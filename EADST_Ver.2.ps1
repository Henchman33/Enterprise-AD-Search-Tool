# ADSearchTool.ps1 - Enterprise GUI for AD searches & exports
# Run on a Domain Controller or workstation with AD module installed.
# Requires PowerShell 5.1+ or 7+ (GUI uses WinForms)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === Config & Helpers ===
$ExportRoot = "C:\Temp\ADSearchTool\Export"
If (!(Test-Path $ExportRoot)) { New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null }

function Ensure-ModuleLoaded {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        return $false
    } else {
        Import-Module $Name -ErrorAction Stop
        return $true
    }
}
$HasAD = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO = Ensure-ModuleLoaded -Name GroupPolicy

function Resolve-SIDToName {
    param([string]$sid)
    try {
        $nt = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $acct = $nt.Translate([System.Security.Principal.NTAccount])
        return $acct.Value
    } catch {
        return $sid
    }
}

function SafeFileName {
    param([string]$name)
    return ($name -replace '[^\w\-\._ ]','_').Trim()
}

# === Export function (includes header/summary) ===
function Export-Results {
    param(
        [Parameter(Mandatory=$true)][array]$Results,
        [Parameter(Mandatory=$true)][string]$Category,
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][string[]]$Formats
    )

    if (!(Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $base = Join-Path $ExportPath ("{0}_{1}" -f (SafeFileName $Category), $timestamp)
    $total = $Results.Count

    $header = @"
====================================================================
 Active Directory Search Export
--------------------------------------------------------------------
 Search Type : $Category
 Filter Used : $Filter
 Export Time : $(Get-Date)
 Total Found : $total
====================================================================
"@.Trim()

    # If no results, create a 'no results' file for each requested format (or a summary)
    if ($total -eq 0) {
        foreach ($fmt in $Formats) {
            switch -Wildcard ($fmt.ToLower()) {
                "csv" {
                    $file = $base + ".csv"
                    "# $header" | Out-File -FilePath $file -Encoding UTF8
                }
                "xml" {
                    $file = $base + ".xml"
                    "<results><summary> $([System.Security.SecurityElement]::Escape($header)) </summary></results>" | Out-File -FilePath $file -Encoding UTF8
                }
                "html" {
                    $file = $base + ".html"
                    "<html><body><pre>$($header)</pre><h3>No results</h3></body></html>" | Out-File -FilePath $file -Encoding UTF8
                }
                "txt" {
                    $file = $base + ".txt"
                    $header | Out-File -FilePath $file -Encoding UTF8
                }
                "excel(.xlsx)" {
                    $file = $base + ".csv"
                    "# $header" | Out-File -FilePath $file -Encoding UTF8
                }
                "pdf" {
                    $file = $base + ".html"
                    "<html><body><pre>$($header)</pre><h3>No results</h3></body></html>" | Out-File -FilePath $file -Encoding UTF8
                }
                "docx" {
                    $file = $base + ".txt"
                    $header | Out-File -FilePath $file -Encoding UTF8
                }
            }
        }
        Write-Host "[!] No results - wrote placeholder files to $ExportPath"
        return
    }

    foreach ($fmt in $Formats) {
        switch -Wildcard ($fmt.ToLower()) {
            "csv" {
                $file = $base + ".csv"
                # Prepend header lines as commented lines starting with '#'
                $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                $csvLines = $Results | ConvertTo-Csv -NoTypeInformation
                $headerLines + $csvLines | Out-File -FilePath $file -Encoding UTF8
                Write-Host "CSV exported: $file"
            }
            "xml" {
                $file = $base + ".xml"
                # Add a comment with header then the CLIXML payload
                $xmlComment = "<!-- " + ($header -replace '--','- -') + " -->`n"
                $Results | Export-Clixml -Path $file
                # Prepend comment (safe small hack: write temp and prepend)
                $tmp = Get-Content -Path $file -Raw
                $xmlComment + $tmp | Out-File -FilePath $file -Encoding UTF8
                Write-Host "XML exported: $file"
            }
            "html" {
                $file = $base + ".html"
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title "AD Search Results - $Category"
                $html | Out-File -FilePath $file -Encoding UTF8
                Write-Host "HTML exported: $file"
            }
            "txt" {
                $file = $base + ".txt"
                $header | Out-File -FilePath $file -Encoding UTF8
                $Results | Out-String | Out-File -FilePath $file -Append -Encoding UTF8
                Write-Host "TXT exported: $file"
            }
            "excel(.xlsx)" {
                $file = $base + ".xlsx"
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    try {
                        $Results | Export-Excel -Path $file -WorksheetName "Results" -AutoSize -Title ("AD Search Results - " + $Category)
                        Write-Host "Excel exported: $file"
                    } catch {
                        Write-Warning "Excel export failed: $_. Falling back to CSV."
                        $csvfile = $base + ".csv"
                        $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                        $csvLines = $Results | ConvertTo-Csv -NoTypeInformation
                        $headerLines + $csvLines | Out-File -FilePath $csvfile -Encoding UTF8
                    }
                } else {
                    Write-Warning "ImportExcel module not present. Saving CSV fallback."
                    $csvfile = $base + ".csv"
                    $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                    $csvLines = $Results | ConvertTo-Csv -NoTypeInformation
                    $headerLines + $csvLines | Out-File -FilePath $csvfile -Encoding UTF8
                }
            }
            "pdf" {
                # Create HTML first and inform user how to convert to PDF automatically if they have wkhtmltopdf
                $fileHtml = $base + ".html"
                $filePdf  = $base + ".pdf"
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title "AD Search Results - $Category"
                $html | Out-File -FilePath $fileHtml -Encoding UTF8
                $wk = (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue).Path
                if ($wk) {
                    & $wk $fileHtml $filePdf
                    Write-Host "PDF exported: $filePdf"
                } else {
                    Write-Warning "wkhtmltopdf not found. Saved HTML: $fileHtml. Use wkhtmltopdf to convert to PDF."
                }
            }
            "docx" {
                # Create HTML and attempt Word COM conversion if Word installed; otherwise save HTML and TXT
                $fileHtml = $base + ".html"
                $fileDocx = $base + ".docx"
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title "AD Search Results - $Category"
                $html | Out-File -FilePath $fileHtml -Encoding UTF8
                try {
                    $word = New-Object -ComObject Word.Application -ErrorAction Stop
                    $doc = $word.Documents.Add($fileHtml)
                    $doc.SaveAs([ref] $fileDocx, [ref] 16)
                    $doc.Close()
                    $word.Quit()
                    Write-Host "DOCX exported: $fileDocx"
                } catch {
                    Write-Warning "Word COM not available. Saved HTML: $fileHtml. Open and Save As DOCX manually."
                }
            }
        }
    }

    Write-Host "[✓] Exports complete. Total results: $total"
}

# === Search functions ===
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

function Search-Users {
    param([string]$filter)
    $props = @("Name","sAMAccountName","distinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","memberOf","userPrincipalName")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADUser -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADUser -Filter { Name -like $f -or sAMAccountName -like $f -or mail -like $f -or userPrincipalName -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    $res | Select-Object @{n='Type';e={'User'}}, Name,sAMAccountName,distinguishedName,Enabled,LockedOut,LastLogonDate,whenCreated,userPrincipalName,@{n='MemberOf';e={$_.memberOf -join '; '}}
}

function Search-Computers {
    param([string]$filter)
    $props = @("Name","OperatingSystem","OperatingSystemVersion","distinguishedName","whenCreated","lastLogonDate")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADComputer -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADComputer -Filter { Name -like $f -or OperatingSystem -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    $res | Select-Object @{n='Type';e={'Computer'}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}}
}

function Search-OUs {
    param([string]$filter)
    $res = Get-ADOrganizationalUnit -Filter { Name -like $filter } -Properties distinguishedName,whenCreated -ErrorAction SilentlyContinue
    $res | Select-Object @{n='Type';e={'OU'}}, Name,distinguishedName,whenCreated
}

function Search-GPOs {
    param([string]$filter)
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        throw "GroupPolicy module not available."
    }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All | Where-Object { $_.DisplayName -like $filter }
    $out = foreach ($g in $gpos) {
        $links = (Get-GPOLink -Guid $g.Id).LinksTo | ForEach-Object { $_.Scope } -join "; "
        [pscustomobject]@{
            Type = "GPO"
            Name = $g.DisplayName
            Id = $g.Id
            Owner = $g.Owner
            CreationTime = $g.CreationTime
            ModificationTime = $g.ModificationTime
            Links = $links
        }
    }
    return $out
}

function Search-Groups {
    param([string]$filter)
    $props = @("Name","distinguishedName","GroupCategory","GroupScope","member")
    $res = Get-ADGroup -Filter { Name -like $filter } -Properties $props -ErrorAction SilentlyContinue
    $res | Select-Object @{n='Type';e={'Group'}}, Name,GroupScope,GroupCategory,distinguishedName,@{n='Members';e={$_.member -join '; '}}
}

function Search-ServiceAccounts {
    param([string]$filter)
    $res = Get-ADUser -Filter { servicePrincipalName -like $filter -or sAMAccountName -like $filter } -Properties servicePrincipalName,description,distinguishedName -ErrorAction SilentlyContinue
    $res | Select-Object @{n='Type';e={'ServiceAccount'}}, Name,sAMAccountName,servicePrincipalName,distinguishedName,description
}

function Search-ServersOrWorkstations {
    param([string]$filter,[switch]$Servers)
    $osFilter = if ($Servers) { "*Server*" } else { "*Windows*" }
    $res = Get-ADComputer -Filter { OperatingSystem -like $osFilter -and Name -like $filter } -Properties OperatingSystem,OperatingSystemVersion,distinguishedName,lastLogonDate -ErrorAction SilentlyContinue
    $res | Select-Object @{n='Type';e={if ($Servers) {'Server'}else{'Workstation'}}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.lastLogonDate}}
}

function Search-LockedOutUsers {
    param([switch]$ResolveOrigin)
    # Get locked out accounts (AD attribute)
    $locked = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue |
        Get-ADUser -Properties LockedOut,LastLogonDate,whenCreated,sAMAccountName,distinguishedName -ErrorAction SilentlyContinue

    $out = @()
    foreach ($u in $locked) {
        $record = [pscustomobject]@{
            Type = "LockedUser"
            Name = $u.Name
            sAMAccountName = $u.sAMAccountName
            DistinguishedName = $u.DistinguishedName
            LockedOut = $u.LockedOut
            LastLogon = $u.LastLogonDate
        }

        if ($ResolveOrigin) {
            $dcs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
            $origin = $null
            foreach ($dc in $dcs) {
                try {
                    $query = @"
<QueryList>
  <Query Id='0' Path='Security'>
    <Select Path='Security'>
      *[System[(EventID=4740)]] and *[EventData[Data and (Data='$($u.sAMAccountName)')]]
    </Select>
  </Query>
</QueryList>
"@
                    $events = Get-WinEvent -ComputerName $dc.HostName -FilterXml $query -MaxEvents 1 -ErrorAction SilentlyContinue
                    if ($events -and $events.Count -gt 0) {
                        $ev = $events[0]
                        $data = [xml]$ev.ToXml()
                        $td = $data.Event.EventData.Data
                        $caller = ($td | Where-Object { $_.Name -eq "CallerComputerName" }).'#text'
                        $origin = @{
                            DomainController = $dc.HostName
                            CallerComputer   = $caller
                            EventTime        = $ev.TimeCreated
                        }
                        break
                    }
                } catch { }
            }
            if ($origin) {
                $originValue = ($origin | Out-String).Trim()
            } else {
                $originValue = "Origin not found or insufficient permissions"
            }
            $record | Add-Member -MemberType NoteProperty -Name "LockoutOrigin" -Value $originValue -Force
        }

        $out += $record
    }

    return $out
}

function Search-Subnets {
    try {
        $configNaming = (Get-ADRootDSE).configurationNamingContext
        $base = "CN=Subnets,CN=Sites,$configNaming"
        $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -ErrorAction SilentlyContinue
        return $subnets | Select-Object @{n='Type';e={'ADSubnet'}}, name,@{n='Location';e={$_.location}},@{n='DistinguishedName';e={$_.DistinguishedName}}
    } catch {
        return @()
    }
}

function Search-GPOFirewallRules {
    param([string]$filter)
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) { throw "GroupPolicy module unavailable." }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All
    $matches = @()
    foreach ($g in $gpos) {
        $xml = Get-GPOReport -Guid $g.Id -ReportType Xml
        [xml]$gxml = $xml
        $policies = @()
        if ($gxml.GPO.Computer.ExtensionData.Extension.Policy) { $policies = $gxml.GPO.Computer.ExtensionData.Extension.Policy }
        foreach ($p in $policies) {
            if ($p.Name -like "*Firewall*" -or $p.Setting -like "*Firewall*" -or $p.Name -like $filter -or $p.Setting -like $filter) {
                $matches += [pscustomobject]@{
                    Type     = "GPOFirewall"
                    GPOName  = $g.DisplayName
                    Policy   = $p.Name
                    Setting  = $p.Setting
                    Links    = ((Get-GPOLink -Guid $g.Id).LinksTo | ForEach-Object { $_.Scope }) -join "; "
                }
            }
        }
    }
    return $matches
}

# === UI Event Handlers ===
$btnOpenExport.Add_Click({
    $path = $txtExportPath.Text
    if (!(Test-Path $path)) { New-Item -Path $path -ItemType Directory -Force | Out-Null }
    Start-Process -FilePath $path
})

$btnSearch.Add_Click({
    $lblStatus.Text = "Running search..."
    $Form.Refresh()

    $sel = $comboType.SelectedItem
    $filter = $txtSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filter)) { $filter = "*" }
    $results = @()

    try {
        switch ($sel) {
            "User" { $results = Search-Users -filter $filter }
            "Computer" { $results = Search-Computers -filter $filter }
            "OU" { $results = Search-OUs -filter $filter }
            "GPO" {
                if (-not $HasGPO) { [System.Windows.Forms.MessageBox]::Show("GroupPolicy module not available on this host.","Missing Module","OK","Warning"); $results = @() }
                else { $results = Search-GPOs -filter $filter }
            }
            "Security Group" { $results = Search-Groups -filter $filter }
            "Service Accounts" { $results = Search-ServiceAccounts -filter $filter }
            "Servers (by OS or group)" { $results = Search-ServersOrWorkstations -filter $filter -Servers }
            "Workstations (by OS or group)" { $results = Search-ServersOrWorkstations -filter $filter }
            "Locked-out Users (basic)" { $results = Search-LockedOutUsers -ResolveOrigin:$false }
            "Locked-out Users (with origin/event lookup)" {
                if (-not $chkEventLookup.Checked) {
                    if ([System.Windows.Forms.MessageBox]::Show("You selected origin lookup but didn't enable the option. Enable it now?","Event lookup","YesNo","Question") -eq "Yes") {
                        $chkEventLookup.Checked = $true
                    }
                }
                $results = Search-LockedOutUsers -ResolveOrigin:$chkEventLookup.Checked
            }
            "Subnets (AD Sites & Services)" { $results = Search-Subnets }
            "Firewall (GPO firewall rules)" {
                if (-not $HasGPO) { [System.Windows.Forms.MessageBox]::Show("GroupPolicy module not available: cannot search firewall rules.","Missing Module","OK","Warning"); $results = @() }
                else { $results = Search-GPOFirewallRules -filter $filter }
            }
            "All: Users+Computers" {
                $results = @()
                $results += (Search-Users -filter $filter)
                $results += (Search-Computers -filter $filter)
            }
            Default { $results = @() }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Search error: $($_.Exception.Message)","Error","OK","Error")
        $lblStatus.Text = "Error during search."
        return
    }

    # show results
    if ($results -and $results.Count -gt 0) {
        $grid.DataSource = $results
        $lblStatus.Text = "Found $($results.Count) item(s)."
        $grid.AutoSizeColumnsMode = "DisplayedCells"
    } else {
        $grid.DataSource = $null
        $lblStatus.Text = "No results found."
    }

    # Prepare export formats list
    $formats = @()
    for ($i=0; $i -lt $chkExport.Items.Count; $i++) {
        if ($chkExport.GetItemChecked($i)) { $formats += $chkExport.Items[$i] }
    }

    if ($formats.Count -gt 0) {
        # Map checked item labels to normalized tokens used by Export-Results
        $normFormats = $formats | ForEach-Object {
            switch ($_){
                "CSV" { "csv" }
                "XML" { "xml" }
                "HTML" { "html" }
                "TXT" { "txt" }
                "Excel(.xlsx)" { "excel(.xlsx)" }
                "PDF" { "pdf" }
                "DOCX" { "docx" }
                default { $_.ToLower() }
            }
        }

        Export-Results -Results ($results | ForEach-Object { $_ }) `
            -Category $sel -Filter $filter -ExportPath $txtExportPath.Text -Formats $normFormats

        $lblStatus.Text += " Exported to $($txtExportPath.Text) in " + ($formats -join ", ")
    } else {
        $lblStatus.Text += " (No export formats selected.)"
    }
})

# === Show form ===
[void]$Form.ShowDialog()
