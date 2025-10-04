# ADSearchTool.ps1
# PowerShell WinForms GUI to search Active Directory and export results

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Ensure Export folder exists
$ExportRoot = "C:\Temp\ADSearchTool\Export"
If (!(Test-Path $ExportRoot)) { New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null }

# --- Load AD & GPO modules gracefully
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

# --- UI building
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "AD Search Tool"
$Form.Size = New-Object System.Drawing.Size(1100,720)
$Form.StartPosition = "CenterScreen"

# Search type dropdown
$lblType = New-Object System.Windows.Forms.Label
$lblType.Location = New-Object System.Drawing.Point(12,14)
$lblType.Size = New-Object System.Drawing.Size(120,20)
$lblType.Text = "Search Category:"
$Form.Controls.Add($lblType)

$comboType = New-Object System.Windows.Forms.ComboBox
$comboType.Location = New-Object System.Drawing.Point(140,10)
$comboType.Size = New-Object System.Drawing.Size(320,24)
$comboType.DropDownStyle = "DropDownList"
$comboItems = @(
    "User",
    "Computer",
    "OU",
    "GPO",
    "Security Group",
    "Service Accounts",
    "Servers (by OS or group)",
    "Workstations (by OS or group)",
    "Locked-out Users (basic)",
    "Locked-out Users (with origin/event lookup)",
    "Subnets (AD Sites & Services)",
    "Firewall (GPO firewall rules)",
    "All: Users+Computers"
)
$comboType.Items.AddRange($comboItems)
$comboType.SelectedIndex = 0
$Form.Controls.Add($comboType)

# Search filter
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Location = New-Object System.Drawing.Point(12,50)
$lblSearch.Size = New-Object System.Drawing.Size(120,20)
$lblSearch.Text = "Search Filter:"
$Form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(140,46)
$txtSearch.Size = New-Object System.Drawing.Size(640,24)
$txtSearch.Text = "*"
$Form.Controls.Add($txtSearch)

# --- Options group
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Location = New-Object System.Drawing.Point(12,80)
$grpOptions.Size = New-Object System.Drawing.Size(1060,60)
$grpOptions.Text = "Options"
$Form.Controls.Add($grpOptions)

$chkResolveSIDs = New-Object System.Windows.Forms.CheckBox
$chkResolveSIDs.Location = New-Object System.Drawing.Point(12,22)
$chkResolveSIDs.Size = New-Object System.Drawing.Size(220,20)
$chkResolveSIDs.Text = "Resolve SIDs to account names"
$chkResolveSIDs.Checked = $true
$grpOptions.Controls.Add($chkResolveSIDs)

$chkIncludeDisabled = New-Object System.Windows.Forms.CheckBox
$chkIncludeDisabled.Location = New-Object System.Drawing.Point(250,22)
$chkIncludeDisabled.Size = New-Object System.Drawing.Size(200,20)
$chkIncludeDisabled.Text = "Include disabled accounts"
$chkIncludeDisabled.Checked = $true
$grpOptions.Controls.Add($chkIncludeDisabled)

$chkEventLookup = New-Object System.Windows.Forms.CheckBox
$chkEventLookup.Location = New-Object System.Drawing.Point(470,22)
$chkEventLookup.Size = New-Object System.Drawing.Size(300,20)
$chkEventLookup.Text = "For lockouts: attempt event-log lookup (slow)"
$chkEventLookup.Checked = $false
$grpOptions.Controls.Add($chkEventLookup)

# --- Buttons
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Location = New-Object System.Drawing.Point(800,44)
$btnSearch.Size = New-Object System.Drawing.Size(120,30)
$btnSearch.Text = "Run Search"
$Form.Controls.Add($btnSearch)

$lblExportTo = New-Object System.Windows.Forms.Label
$lblExportTo.Location = New-Object System.Drawing.Point(12,152)
$lblExportTo.Size = New-Object System.Drawing.Size(200,20)
$lblExportTo.Text = "Export format(s):"
$Form.Controls.Add($lblExportTo)

# Export format list
$chkExport = New-Object System.Windows.Forms.CheckedListBox
$chkExport.Location = New-Object System.Drawing.Point(12,178)
$chkExport.Size = New-Object System.Drawing.Size(460,80)
$chkExport.CheckOnClick = $true
$exportOptions = @("CSV","XML","HTML","TXT","Excel(.xlsx)","PDF","DOCX")
$chkExport.Items.AddRange($exportOptions)
$chkExport.SetItemChecked(0,$true)
$chkExport.SetItemChecked(2,$true)
$Form.Controls.Add($chkExport)

# Export path
$lblExportPath = New-Object System.Windows.Forms.Label
$lblExportPath.Location = New-Object System.Drawing.Point(490,178)
$lblExportPath.Size = New-Object System.Drawing.Size(120,20)
$lblExportPath.Text = "Export Folder:"
$Form.Controls.Add($lblExportPath)

$txtExportPath = New-Object System.Windows.Forms.TextBox
$txtExportPath.Location = New-Object System.Drawing.Point(490,200)
$txtExportPath.Size = New-Object System.Drawing.Size(440,24)
$txtExportPath.Text = $ExportRoot
$Form.Controls.Add($txtExportPath)

$btnOpenExport = New-Object System.Windows.Forms.Button
$btnOpenExport.Location = New-Object System.Drawing.Point(940,200)
$btnOpenExport.Size = New-Object System.Drawing.Size(80,24)
$btnOpenExport.Text = "Open"
$Form.Controls.Add($btnOpenExport)

# Results grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12,270)
$grid.Size = New-Object System.Drawing.Size(1060,380)
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = "AllCells"
$Form.Controls.Add($grid)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12,660)
$lblStatus.Size = New-Object System.Drawing.Size(1060,20)
$lblStatus.Text = "Ready."
$Form.Controls.Add($lblStatus)

# --- Core Search Functions (Users, Computers, etc.) ---
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# [Keeping your existing search functions exactly as-is for Users, Computers, OUs, GPOs, etc.]

# --- Fixed Locked-Out Users Function ---
function Search-LockedOutUsers {
    param([switch]$ResolveOrigin)

    $locked = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue |
        Get-ADUser -Properties LockedOut,LastLogonDate,whenCreated,sAMAccountName,distinguishedName

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
                            CallerComputer = $caller
                            EventTime = $ev.TimeCreated
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

            Add-Member -InputObject $record -NotePropertyName "LockoutOrigin" -NotePropertyValue $originValue
        }

        $out += $record
    }

    return $out
}

# --- [Keep Export-Results and event handlers from your original script here, unchanged] ---

# Show GUI
[void]$Form.ShowDialog()
