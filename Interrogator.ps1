# Interrogator_v3.ps1
# Pulls auditing information
#   This is a major rewrite to functionalize and export to a single XLSX rather than individual CSVs
# Author: Stephen Kleine [kleines2015@gmail.com]
# Version 2.0 - 20201216
# Revision  

# USAGE
# .\Interrogator.ps1 [-ShowMagic] [-InspectServiceAccounts]

# KNOWN BUGS
#   Creating new worksheets within a function doesn't work.

# HOUSEKEEPING
# The below fixes a code validation issue for powershell in VSCode https://github.com/PowerShell/PSScriptAnalyzer/issues/827
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseDeclaredVarsMoreThanAssignments", "")]
Param()

# Parameters
param (
    [switch]$ShowMagic = $false,
    [switch]$InspectServiceAccounts = $false
)

# Import all the needed modules

Import-Module -Name GroupPolicy, ActiveDirectory -ErrorAction stop

# Global variables

$Root = [ADSI]"LDAP://RootDSE" # Used for multi-domain environments
$RootDN = $Root.rootDomainNamingContext # pulls the root domain's DN, needed for polling ADSI directly
$ConfigurationSearchBase = "cn=configuration,$RootDN"
$DomainControllerADWS = (get-addomaincontroller -discover -service ADWS).Name
$UserName = $env:USERNAME
$UserTempDir = $env:TEMP
$StartTimeStamp = get-date -f FileDateTime
$AnalysisTempDir = "$UserTempDir\AnalysisReport_$StartTimeStamp" #Put a subdirectory into the TEMP folder
$90Days = (get-date).ticks - 504988992000000000 #90 days ago, needed for stale users report
$ValidServiceAccounts = @('localSystem', 'NT AUTHORITY\NetworkService', 'NT AUTHORITY\LocalService') # used for service detections
[int]$WorksheetIndex = 1  

function BuildHeaders($WorksheetVariableName, $Column, $Header) {
    # Writes your headers
    $WorksheetVariableName.Cells.Item(1, $Column) = $Header
}
function RenameWorksheet($WorksheetVariableName, $TabName) { $WorksheetVariableName.Name = $TabName }

function FillNewRow ($Row, $Column, $Value) { $excel.cells.item($Row, $Column) = "$Value" }

Function CleanupAndClose {
    $excel.workbooks.close()
    $excel.Quit()
    $excel = $Workbook = $uregwksht = $null
    remove-item $AnalysisTempDir -Recurse -Force
    [System.GC]::Collect()
}

Function ChangeWorksheet ($Name) {
    $Worksheet = $Workbook.Worksheets.item("$Name")
    $Worksheet.Activate()
}

# Prepartion for Excel build
mkdir $AnalysisTempDir -ea Stop -wa stop | out-null # Path to temp directory and create folder

# Build Workbook
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$excel = New-Object -ComObject excel.application
$Workbook = $excel.Workbooks.Add()
if ($ShowMagic) {
    $excel.visible = $True # For debug only, will slow the processing markedly and cause errors
}

# Pull all GPOs and export as XML and HTML

Write-host "Dumping GPOs..."
# This builds new worksheets - functionizing doesn't work?
$uregwksht = $workbook.Worksheets.Item($WorksheetIndex)
$WorksheetIndex++
RenameWorksheet $uregwksht 'GPOs' $WorksheetIndex
BuildHeaders $uregwksht 1 'DisplayName'
BuildHeaders $uregwksht 2 'GUID' 
BuildHeaders $uregwksht 3 'Description' 
BuildHeaders $uregwksht 4 'ComputerChanges'
BuildHeaders $uregwksht 5 'ComputerActive'
BuildHeaders $uregwksht 6 'User Changes'
BuildHeaders $uregwksht 7 'User Active'
BuildHeaders $uregwksht 8 'LinksTo'
$i = 2

# Mainline
$AllGPOs = get-gpo -All 
Foreach ($Policy in $AllGPOs) {
    $ID = $Policy.id
    Get-GPOReport -Guid $Policy.id -ReportType XML | out-file -filepath $AnalysisTempDir\$ID.xml -Encoding utf8 # Did it this way because special characters in a GPO's name cause problems with writing to disk
    [XML]$GPOFile = Get-Content "$AnalysisTempDir\$ID.xml"
    foreach ($item in $GPOfile.GPO) { 
        if ($null -eq $item.computer.ExtensionData.IsEmpty) { $ComputerChanges = $false } else { $ComputerChanges = $true }
        if ($null -eq $item.user.ExtensionData.IsEmpty) { $UserChanges = $false } else { $UserChanges = $true }
        $ComputerStatus = $item.computer.Enabled
        $UserStatus = $item.user.Enabled
        $LinksFound = $item.LinksTo
        FillNewRow $i 1 $Policy.DisplayName
        FillNewRow $i 2 $Policy.id.GUID
        FillNewRow $i 3 $Policy.Description
        FillNewRow $i 4 $ComputerChanges
        FillNewRow $i 5 $ComputerStatus
        FillNewRow $i 6 $UserChanges
        FillNewRow $i 7 $UserStatus
        if ($LinksFound) { 
            $j = 8
            ForEach ($Link in $LinksFound) {
                FillNewRow $i $j $Link.SOMPath
                $j++
            }
        }
        $i++
    }
}

# Create and config sheets 
# Disabled user identities
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'DisabledUsers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Description' 
BuildHeaders $uregwksht 3 'WhenChanged'
BuildHeaders $uregwksht 4 'PwdLastSet'
$DisabledIndex = 2

# No Password Expiry User Identities
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'No Password Expiry' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Description' 
BuildHeaders $uregwksht 3 'WhenChanged'
BuildHeaders $uregwksht 4 'PwdLastSet'
BuildHeaders $uregwksht 5 'IsEnabled'
$UnexpiringIndex = 2

# Over ninety days since user identity logged on
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Aged Users' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Description' 
BuildHeaders $uregwksht 3 'LastLogonTimestamp'
BuildHeaders $uregwksht 4 'IsEnabled'
$AgedUsersIndex = 2

# Password not required for user identity
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'No Password Required' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'IsEnabled'
BuildHeaders $uregwksht 3 'Description' 
$NoPasswordIndex = 2

# Never used user identity
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Never Used User' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'IsEnabled'
BuildHeaders $uregwksht 3 'Description' 
$NeverLoggedOnUserIndex = 2

# Stale Password User Identities
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Stale Password' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Description' 
BuildHeaders $uregwksht 3 'PwdLastSet'
BuildHeaders $uregwksht 4 'IsEnabled'
$StaleUserPasswordIndex = 2

# Domain Admins
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Domain Admins' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'IsEnabled'
$DomainAdminsIndex = 2

# Enterprise Admins
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Enterprise Admins' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'IsEnabled'
$EnterpriseAdminsIndex = 2

# Schema Admins
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Schema Admins' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'IsEnabled'
$SchemaAdminsIndex = 2

Write-Host "Enumerating User identity issues..."
$AllUserAccounts = get-aduser -server $DomainControllerADWS -f * -properties Name, Description, PasswordNeverExpires, PasswordNotRequired, Lastlogontimestamp, Enabled, PwdLastSet, WhenChanged, MemberOf
ForEach ($UserName in $AllUserAccounts) {
    If ($UserName.Enabled -eq $false) {
        ChangeWorksheet "DisabledUsers"
        FillNewRow $DisabledIndex 1 $UserName.Name
        FillNewRow $DisabledIndex 2 $UserName.Description
        if ($null -ne $UserName.WhenChanged) { FillNewRow $DisabledIndex 3 $UserName.WhenChanged }
        if ($null -ne $UserName.PwdLastSet) { FillNewRow $DisabledIndex 4 ([datetime]::FromFileTimeutc($UserName.pwdlastset).ToString('yyyy-MM-dd')) }
        $DisabledIndex++
    }
    if ($Username.PasswordNeverExpires -eq $True) {
        ChangeWorksheet "No Password Expiry"
        FillNewRow $UnexpiringIndex 1 $UserName.Name
        FillNewRow $UnexpiringIndex 2 $UserName.Description
        if ($null -ne $UserName.WhenChanged) { FillNewRow $UnexpiringIndex 3 $UserName.WhenChanged }
        if ($null -ne $UserName.PwdLastSet) { FillNewRow $UnexpiringIndex 4 ([datetime]::FromFileTimeutc($UserName.pwdlastset).ToString('yyyy-MM-dd')) }
        FillNewRow $UnexpiringIndex 5 $UserName.Enabled
        $UnexpiringIndex++
    }
    if ($Username.LastLogontimestamp -lt $90Days) {
        ChangeWorksheet "Aged Users"
        FillNewRow $AgedUsersIndex 1 $UserName.Name
        FillNewRow $AgedUsersIndex 2 $UserName.Description
        FillNewRow $AgedUsersIndex 3 $UserName.LastLogonTimestamp
        if ($null -ne $UserName.LastLogonTimestamp) { FillNewRow $AgedUsersIndex 3 ([datetime]::FromFileTimeutc($UserName.Lastlogontimestamp).ToString('yyyy-MM-dd')) }
        FillNewRow $AgedUsersIndex 4 $UserName.Enabled
        $AgedUsersIndex++
    }
    if ($Username.PasswordNotRequired) {
        ChangeWorksheet "No Password Required"
        FillNewRow $NoPasswordIndex 1 $UserName.Name
        FillNewRow $NoPasswordIndex 1 $UserName.Enabled
        FillNewRow $NoPasswordIndex 3 $UserName.Description
        $NoPasswordIndex++
    }
    if (($null -eq $Username.lastlogontimestamp) -and ($Username.enabled -eq $true)) {
        ChangeWorksheet "Never Used User"
        FillNewRow $NeverLoggedOnUserIndex 1 $UserName.Name
        FillNewRow $NeverLoggedOnUserIndex 2 $UserName.Enabled
        FillNewRow $NeverLoggedOnUserIndex 3 $UserName.Description
        $NeverLoggedOnUserIndex++
    }
    if ($Username.pwdlastset -lt $90Days) {
        ChangeWorksheet "Stale Password"
        FillNewRow $StaleUserPasswordIndex 1 $UserName.Name
        FillNewRow $StaleUserPasswordIndex 2 $UserName.Description
        if ($null -ne $UserName.PwdLastSet) { FillNewRow $UnexpiringIndex 3 ([datetime]::FromFileTimeutc($UserName.pwdlastset).ToString('yyyy-MM-dd')) }
        FillNewRow $StaleUserPasswordIndex 4 $UserName.Enabled
        $StaleUserPasswordIndex++
    }
    if ($UserName.Memberof -like "*Domain Admins*") {
        ChangeWorksheet "Domain Admins"
        FillNewRow $DomainAdminsIndex 1 $UserName.Name
        FillNewRow $DomainAdminsIndex 2 $UserName.Enabled
        $DomainAdminsIndex++
    }
    if ($UserName.Memberof -like "*Enterprise Admins*") {
        ChangeWorksheet "Enterprise Admins"
        FillNewRow $EnterpriseAdminsIndex 1 $UserName.Name
        FillNewRow $EnterpriseAdminsIndex 2 $UserName.Enabled
        $EnterpriseAdminsIndex++
    } 
    if ($UserName.Memberof -like "*Schema Admins*") {
        ChangeWorksheet "Schema Admins"
        FillNewRow $SchemaAdminsIndex 1 $UserName.Name
        FillNewRow $SchemaAdminsIndex 2 $UserName.Enabled
        $SchemaAdminsIndex++
    }
}

# Now onto groups
Write-host "Enumerating Group issues..."
$AllGroups = get-adgroup -f * -Properties Name, GroupCategory, GroupScope, Description, member, mail, memberOf

# Build Group pages
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Mail-enabled groups' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Group Category' 
BuildHeaders $uregwksht 3 'Group Scope'
BuildHeaders $uregwksht 4 'Description'
BuildHeaders $uregwksht 5 'Mail Address'
$MailEnabledGroupIndex = 2

$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'No Group Members' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Group Category' 
BuildHeaders $uregwksht 3 'Group Scope'
BuildHeaders $uregwksht 4 'Description'
$NoGroupMembersIndex = 2

$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Nested groups' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Group Category' 
BuildHeaders $uregwksht 3 'Group Scope'
BuildHeaders $uregwksht 4 'Description'
BuildHeaders $uregwksht 5 'Description'
BuildHeaders $uregwksht 6 'Member Of'
$NestedGroupIndex = 2

Foreach ($Group in $AllGroups) {
    if ($Group.mail) {
        ChangeWorksheet "Mail-enabled groups"
        FillNewRow $MailEnabledGroupIndex 1 $Group.Name
        FillNewRow $MailEnabledGroupIndex 2 $Group.GroupCategory
        FillNewRow $MailEnabledGroupIndex 3 $Group.GroupScope
        FillNewRow $MailEnabledGroupIndex 4 $Group.Description
        FillNewRow $MailEnabledGroupIndex 5 $Group.mail
        $MailEnabledGroupIndex++
    }
    if ($Group.member.count -eq "0") {
        ChangeWorksheet "No Group Members"
        FillNewRow $NoGroupMembersIndex 1 $Group.Name
        FillNewRow $NoGroupMembersIndex 2 $Group.GroupCategory
        FillNewRow $NoGroupMembersIndex 3 $Group.GroupScope
        FillNewRow $NoGroupMembersIndex 4 $Group.Description
        $NoGroupMembersIndex++
    }
    If ($Group.MemberOf) {  
        ChangeWorksheet "Nested Groups"
        FillNewRow $NestedGroupIndex 1 $Group.name
        FillNewRow $NestedGroupIndex 2 $Group.GroupCategory
        FillNewRow $NestedGroupIndex 3 $Group.GroupScope    
        FillNewRow $NestedGroupIndex 4 $Group.Description
        FillNewRow $NestedGroupIndex 5 $Group.Description
        $j = 6
        ForEach ($Subgroup in $Group.MemberOf) {
            FillNewRow $NestedGroupIndex $j $Subgroup
            $j++
        }
        $NestedGroupIndex++
    }
}

# Infrastructure
# Build worksheets
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'DHCP Servers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
$DhcpServersIndex = 2

$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'PKI Servers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
$PkiServersIndex = 2

# Grab all Objects
$AllObjects = Get-ADObject -SearchBase $ConfigurationSearchBase -f *

Foreach ($ObjectFound in $AllObjects) {
    if (($ObjectFound.objectclass -eq "dHCPClass") -and ($ObjectFound.Name -ne "DhcpRoot")) {
        ChangeWorksheet "DHCP Servers"
        FillNewRow $DhcpServersIndex 1 $ObjectFound.Name
        $DhcpServersIndex++
    }
    if ($ObjectFound.DistinguishedName -ilike "*CN=Certification Authorities*") {
        ChangeWorksheet "PKI Servers"
        FillNewRow $PkiServersIndex 1 $ObjectFound.Name
        $PkiServersIndex++
    }
}

# Computer Systems
# Servers listing
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Windows Servers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'OperatingSystemServicePack' 
BuildHeaders $uregwksht 4 'OperatingSystemVersion' 
$WindowsServersIndex = 2

# Windows workstations listing
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Windows Workstations' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'OperatingSystemServicePack' 
BuildHeaders $uregwksht 4 'OperatingSystemVersion' 
$WindowsWorkstationsIndex = 2

# Non- Microsoft listing
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Non-Microsoft OS' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'OperatingSystemServicePack' 
BuildHeaders $uregwksht 4 'OperatingSystemVersion' 
$NonMicrosoftOsIndex = 2

# Disabled systems
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Disabled Computers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'OperatingSystemServicePack' 
BuildHeaders $uregwksht 4 'OperatingSystemVersion' 
BuildHeaders $uregwksht 5 'LastLogon'
$DisabledComputersIndex = 2

# Stale systems 
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Stale Computers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'OperatingSystemServicePack' 
BuildHeaders $uregwksht 4 'OperatingSystemVersion' 
BuildHeaders $uregwksht 5 'LastLogon'
$StaleComputersIndex = 2

# Poll and categorize
$AllSystems = Get-ADComputer -f * -properties Name, OperatingSystem, LastLogon, WhenCreated, OperatingSystemServicePack, OperatingSystemVersion, lastlogontimestamp
foreach ($SystemFound in $AllSystems) {
    if ($SystemFound.OperatingSystem -inotlike "*Windows*") {
        ChangeWorksheet "Non-Microsoft OS"
        FillNewRow $NonMicrosoftOsIndex 1 $SystemFound.Name
        FillNewRow $NonMicrosoftOsIndex 2 $SystemFound.OperatingSystem
        FillNewRow $NonMicrosoftOsIndex 3 $SystemFound.OperatingSystemServicePack
        FillNewRow $NonMicrosoftOsIndex 4 $SystemFound.OperatingSystemVersion
        $NonMicrosoftOsIndex++
    }
    elseif (($SystemFound.OperatingSystem -ilike "*Server*") -and ($SystemFound.OperatingSystem -ilike "*Windows*")) {
        ChangeWorksheet "Windows Servers"
        FillNewRow $WindowsServersIndex 1 $SystemFound.Name
        FillNewRow $WindowsServersIndex 2 $SystemFound.OperatingSystem
        FillNewRow $WindowsServersIndex 3 $SystemFound.OperatingSystemServicePack
        FillNewRow $WindowsServersIndex 4 $SystemFound.OperatingSystemVersion
        $WindowsServersIndex++
    }
    else {
        ChangeWorksheet "Windows Workstations"
        FillNewRow $WindowsWorkstationsIndex 1 $SystemFound.Name
        FillNewRow $WindowsWorkstationsIndex 2 $SystemFound.OperatingSystem
        FillNewRow $WindowsWorkstationsIndex 3 $SystemFound.OperatingSystemServicePack
        FillNewRow $WindowsWorkstationsIndex 4 $SystemFound.OperatingSystemVersion
        $WindowsWorkstationsIndex++
    }
    if ($SystemFound.enabled -ne $True) {
        ChangeWorksheet "Disabled Computers"
        FillNewRow $DisabledComputersIndex 1 $SystemFound.Name
        FillNewRow $DisabledComputersIndex 2 $SystemFound.OperatingSystem
        FillNewRow $DisabledComputersIndex 3 $SystemFound.OperatingSystemServicePack
        FillNewRow $DisabledComputersIndex 4 $SystemFound.OperatingSystemVersion
        if ($null -ne $SystemFound.LastLogonTimestamp) { FillNewRow $DisabledComputersIndex 5 ([datetime]::FromFileTimeutc($SystemFound.Lastlogontimestamp).ToString('yyyy-MM-dd')) }
        $DisabledComputersIndex++
    }
    if ($SystemFound.LastLogonTimestamp -lt $90Days) {
        ChangeWorksheet "Stale Computers"
        FillNewRow $StaleComputersIndex 1 $SystemFound.Name
        FillNewRow $StaleComputersIndex 2 $SystemFound.OperatingSystem
        FillNewRow $StaleComputersIndex 3 $SystemFound.OperatingSystemServicePack
        FillNewRow $StaleComputersIndex 4 $SystemFound.OperatingSystemVersion
        if ($null -ne $SystemFound.LastLogonTimestamp) { FillNewRow $StaleComputersIndex 5 ([datetime]::FromFileTimeutc($SystemFound.Lastlogontimestamp).ToString('yyyy-MM-dd')) }     
        $StaleComputersIndex++
    }

}

# Domain Controllers - current domain
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Domain Controllers' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'OperatingSystem'
BuildHeaders $uregwksht 3 'IPv4Address' 
BuildHeaders $uregwksht 4 'Enabled' 
$DomainControllersIndex = 2

$DomainControllers = Get-ADDomainController -filter * 
ChangeWorksheet "Domain Controllers"
foreach ($DomainController in $DomainControllers) {
    FillNewRow $DomainControllersIndex 1 $DomainController.Name
    FillNewRow $DomainControllersIndex 2 $DomainController.OperatingSystem
    FillNewRow $DomainControllersIndex 3 $DomainController.IPv4Address
    FillNewRow $DomainControllersIndex 4 $DomainController.Enabled
    $DomainControllersIndex++
}

# Trusts 
$uregwksht = $workbook.Worksheets.add()
$WorksheetIndex++
RenameWorksheet $uregwksht 'Trusts' $WorksheetIndex
BuildHeaders $uregwksht 1 'Name'
BuildHeaders $uregwksht 2 'Target'
BuildHeaders $uregwksht 3 'Direction'
BuildHeaders $uregwksht 4 'Type'
BuildHeaders $uregwksht 5 'Transitive' 
$DomainTrustsIndex = 2

# Poll and categorize
$DomainTrusts = Get-ADtrust -filter * 
ChangeWorksheet "Trusts"
foreach ($Trust in $DomainTrusts) {
    if ($Trust.direction -ne "incoming") {
        ChangeWorksheet "Trusts"
        FillNewRow $DomainTrustsIndex 1 $Trust.Name
        FillNewRow $DomainTrustsIndex 2 $Trust.Target
        FillNewRow $DomainTrustsIndex 3 $Trust.Direction
        if ($trust.ForestTransitive -eq $true) {
            FillNewRow $DomainTrustsIndex 4 "Forest"
        }
        else { FillNewRow $DomainTrustsIndex 4 "External" }
        FillNewRow $DomainTrustsIndex 5 $Trust.ForestTransitive
    }
    $DomainTrustsIndex++
}
# Make visible, user saves, and clean up & close
$excel.visible
CleanupAndClose


# Inspect online system services for default accounts types
if ($InspectServiceAccounts) {
    Write-Host "Analyzing all systems in AD for service account usage, please be patient..."

    $NonLocalServices = @()
    $NonResponsiveSystems = @()

    $AllComputers = get-adcomputer -filter *
    $ValidServiceAccounts = @('localSystem', 'NT AUTHORITY\NetworkService', 'NT AUTHORITY\LocalService')
    Foreach ($system in $AllComputers) {
        try {
            $Services = Get-WmiObject -Class Win32_Service -ComputerName $system.name -EA Stop
            Foreach ($Servicename in $Services) {
                $count = 0
                Foreach ($ValidName in $ValidServiceAccounts) {
                    if ($ServiceName.Startname -ine $ValidName) { $count++ }
                    if ($count -eq 3) {
                        $Device = $system.name, $Servicename.DisplayName, $Servicename.StartName, $Servicename.State -join ','
                        $NonLocalServices += $Device
                    }
                }
            }            
        }
        catch { $NonResponsiveSystems += $system.Name }
    }
        
    $NonLocalServices | out-file $AnalysisTempDir\NonLocalServices.csv
    $NonResponsiveSystems | out-file $AnalysisTempDir\NonResponsiveLocalServiceSystems.csv


    # Finally, open the folder created waaaaay back in the beginning
    invoke-item $AnalysisTempDir
}
