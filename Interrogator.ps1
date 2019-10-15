# Interrogator_v2.ps1
# Pulls auditing information for clients; requires an account that can read all GPOs and a system on their domain
# Author: Stephen Kleine [kleines2015@gmail.com]
# Version 1.0 - 20180525
# Revision 20191011 - Major rework to move into functions

# KNOWN BUGS
#    On a Windows 2012 R2 DC in a Windows 2003 Functional level domain and forest, GPO analysis does not work


#Import all the needed modules

Import-Module -Name GroupPolicy, ActiveDirectory -ErrorAction stop

# Global variables

$Root = [ADSI]"LDAP://RootDSE" # Used for multi-domain environments
$RootDN = $Root.rootDomainNamingContext # pulls the root domain's DN, needed for polling ADSI directly
$UserLogonServer=$env:LOGONSERVER #this *should* return a DC in the local site, no guarantees though
$DomainController = $UserLogonServer.trimstart('\\')
$UserNetBIOSDomain = $env:USERDOMAIN
$DomainControllerADWS = (get-addomaincontroller -discover -service ADWS).Name
$userFQDNdomain = (get-addomain).dnsroot
$UserName = $env:USERNAME
$UserTempDir = $env:TEMP
$StartTimeStamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." } # ISO UTC datetime
$AnalysisTempDir = "$UserTempDir\AnalysisReport_$StartTimeStamp" #Put a subdirectory into the TEMP folder
$90Days = (get-date).ticks - 504988992000000000 #90 days ago, needed for stale users report
$ValidServiceAccounts = @('localSystem','NT AUTHORITY\NetworkService','NT AUTHORITY\LocalService') # used for service detections

New-Item -ItemType directory -Path $AnalysisTempDir

#Pull all GPOs and export as XML and HTML

Write-host "Dumping all GPOs to XML..."

Get-GPOReport -All -Domain $userFQDNdomain -server $DomainController -ReportType xml -path $AnalysisTempDir\GPOReport.xml

Write-host "Dumping all GPOs to HTML..."

Get-GPOReport -All -Domain $userFQDNdomain -server $DomainController -ReportType html -path $AnalysisTempDir\GPOReport.html

# Functions for GPO reporting

function IsNotLinked($xmldata){ 
    If ($xmldata.GPO.LinksTo -eq $null) {Return $true}
    Return $false 
} 
 
function NoUserChanges($xmldata){ 
    If ($xmldata.GPO.User.ExtensionData -eq $null) {Return $true} 
    Return $false 
} 

function NoComputerChanges($xmldata){ 
    If ($xmldata.GPO.Computer.ExtensionData -eq $null) {Return $true} 
    Return $false 
} 

#Function for Group membership exporting
function GetGroupMembers($GroupName) {
    $OutputList = @()
    $MemberList = Get-ADGroupMember -Identity $GroupName -Recursive
    foreach ($user in $Memberlist) { $Outputlist += $user }
    $OutputList | Select Name, samAccountName | Export-csv "$AnalysisTempDir\$GroupName.csv"
}


# GPO report mainline
 Write-host "Analyzing GPOs for issues..."

    $unlinkedGPOs = @()
    $noUserConfigs = @()
    $noComputerConfigs = @()
    $AllGPOs = @() 

    Get-GPO -All -server $DomainController | ForEach { $gpo = $_ | Get-GPOReport -ReportType xml | ForEach { 
        If(IsNotLinked([xml]$_)){$unlinkedGPOs += $gpo} 
        If(NoUserChanges([xml]$_)){$NoUserConfigs += $gpo} #actually detects user part disabled
        If(NoComputerChanges([xml]$_)){$NoComputerConfigs += $gpo} } #actually detects computer part disabled
        $AllGPOs += $_
    }
 
    $unlinkedGPOs | Select DisplayName,ID | export-csv $AnalysisTempDir\UnlinkedGPO.csv
    $noUserConfigs| Select DisplayName,ID | export-csv $AnalysisTempDir\NoUserConfigs.csv
    $NoComputerConfigs | Select DisplayName,ID | export-csv $AnalysisTempDir\NoComputerConfigs.csv
    $AllGPOs | Select Displayname, Description, GPOstatus, CreationTime, ModificationTime, WMIfilter, Owner | export-csv $AnalysisTempDir\AllGPOs.csv

# User Reporting
Write-host "Analysing user accounts..."

$DisabledUsers = @()
$NonExpiringUsers = @()
$NinetyDayUsers = @()
$PasswordNotRequired = @()
$NeverUsedAccounts = @()
$StalePasswords = @()

get-aduser -server $DomainControllerADWS -f * -properties Name, PasswordNeverExpires, PasswordNotRequired, Lastlogontimestamp, Enabled, PwdLastSet | foreach ($_) { 
    $Identity = $_
    If ($_.Enabled -eq $false) { $DisabledUsers += $_ }
    If ($_.PasswordNeverExpires) { $NoNExpiringUsers += $_ }
    If ($_.LastLogontimestamp -lt $90Days) { $NinetyDayUsers += $_ }
	If ($_.PasswordNotRequired) { $PasswordNotRequired += $_ }
    if (($_.lastlogontimestamp -eq $null) -and ($_.enabled -eq $true)) {$NeverUsedAccounts+= $_}
    If ($_.pwdlastset -lt $90Days) { $StalePasswords += $_ }   
}

$DisabledUsers| Select Name | export-csv $AnalysisTempDir\DisabledUsers.csv    
$NonExpiringUsers| Select Name | export-csv $AnalysisTempDir\NonExpiringUsers.csv    
$NinetyDayUsers| Select Name | export-csv $AnalysisTempDir\NinetyDayUsers.csv    
$PasswordNotRequired| Select Name | export-csv $AnalysisTempDir\PasswordNotRequired.csv    
$NeverUsedAccounts | Select Name | export-csv $AnalysisTempDir\NeverUsedAccounts.csv
$StalePasswords | Select Name | export-csv $AnalysisTempDir\Stalepasswords.csv

# Group analysis
Write-host "Analyzing groups..."

$AllGroups = @()
$EmptyGroups = @()
$NestedGroups = @()
$MailEnabledGroups =  @()

get-adgroup -f * -Properties Name,GroupCategory,GroupScope,Description,member,mail,ManagedBy,memberOf | Foreach ($_) {
    if ($_.mail) {$MailEnabledGroups += $_}
    if ($_.member.count -eq "0") {$EmptyGroups += $_}
    if ($_.MemberOf) {$NestedGroups += $_}
    $AllGroups += $_
}

$MailEnabledGroups | Select Name,GroupCategory,GroupScope,Description,Mail | export-csv $AnalysisTempDir\MailEnabledGroups.csv
$EmptyGroups | Select Name,GroupCategory,GroupScope,Description| export-csv $AnalysisTempDir\EmptyGroups.csv
$NestedGroups| Select Name,GroupCategory,GroupScope,Description,@{l='MemberOf'; e= { ( $_.memberof | % { (Get-ADObject $_).Name }) -join "," }} | export-csv $AnalysisTempDir\NestedGroups.csv -notypeinformation
$AllGroups | Select Name,GroupCategory,GroupScope,Description,mail,ManagedBy,@{l='MemberOf'; e= { ( $_.member | % { (Get-ADObject $_).Name }) -join "," }} | export-csv $AnalysisTempDir\AllGroups.csv -notypeinformation

#Get all privilege group members
GetGroupMembers("Domain Admins")
GetGroupMembers("Enterprise Admins")
GetGroupMembers("Schema Admins")

# Poll for all DHCP Servers
Write-host "Enumerating DHCP servers..."

$ConfigurationSearchBase = "cn=configuration,$RootDN"

$DHCPServers = @()

#$ConfigurationSearchBase = "cn=configuration,$RootDN" #direct link into ADSI
Get-ADObject -SearchBase $ConfigurationSearchBase -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'" -properties Name | foreach ($_) {
	$DHCPServers += $_
}

$DHCPServers | select Name | export-csv $AnalysisTempDir\DHCPServers.csv

# PKI servers
Write-host "Enumerating PKI..."

$CertificateAuthorities = @()

Get-ADObject -SearchBase $ConfigurationSearchBase  -Filter "objectclass -eq 'certificationAuthority' " -properties Name | foreach ($_) {
	$CertificateAuthorities += $_
}

$CertificateAuthorities | select Name | export-csv $AnalysisTempDir\CertificateAuthorities.csv

# Pull all system OS from AD
Write-host "Analyzing computer objects..."

#Workstations
$WindowsXP = @()
$Windows8 = @()
$Windows7 = @()
$Windows10 = @()
#Server
$Windows2000 = @()
$Windows2003 = @()
$Windows2008 = @()
$Windows2012 = @()
$Windows2016 = @()
$Windows2019 = @()
#Miscellaneous
$UnknownOS = @()

# For disabled, unused, and missing systems

$DisabledComputers = @()
$NinetyDayComputers = @()
$NeverUsedComputers = @()

Get-ADComputer -f * -properties Name,OperatingSystem,LastLogon,WhenCreated,OperatingSystemVersion,lastlogontimestamp| foreach ($_) {
    switch ($_) {
        {$PSItem.OperatingSystem -like 'Windows Server 2019*'} {$Windows2019 += $_;continue} 
        {$PSItem.OperatingSystem -like 'Windows Server 2016*'} {$Windows2016 += $_;continue} 
        {$PSItem.OperatingSystem -like 'Windows Server 2012*'} {$Windows2012 += $_;continue}
        {$PSItem.OperatingSystem -like 'Windows Server 2008*'} {$Windows2008 += $_;continue}
        {$PSItem.OperatingSystem -like 'Windows Server 2003*'} {$Windows2003 += $_;continue}
        {$PSItem.OperatingSystem  -like 'Windows 2000 Server*'} {$Windows2000 += $_;continue}
        {$PSItem.OperatingSystem -like 'Windows 7*'} {$Windows7 += $_;continue}
        {$PSItem.OperatingSystem -like 'Windows 8*'} {$Windows8 += $_;continue}
        {$PSItem.OperatingSystem  -like 'Windows 10*'} {$Windows10 += $_;continue}
        {$PSItem.OperatingSystem  -like 'Windows XP*'} {$WindowsXP += $_;continue}
        default {$UnknownOS += $_}
    }
    If ($_.Enabled -eq $false) { $DisabledComputers += $_ }
    If ($_.LastLogontimestamp -lt $90Days) { $NinetyDayComputers += $_ }
    if (($_.lastlogontimestamp -eq $null) -and ($_.enabled -eq $true)) {$NeverUsedComputers+= $_}   
}

$WindowsXP | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\WindowsXP.csv
$Windows8 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows8.csv
$Windows7 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows7.csv
$Windows10 | Select Name,OperatingSystem,OperatingSystemVersion,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows10.csv
$Windows2000 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2000.csv
$Windows2003 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2003.csv
$Windows2008 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2008.csv
$Windows2012 | Select Name,OperatingSystem,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2012.csv
$Windows2016 | Select Name,OperatingSystem,OperatingSystemVersion,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2016.csv
$Windows2019 | Select Name,OperatingSystem,OperatingSystemVersion,WhenCreated,LastLogon | export-csv $AnalysisTempDir\Windows2019.csv

$DisabledComputers | Select Name, OperatingSystem, whencreated, lastlogon | export-csv $AnalysisTempDir\DisabledComputers.csv
$NinetyDayComputers | Select Name, OperatingSystem, whencreated, lastlogon | export-csv $AnalysisTempDir\NinetyDayComputers.csv
$NeverUsedComputers | Select Name, OperatingSystem, whencreated, lastlogon | export-csv $AnalysisTempDir\NeverUsedComputers.csv

# Inspect online system services for default accounts types

Write-Host "Analyzing all systems in AD for service account usage, please be patient..."

$NonLocalServices = @()
$NonResponsiveSystems = @()

$AllComputers = get-adcomputer -filter *
$ValidServiceAccounts = @('localSystem','NT AUTHORITY\NetworkService','NT AUTHORITY\LocalService')
Foreach ($system in $AllComputers) {
    try {
        $Services = Get-WmiObject -Class Win32_Service -ComputerName $system.name -EA Stop
        Foreach ($Servicename in $Services) {
            $count = 0
            Foreach ($ValidName in $ValidServiceAccounts) {
                if ($ServiceName.Startname -ine $ValidName) {$count++}
                if ($count -eq 3) {
                    $Device = $system.name, $Servicename.DisplayName, $Servicename.StartName, $Servicename.State -join ','
                    $NonLocalServices +=  $Device
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
