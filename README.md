# Interrogator
This is the development of a powershell script to perform analysis of an Active Directory domain for common failures of best practices.

Upon completion it opens the temporary folder location for your inspection.

At this time (20210330) it inspects and outputs a single XLSX file all of the following:

GPOs

    Exports as XML
    Export as HTML
    Shows if changes made to computer and user halves of a GPO
    Shows if computer and user halves of a GPO are enabled
    Shows all places the GPO links to in the AD tree
    Displays comments    

USER ACCOUNTS

    Disabled users accounts
    Non-Expiring user accounts
    User accounts who haven't logged in within ninety (90) days
    User accounts who haven't changed passwords in more than ninety (90) days
    User accounts with no password required
    Enabled user accounts that have never been used (as determined by null last logon date)

GROUPS

    Dumps all groups
    Mail-enabled groups
    Nested groups with their subgroups members
    Groups with no members
    Exports Domain, Enterprise, and Schema admins
    
COMPUTERS

    Disabled computer accounts
    Computers that haven't haven't logged in within ninety (90) days
    Computer  accounts that have never been used (as determined by null last logon date)

DHCP

    Servers defined in AD

PKI

    CAs defined in AD

OPERATING SYSTEMS

    Windows XP
    Windows 8
    Windows 10
    Windows 2000
    Windows 2003
    Windows 2008
    Windows 2012
    Windows 2016
    Windows 2019
    Non-Windows, and attempts to determine what OS (badly)

SERVICES

    All services on all systems NOT using localSystem, NT AUTHORITY\NetworkService or NT AUTHORITY\LocalService a credentials
    Separate tab non-responsive systems (offline, WMI issues, etc.) for this function as well
    NOTE: This piece has been commented out as it's dog-slow.
