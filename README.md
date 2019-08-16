# Interrogator
This is the development of a powershell script to perform analysis of an Active Directory domain for common failures of best practices.

Upon completion it opens the temporary folder location for your inspection.

At this time (20190816) it inspects and outputs as CSV files all of the following:

GPOs

    Exports as XML
    Export as HTML

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
    Others

SERVICES

    All services on all systems NOT using localSystem, NT AUTHORITY\NetworkService or NT AUTHORITY\LocalService a credentials
    Exports as a separate file non-responsive systems (offline, WMI issues, etc.)
