# Interrogator
This is the development of a powershell script to perform analysis of an Active Directory domain for common failures of best practices

At this time (20190812) it inspects and outputs as CSV files all of the following:

GPOs (exports as XML and HTML)

USERS

    Disabled users accounts
    Non-Expiring user accounts
    Users who haven't logged in within ninety (90) days
    Users with no password required
    User accounts that have never been used (as determined by null last logon date)

GROUPS

    All
    Mail-enabled
    Nested and the subgroups
    Groups with no members
    
COMPUTERS

    Disabled computer accounts
    Computers that haven't haven't logged in within ninety (90) days
    Computer  accounts that have never been used (as determined by null last logon date)

DHCP Servers defined in AD

PKI CAs defined in AD

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
