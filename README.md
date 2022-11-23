# ADAudit
Mail a formatted list of account audit events from a domain controller.

A simple PowerShell script to be run on schedule. Extracts events from the Security event log for account and group changes formatted as a simple table.  
It is very hacky.

The source includes an annotated list of EventIDs to be extracted.

## Requirements before running
- Change the three mail variables at the top of the source file.
- Change the two DN strings at source [line 100](https://github.com/bluikko/ADAudit/blob/45c9fcdd93385172b4ffe24259421256eb61e4fa/ADAudit.ps1#L100) used for making LDAP paths shorter.
- The script is expected to be installed in `C:\Program Files\Scripts\` and as such a subdirectory `Data` (`C:\Program Files\Scripts\Data\`) is expected to exist for saving CSV-formatted reports.
