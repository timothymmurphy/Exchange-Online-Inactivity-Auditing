# Exchange Online Inactivity Auditing
PowerShell scripts to proactively identify orphaned distribution lists and shared mailboxes in Exchange Online.

# Overview
Exports Message Trace statistics to a .csv for each distribution list and shared mailbox in your Exchange Online tenant. After a certain amount of time, the data from each report is parsed and a new report created for any recipient addresses that have not received any mail in the desired amount of time.

# Recommended Usage

- Schedule `Get-DistributionListReceivedMessages` and `Get-SharedMailboxReceivedMessages` to run weekly
- Schedule `New-DistributionListInactivityReport` and `New-SharedMailboxInactivityReport` to run at the appropriate cadence for your organization. Recommended to run monthly (4 weeks) or quarterly (13 weeks)

It is recommended that you update line 29/35 (depending on script) as it will currently require manual interaction to authenticate using `Connect-ExchangeOnline` so that you can use it in an automated fashion, e.g. via a Scheduled Task.

