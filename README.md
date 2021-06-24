# change-ad-user-default-smtp

Script to change SMTP: Attribute for one or more users

To install or update this module:

`Import-Module .\ChangePrimarySmtpAddress\ChangePrimarySmtpAddress.psm1 -Force`

Make sure to connect to Exchange Online with `Connect-ExchangeOnline` before attempting to run any commands against
Exchange Online.

Use `Get-Command -Module ChangePrimarySmtpAddress` for a list of available cmdlets.

To check for all cloud mailboxes which have not yet been converted to the new desired primary smtp domain:

`Get-sbMailboxesWithoutNewSMTP -NewSmtpDomain new-primary-domain.com`

This will return all non-DirSynced mailboxes which have not yet had the new desired primary email domain applied,
it will not return Shared mailboxes.

If you want to include mailboxes synced from AD, or Shared mailboxes, use the `-IncludeSynced` or `-IncludeShared`
parameters, respectively.

You can pipe the results of `Get-sbMailboxesWithoutNewSMTP` into `Set-sbExoUserNewPrimarySMTP` to update all
the listed mailboxes with the new primary smtp domain.

A log of all changes is kept if you need to revert.
