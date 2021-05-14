#requires -RunAsAdministrator
function Get-sbAdUserPrimarySmtp {
    [OutputType('Custom.SB.ADUser')]
    [cmdletbinding()]
    param(
        # We want to accept multiple user names from a CSV file with a UserName header
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String[]]$UserName
    )

    Begin {
    }

    Process {

        foreach ($user in $UserName) {
            $aduser = Get-AdUser -Filter "samAccountName -eq '$user'" -Properties 'proxyAddresses' -ErrorAction SilentlyContinue

            # Making sure the user exists - will write warning if not.
            if ($aduser) {

                $primarysmtp = $aduser.proxyAddresses | Where-Object { $_ -clike '*SMTP*' }
                $primarysmtp = $primarysmtp -replace "SMTP:", ""

                [PSCustomObject]@{
                    PSTypeName     = "Custom.SB.ADUser"
                    SamAccountName = $aduser.SamAccountName
                    PrimarySMTP    = $primarysmtp
                }

            } else {
                Write-Warning "User [$user] not found."
            } #if $aduser
        } #foreach
    } #process

    End {
    }
} #function

function Set-sbADUserNewPrimarySMTP {

    [cmdletbinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $false)]
        [string]$NewSmtpDomain,
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [PSTypeName("Custom.SB.ADUser")][Object[]]$sbADUserPrimarySmtp
    )

    Begin {
    }

    Process {

        if ($PSCmdlet.ShouldProcess("User [$($sbADUserPrimarySmtp.SamAccountName)] with Primary SMTP [$($sbADUserPrimarySmtp.PrimarySMTP)]", "Change Primary SMTP to [$NewSmtpDomain]")) {

            $aduser = Get-AdUser -Filter "samAccountName -eq '$($sbADUserPrimarySmtp.samAccountName)'" -Properties 'proxyAddresses'
            $currentproxyaddresses = $aduser.proxyAddresses
            $newdomain = $NewSmtpDomain
            $currentprimarysmtp = $currentproxyaddresses | Where-Object { $_ -clike '*SMTP*' }
            $localpart = $currentprimarysmtp -replace "^SMTP:|@(.*)$", ""
            $newprimary = $localpart, $newdomain -join "@"
            $oldprimary = $currentprimarysmtp -replace "^SMTP:", ""


            if ($currentproxyaddresses -ccontains "SMTP:$newprimary") {
                Write-Warning "Already Done!`n"
            } else {
                $backuppath = "c:\users\$env:username\desktop\backups"
                $backupfile = "$backuppath\$($aduser.samAccountName)-proxyAddresses.txt"
                if (!(Test-Path $backuppath)) {
                    New-Item -ItemType Directory -Path $backuppath
                } else { }
                Write-Verbose "Backing up current proxyAddresses to [$backupfile]`n"
                $currentproxyaddresses | Out-File $backupfile
                $newproxyaddresses = $currentproxyaddresses -creplace "SMTP:$oldprimary", "SMTP:$newprimary"

                if ($newproxyaddresses -ccontains "smtp:$oldprimary") {
                    Write-Warning "Old smtp address already exists for [$($aduser.samAccountName)]"
                } elseif ($newproxyaddresses -ccontains "smtp:$newprimary") {
                    Write-Warning "New address present as secondary - converting to old`n"
                    $newproxyaddresses = $newproxyaddresses -creplace "smtp:$newprimary", "smtp:$oldprimary"
                } else {
                    Write-Warning "Adding old smtp: [$oldprimary]`n"
                    $newproxyaddresses += "smtp:$oldprimary"

                } #if new/old addresses already exist as secondary smtp

                Write-Warning "Committing Changes..."
                Set-ADUser $aduser.samAccountName -replace @{proxyaddresses = $newproxyaddresses }

            } #if alreadydone

            # Outputting Results - getting fresh proxyAddresses for user
            Write-Verbose "Getting current user proxyAddresses"
            Get-AdUser -Filter "samAccountName -eq '$($sbADUserPrimarySmtp.samAccountName)'" -Properties 'proxyAddresses'
        } #if shouldprocess

    }

    End {
        Write-Host "Conversion(s) complete. Old settings backed up to [$backuppath]." -ForegroundColor Green
    }
} #function

function Get-sbExoUserPrimarySmtp {
    [OutputType('Custom.SB.ExoUser')]
    [cmdletbinding()]
    param(
        # We want to accept multiple user names from a CSV file with a UserName header
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias ('UserName')]
        [String[]]$Identity
    )

    Begin {
        # Checking that we're connected to Exchange Online
        $pssession = Get-PSSession | Where-Object { $_.ComputerName -eq "outlook.office365.com" }
        if (!($pssession)) {
            Throw "Please connect to Exchange Online Powershell and try again."
        }

    }

    Process {

        foreach ($user in $Identity) {
            $exouser = Get-Mailbox -Identity "$User" -ErrorAction SilentlyContinue

            # Making sure the user exists - will write warning if not.
            if ($exouser) {

                $primarysmtp = $exouser.EmailAddresses | Where-Object { $_ -clike '*SMTP*' }
                $primarysmtp = $primarysmtp -replace "SMTP:", ""

                [PSCustomObject]@{
                    PSTypeName  = "Custom.SB.ExoUser"
                    Identity    = $exouser.Identity
                    PrimarySMTP = $primarysmtp
                }

            } else {
                Write-Warning "User [$user] not found."
            } #if $exouser
        } #foreach
    } #process

    End {
    }
} #function

function TS {
    Get-Date -Format 'HH:mm:ss'
} #function

function Write-Log {
    Param(
        [Parameter()]
        [string]$Message,
        [Parameter()]
        [string]$Path = "c:\users\$env:username\desktop\backups\$(Get-Date -Format 'dd MMM yy')-log.txt",
        [Parameter()]
        [switch]$Warning = $false,
        [Parameter()]
        [switch]$Silent = $false
    )
    if ( $Silent ) {
        "[$(TS)] $Message" | Tee-Object -FilePath $Path -Append | Write-Verbose
    } else {
        if ( $Warning ) {
            "[$(TS)] $Message" | Tee-Object -FilePath $Path -Append | Write-Warning
        } else {
            "[$(TS)] $Message" | Tee-Object -FilePath $Path -Append | Write-Host -ForegroundColor Green
        } #ifwarning
    } #ifsilent

} #function


function Set-sbExoUserNewPrimarySMTP {
    # TODO Add/change a parameter to apply this to an individual user outside the pipeline
    [cmdletbinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $false)]
        [string]$NewSmtpDomain,
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [PSTypeName("Custom.SB.ExoUser")][Object[]]$sbExoUserPrimarySmtp
    )

    Begin {
        Write-Log "+++==========================================+++" -Silent
        Write-Log "Primary SMTP address conversion process started." -Silent
        Write-Log "Verifying connection to Microsoft 365..." -Silent
        $pssession = Get-PSSession | Where-Object { $_.ComputerName -eq "outlook.office365.com" }
        if (!($pssession)) {
            Throw "Please connect to Exchange Online Powershell and try again."
        }
        Write-Log "Connection verified. Continuing.`n" -Silent

        $ChangesMade = $false
    }

    Process {

        if ($PSCmdlet.ShouldProcess("User [$($sbExoUserPrimarySmtp.Identity)] with Primary SMTP [$($sbExoUserPrimarySmtp.PrimarySMTP)]", "Change Primary SMTP to [$NewSmtpDomain]")) {

            $exouser = Get-Mailbox -Identity "$($sbExoUserPrimarySmtp.Identity)"
            $currentemailaddresses = $exouser.emailaddresses
            $newdomain = $NewSmtpDomain
            $currentprimarysmtp = $currentemailaddresses | Where-Object { $_ -clike '*SMTP*' }
            $localpart = $currentprimarysmtp -replace "^SMTP:|@(.*)$", ""
            $newprimary = $localpart, $newdomain -join "@"
            $oldprimary = $currentprimarysmtp -replace "^SMTP:", ""


            if ($currentemailaddresses -ccontains "SMTP:$newprimary") {
                Write-Log "[$exouser] already has [$newprimary] as Primary SMTP!`n" -Warning
            } else {
                $ChangesMade = $true
                $backuppath = "c:\users\$env:username\desktop\backups"
                $backupfile = "$backuppath\$($exouser.Identity)-emailaddresses.txt"
                if (!(Test-Path $backuppath)) {
                    New-Item -ItemType Directory -Path $backuppath
                } else { }
                Write-Log "Backing up current emailaddresses to [$backupfile]`n" -Silent
                $currentemailaddresses | Out-File $backupfile
                $newemailaddresses = $currentemailaddresses -creplace "SMTP:$oldprimary", "SMTP:$newprimary"

                # I think this initial "if" is checking for an impossible situation, where the current Primary SMTP
                # address also exists as an Alias. Which cannot happen as far as I know.
                if ($newemailaddresses -ccontains "smtp:$oldprimary") {
                    Write-Log "Current Primary [$oldprimary] already exists as Alias for [$($exouser.Identity)]." -Warning
                } elseif ($newemailaddresses -ccontains "smtp:$newprimary") {
                    Write-Log "New Primary [$newprimary] present as Alias - will become Primary." -Warning
                    Write-Log "Current Primary [$oldprimary] will become an Alias.`n" -Warning
                    $newemailaddresses = $newemailaddresses -creplace "smtp:$newprimary", "smtp:$oldprimary"
                } else {
                    Write-Log "New Primary not present as Alias. Will add [$newprimary] to email address list." -Warning
                    Write-Log "Current Primary [$oldprimary] will become an Alias.`n" -Warning
                    $newemailaddresses += "smtp:$oldprimary"

                } #if new/old addresses already exist as secondary smtp

                Write-Log "Committing Changes for [$($exouser.Identity)]..." -Warning
                Set-Mailbox -Identity $exouser.Identity -EmailAddresses $newemailaddresses

            } #if alreadydone

        } #if shouldprocess

        # Outputting Results - getting fresh emailaddresses for user

        if ( $ChangesMade ) {

            Write-Log "Getting post-conversion email addresses for [$($exouser.Identity)]" -Silent
            $afterchanges = Get-Mailbox -Identity "$($sbExoUserPrimarySmtp.Identity)"
            Write-Log " [$($exouser)] now has the following email addresses: "
            Write-Log " [$($afterchanges.EmailAddresses)]"

            # [PSCustomObject]@{
            #     Name           = $afterchanges.Identity
            #     EmailAddresses = $afterchanges.EmailAddresses
            # }

        } #ifchangesmade

    }

    End {

        if ( $ChangesMade ) {
            Write-Log "Conversion(s) complete. Old settings backed up to [$backuppath]."
        } else {
            Write-Log "No changes made - no backup taken."
        } #ifchangesmade

    }

} #function

function Get-sbMailboxesWithoutNewSMTP {

    [cmdletbinding()]
    param(
        # We want to accept multiple user names from a CSV file with a UserName header
        [Parameter(Mandatory = $True, ValueFromPipeline = $false)]
        [string]$NewSmtpDomain,
        [Parameter()]
        [switch]$IncludeSharedMailboxes = $false,
        [Parameter()]
        [switch]$IncludeOnPremiseMailboxes = $false
    )

    Begin {
        # Checking that we're connected to Exchange Online
        $pssession = Get-PSSession | Where-Object { $_.ComputerName -eq "outlook.office365.com" }
        if (!($pssession)) {
            Throw "Please connect to Exchange Online Powershell and try again."
        }
    } #begin

    Process {
        # TODO Find a way to filter out Cloud-Only Mailboxes and make this a parameter.Property = IsDirSynced
        # Getting all mailboxes with/without shared mailboxes as required. Excluding DiscoverySearch in all cases.
        if ( $IncludeSharedMailboxes ) {
            Write-Verbose "Getting all User and Shared Mailboxes, excluding DiscoverySearchMailbox..."
            Write-Verbose "Searching for occurrences of [$NewSmtpDomain] in their Email Addresses..."
            $AllMailboxes = Get-Mailbox | Where-Object { $_.Alias -notlike "*DiscoverySearchMailbox*" }
            Write-Verbose "All returned Identities below do NOT have an address (Primary or Alias) ending [@$NewSmtpDomain]."
        } else {
            Write-Verbose "Getting all User Mailboxes, excluding Shared Mailboxes and DiscoverySearchMailbox..."
            Write-Verbose "Searching for occurrences of [$NewSmtpDomain] in their Email Addresses..."
            $AllMailboxes = Get-Mailbox | Where-Object { $_.IsShared -eq $false -and $_.Alias -notlike "*DiscoverySearchMailbox*" }
            Write-Verbose "All returned Identities below do NOT have an address (Primary or Alias) ending [@$NewSmtpDomain]."
        } # ifincludesharedmailboxes

        if ( $IncludeOnPremiseMailboxes ) {
            Write-Verbose "Mailboxes synced from On-Premises Exchange ARE included."
        } else {
            $AllMailboxes = $AllMailboxes | Where-Object { $_.IsDirSynced -eq $false }
        } #ifincludeonpremisemailboxes

        foreach ($Mailbox in $AllMailboxes) {
            $SearchAddress = "smtp:$($Mailbox.Alias)@$NewSmtpDomain"
            if ( $($Mailbox).EmailAddresses -contains $SearchAddress ) {

            } else {

                $MailboxNoMatch = [PSCustomObject]@{
                    PSTypeName = "Custom.SB.ExoUser"
                    Identity   = $Mailbox.Identity
                }

                $MailboxNoMatch

            } #ifmailboxcontainssearchaddress

        } #foreach

    } #process

    End {

    } #end

} #function
