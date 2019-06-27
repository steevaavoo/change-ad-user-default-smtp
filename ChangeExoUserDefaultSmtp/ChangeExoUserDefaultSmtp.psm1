#requires -RunAsAdministrator
function Get-sbExoUserPrimarySmtp {
    [OutputType('Custom.SB.ExoUser')]
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
            $exouser = Get-Mailbox -Identity "*$User*"

            # Making sure the user exists - will write warning if not.
            if ($exouser) {

                $primarysmtp = $exouser.EmailAddresses | Where-Object { $_ -clike '*SMTP*' }
                $primarysmtp = $primarysmtp -replace ‘SMTP:’, ''

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

function Set-sbExoUserNewPrimarySMTP {

    [cmdletbinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [PSTypeName("Custom.SB.ExoUser")][Object[]]$sbExoUserPrimarySmtp,
        [Parameter(Mandatory = $True, ValueFromPipeline = $false)]
        [string]$NewSmtpDomain
    )

    Begin {
    }

    Process {

        if ($PSCmdlet.ShouldProcess("User [$($sbExoUserPrimarySmtp.Identity)] with Primary SMTP [$($sbExoUserPrimarySmtp.PrimarySMTP)]", "Change Primary SMTP to [$NewSmtpDomain]")) {

            $exouser = Get-Mailbox -Identity "*$($sbExoUserPrimarySmtp.Identity)*"
            $currentemailaddresses = $exouser.emailaddresses
            $newdomain = $NewSmtpDomain
            $currentprimarysmtp = $currentemailaddresses | Where-Object { $_ -clike '*SMTP*' }
            $localpart = $currentprimarysmtp -replace '^SMTP:|@(.*)$', ''
            $newprimary = $localpart, $newdomain -join '@'
            $oldprimary = $currentprimarysmtp -replace '^SMTP:', ''


            if ($currentemailaddresses -ccontains "SMTP:$newprimary") {
                Write-Warning "Already Done!`n"
            } else {
                $backuppath = "c:\users\$env:username\desktop\backups"
                $backupfile = "$backuppath\$($exouser.Identity)-emailaddresses.txt"
                if (!(Test-Path $backuppath)) {
                    New-Item -ItemType Directory -Path $backuppath
                } else { }
                Write-Verbose "Backing up current emailaddresses to [$backupfile]`n"
                $currentemailaddresses | Out-File $backupfile
                $newemailaddresses = $currentemailaddresses -creplace "SMTP:$oldprimary", "SMTP:$newprimary"

                if ($newemailaddresses -ccontains "smtp:$oldprimary") {
                    Write-Warning "Old smtp address already exists for [$($exouser.Identity)]"
                } elseif ($newemailaddresses -ccontains "smtp:$newprimary") {
                    Write-Warning "New address present as secondary - converting to old`n"
                    $newemailaddresses = $newemailaddresses -creplace "smtp:$newprimary", "smtp:$oldprimary"
                } else {
                    Write-Warning "Adding old smtp: [$oldprimary]`n"
                    $newemailaddresses += "smtp:$oldprimary"

                } #if new/old addresses already exist as secondary smtp

                Write-Warning "Committing Changes..."
                Set-Mailbox -Identity $exouser.Identity -EmailAddresses $newemailaddresses

            } #if alreadydone

            # Outputting Results - getting fresh emailaddresses for user
            Write-Verbose "Getting current user emailaddresses"
            $afterchanges = Get-Mailbox -Identity "*$($sbExoUserPrimarySmtp.Identity)*"

            [PSCustomObject]@{
                Name           = $afterchanges.Identity
                EmailAddresses = $afterchanges.EmailAddresses
            }

        } #if shouldprocess

    }

    End {
    }
} #function