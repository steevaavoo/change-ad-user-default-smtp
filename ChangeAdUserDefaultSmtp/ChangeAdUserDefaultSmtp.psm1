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
            $aduser = Get-AdUser -Filter "samAccountName -like '*$user*'" -Properties 'proxyAddresses'

            # Making sure the user exists - will write warning if not.
            if ($aduser) {

                $primarysmtp = $aduser.proxyAddresses | Where-Object { $_ -clike '*SMTP*' }
            $primarysmtp = $primarysmtp -replace ‘SMTP:’, ''

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
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [PSTypeName("Custom.SB.ADUser")][Object[]]$sbADUserPrimarySmtp,
        [Parameter(Mandatory = $True, ValueFromPipeline = $false)]
        [string]$NewSmtpDomain
    )

    Begin {
    }

    Process {

        if ($PSCmdlet.ShouldProcess("User [$($sbADUserPrimarySmtp.SamAccountName)] with Primary SMTP [$($sbADUserPrimarySmtp.PrimarySMTP)]", "Change Primary SMTP to [$NewSmtpDomain]")) {

            $aduser = Get-AdUser -Filter "samAccountName -like '*$($sbADUserPrimarySmtp.samAccountName)*'" -Properties 'proxyAddresses'
            $currentproxyaddresses = $aduser.proxyAddresses
            $newdomain = $NewSmtpDomain
            $currentprimarysmtp = $currentproxyaddresses | Where-Object { $_ -clike '*SMTP*' }
        $localpart = $currentprimarysmtp -replace '^SMTP:|@(.*)$', ''
        $newprimary = $localpart, $newdomain -join '@'
        $oldprimary = $currentprimarysmtp -replace '^SMTP:', ''


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
    Get-AdUser -Filter "samAccountName -like '*$($sbADUserPrimarySmtp.samAccountName)*'" -Properties 'proxyAddresses'
} #if shouldprocess

}

End {
}
} #function