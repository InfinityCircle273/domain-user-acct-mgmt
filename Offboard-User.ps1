<#

.DESCRIPTION
    Interactive script for user offboarding in an AD/AAD-synced environment.
    - AD account disabled
    - AD account password set to random string
    - Start AD Sync
    - Connect to Exchange Online (v2)
    - Connect to Azure AD
    - Remove from distros and O365 groups
    - Convert mailbox to shared
    - Block sign-in on AAD
    - Revoke AAD refresh tokens
    - Remove assigned licenses and output for removal from billing
.NOTES
    Following modules required for successful execution:
    - ExchangeOnlineManagement (Install-Module ExchangeOnlineManagement -Force)
    - AzureAD (Install-Module AzureAD -Force)

    Version:        0.1
    Updated:        07/02/2020
    Created:        07/02/2020
    Author:         Zach Choate
    URL:            https://raw.githubusercontent.com/KSMC-TS/domain-user-acct-mgmt/Offboard-User.ps1

#>

function New-OptionList {

    <#

    .DESCRIPTION
        Use to create a Powershell option list and return the selected value.
    .PARAMETER title
        Title of the option list menu you'd like to specifiy.
    .PARAMETER optionsList
        Comma separated list of values to build your option menu from.
    .PARAMETER message
        Message to prepend before the option values and on the post selection page.
    .EXAMPLE
        $searchBy = New-OptionList -Title "Search for user by" -optionsList "displayname", "email", "userprincipalname" -message "to search by"
            This will create a menu with the options specified and return the $selectedValue to the $searchBy variable for use in the script.
    .NOTES
        Adapted from script on https://www.business.com/articles/powershell-interactive-menu/

        Version:        1.0
        Updated:        07/01/2020
        Created:        07/01/2020
        Author:         Zach Choate

    #>

    param (
           [string]$title,
           [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
           [string[]]$optionsList=@(),
           [string]$message
    )

    Clear-Host
    Write-Host "================ $title ================"
    
    $n = 1
    ForEach($option in $optionsList) {

        Write-Host "$n`: Press `'$n`' $message $option."
        $n++

    }
    Write-Host "Q: Press 'Q' to quit."
    $validationSet = '^[1-{0}]$|^Q$' -f [regex]::escape($($n-1))
    while(($optionSelected = Read-Host "Please make a selection") -notmatch $validationSet){}
    If($optionSelected -eq "q") {
        Write-Host "Quiting..."
        Exit
    }
    $optionSelected = $optionSelected-1
    $selectedValue = $optionsList[$optionSelected]

    Write-Host "`nYou selected $message $selectedValue. `nMoving on..."
    Return $selectedValue

}

function Get-RandomString {
    param (
        [int] $upperCase = 2,
        [int] $lowerCase = 2,
        [int] $numbers = 4
    )
    
    Write-Output ( -join ([char[]](65..90) | Get-Random -Count $upperCase) + ([char[]](97..122) | Get-Random -Count $lowerCase) + ((1..9) | Get-Random -Count $numbers)).Replace(" ","")

}

# Search for user based on their name, email, UPN, or SAM
$searchBy = New-OptionList -Title "Search for user by" -optionsList "name", "mail", "userprincipalname","samaccountname" -message "to search by"
Clear-Host
$searchString = Read-Host "Enter the user's $searchBy."
$job = Start-Job -ScriptBlock {
    $filter = "$($args[0]) -like `"*$($args[1])*`""
    $users = Get-ADUser -Filter $filter -Properties *
    $users
} -ArgumentList $searchBy, $searchString
while ($job.State -ne "Completed") {
    Write-Host "Click clack..."
    Start-Sleep -Milliseconds 250
    Write-Host "Click clack..."
    Start-Sleep -Milliseconds 275
    Write-Host "Searching for $searchString..."
    Start-Sleep -Milliseconds 300
}
$users = Receive-Job -Job $job
Clear-Host
If($users.Count -gt 1) {
    Write-Host "It looks like the search returned multiple users. Please select the user from the list below:"
    $n = 1
    ForEach($u in $users) {
        Write-Host "$n`: Press `'$n`' to select Name: $($u.name) | Email Address: $($u.mail) | UserPrincipalName: $($u.userprincipalname) | SamAccountName: $($u.samaccountname)"
        $n++
    }
    $optionSelected = Read-Host "Please make a selection"
    $optionSelected = $optionSelected-1
    $user = $users[$optionSelected]
} elseif($users) {
    Write-Host "Your search yielded the following: `nName: $($user.Name) | Email Address: $($user.mail) | UserPrincipalName: $($user.userprincipalname) | SamAccountName: $($user.samaccountname)"
    $user = $users
} else {
    Write-Host "Whoops, it looks like we couldn't find anything matching your search. Try again..."
    Exit
}
Remove-Job -Job $job

Clear-Host
while(($confirm = Read-Host "Are you sure you want to disable $($user.Name) with the username $($user.userprincipalname)? [y/n]") -notmatch '^Y$|^N$'){}
If($confirm -ne "y") {
    Write-Host "Okay, aborting..."
    Exit
}

# Start disabling user, resetting password and syncing to Azure AD.
Clear-Host
$passString = Get-RandomString
$disableResetJob = Start-Job -ScriptBlock {
    Try {
        Set-ADUser -Identity $args[0] -Enabled $false
        Set-ADAccountPassword -Identity $args[0] -Reset -NewPassword $($args[1] | ConvertTo-SecureString -AsPlainText -Force)
        Import-Module ADSync
        Stop-ADSyncSyncCycle
        Start-ADSyncSyncCycle -PolicyType Delta
    } catch {
        $disableError = $true
    }
    $disableError
} -ArgumentList $($user.DistinguishedName), $passString
Write-Host "Disabling $($user.Name) and resetting password..."
Clear-Host
$disableError = Receive-Job -Job $disableResetJob
If($disableError) {
    Write-Host "Disabling $($user.Name) and resetting password appears to have failed. Further investigation required."
    Pause
} else {
    Write-Host "Disabled $($user.Name) and reset password successfully.`nMoving to the next step..."
}

# Select Office 365 Environment and connect to Exchange Online
Clear-Host
$o365Env = New-OptionList -Title "Select Office 365 Environment" -optionsList "Commerical", "21Vianet", "Germany", "Government Community Cloud High"<#, "DoD"#> -message "for"
Clear-Host

switch ($o365Env) {
    "Commerical" {
        $connectionUri = "https://outlook.office365.com/powershell-liveid/"
        $azureEnv = "AzureCloud"
    }
    "21Vianet" {
        $connectionUri = "https://partner.outlook.cn/PowerShell"
        $azureEnv = "AzureChinaCloud"
    }
    "Germany" {
        $connectionUri = "https://outlook.office.de/powershell-liveid/"
        $azureEnv = "AzureGermanyCloud"
    }
    "Government Community Cloud High" {
        $connectionUri = "https://outlook.office365.us/powershell-liveid/"
        $azureEnv = "AzureUSGovernment"
    }
    #"DoD" {$connectionUri = "https://webmail.apps.mil/powershell-liveid/"}
}

Write-Host "`nConnecting to Office 365 and AzureAD..."

Import-Module ExchangeOnlineManagement
Import-Module AzureAD
Connect-ExchangeOnline -ConnectionUri $connectionUri
Connect-AzureAD -AzureEnvironmentName $azureEnv

Write-Host "Getting user's Exchange Online details and group memberships. This may take a minute..."
$userDn = Get-User -Identity $($user.userprincipalname) | Select-Object -ExpandProperty DistinguishedName
$groupsFilter = "Members -eq '" + $userDn + "'"
$groups = Get-Recipient -Filter $groupsFilter -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select-Object Name,RecipientTypeDetails,DistinguishedName
ForEach($group in $groups) {
    If($group.RecipientTypeDetails -eq "GroupMailbox") {
        Remove-UnifiedGroupLinks -Identity $group.DistinguishedName -LinkType Members -Links $userDn
        Write-Host "Removed $($user.Name) from $($group.Name)."
    } else {
        Remove-DistributionGroupMember -Identity $group.DistinguishedName -Member $userDn
        Write-Host "Removed $($user.Name) from $($group.Name)."
    }
}

# Convert to shared mailbox
Set-Mailbox -Identity $userDn -Type Shared
Write-Host "Converted $($user.Name)'s mailbox to a shared mailbox."

# Disable AAD User
$aadUser = Get-AzureADUser -ObjectId $($user.userprincipalname) | Select-Object ObjectId,UserPrincipalName
Set-AzureADUser -ObjectId $aadUser.ObjectId -AccountEnabled $false
Write-Host "Disabled $($user.Name)'s account in Azure AD."


# Immediately revoke any refresh tokens in AAD
Revoke-AzureADUserAllRefreshToken -ObjectId $aadUser.ObjectId
Write-Host "Revoked all existing refresh tokens for Azure AD."

# Remove licenses from user
$aadUserLicenses = Get-AzureADUser -ObjectId $aadUser.ObjectId | Select-Object -ExpandProperty AssignedLicenses | Select-Object SkuID
$license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$license.SkuID = $aadUserLicenses.SkuID
$licenses.AddLicenses = $license
Set-AzureAdUserLicense -ObjectId $aadUser.ObjectId -AssignedLicenses $licenses
$licenses.AddLicenses = @()
$licenses.RemoveLicenses = (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $aadUserLicenses.SkuID -EQ).SkuID
Set-AzureAdUserLicense -ObjectId $aadUser.ObjectId -AssignedLicenses $licenses
$licensesRemoved = $licenses.RemoveLicenses -split " "
$licenseList = Get-AzureADSubscribedSku | Select-Object SkuPartNumber,SkuID
$licensesToRemove = @()
Write-Host "The following license SKUs were removed from $($user.Name): `n$(ForEach($sku in $licensesRemoved) {
    $skuPartNumber = ($licenseList | Where-Object {$_.SkuId -contains $sku}).SkuPartNumber
    Write-Host "$skuPartNumber | $sku"
    $licensesToRemove += $skuPartNumber
})"

Clear-Host
Write-Host "All that is left is to remove licenses from billing if required. These are the licenses that were removed: `n$licensesToRemove"

Remove-Job -Job $disableResetJob
Disconnect-ExchangeOnline
Disconnect-AzureAD