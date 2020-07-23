<#

.DESCRIPTION
    Interactive script for user offboarding in an AD/AAD-synced environment.
    - AD account disabled
    - AD account password set to random string
    - Start AD Sync
    - Connect to Azure AD
    - Connect to Exchange Online (v2)
    - Remove from distros and O365 groups
    - Convert mailbox to shared
    - Block sign-in on AAD
    - Revoke AAD refresh tokens
    - Remove assigned licenses and output for removal from billing
.NOTES
    Following modules required for successful execution:
    - ExchangeOnlineManagement (Install-Module ExchangeOnlineManagement -Force)
    - AzureAD (Install-Module AzureAD -Force)

    Version:        0.2
    Updated:        07/22/2020
    Created:        07/02/2020
    Author:         Zach Choate
    URL:            https://raw.githubusercontent.com/KSMC-TS/domain-user-acct-mgmt/main/Offboard-User.ps1

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

function Disable-UserAccount {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$username,
        [Parameter(Mandatory=$true)]
        [string]$pswd
    )
    # Disable the AD account and change the password.
    Set-ADUser -Identity $userName -Enabled $false
    Set-ADAccountPassword -Identity $userName -Reset -NewPassword $($pswd | ConvertTo-SecureString -AsPlainText -Force)

    # Import ADSync module to get ADSync configuration
    Import-Module ADSync

    # Get ADSync group if applicable
    $adSyncGroup = Get-ADGroup $(((Get-ADSyncConnector).GlobalParameters | Where-Object {$_.name -eq "Connector.GroupFilteringGroupDn"} | Select-Object Value).value) -ErrorAction Ignore

    # Get user's primary group
    $primaryGroup = (Get-ADUser -Identity $userName -Properties PrimaryGroup | Select-Object PrimaryGroup).PrimaryGroup

    # Get the groups the user is a member of filtering out the AD Sync and user's primary group.
    $groups = Get-ADPrincipalGroupMembership -Identity $userName | Where-Object {$_.distinguishedName -ne $adSyncGroup.DistinguishedName -and $_.distinguishedName -ne $primaryGroup}

    # Start removing those groups from the user.
    ForEach($group in $groups) {
        Remove-ADGroupMember -Identity $group.DistinguishedName -Members $userName
    }
    Stop-ADSyncSyncCycle
    Start-Sleep -Seconds 2
    Start-ADSyncSyncCycle -PolicyType Delta
}

# Check for dependencies
$aadModule = Get-Module -Name AzureAD -ListAvailable
$exoModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable
$adsyncModule = Get-Module -Name ADSync -ListAvailable
$adModule = Get-Module -Name ActiveDirectory -ListAvailable
Clear-Host
Write-Host "================ Offboard a User ================"
If(!$adsyncModule) {
    Write-Host "Looks like the AD Sync module isn't installed. Ensure you're running this on a server with AD Sync installed."
}
If(!$adModule) {
    Write-Host "Looks like the ActiveDirectory module isn't installed. Ensure you're running this on a server with the AD management tools installed."
}
If(!$aadModule) {
    While(($aadConfirm = Read-Host "Looks like the AzureAD module isn't installed. Would you like to install the module? Note: This will open a new session as an administrator. [y/n]") -notmatch '^Y$|^N$'){}
    If($aadConfirm -ne "y") {
        Write-Host "Okay, aborting..."
    } else {
        $aadArgs = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Install-Module AzureAD -Force"
        Start-Process powershell -Verb RunAs -ArgumentList "-command $aadArgs" -Wait
        Write-Host "Looks like we have the Azure AD module installed now. Let's verify that ExchangeOnlineManagement is installed. Then we'll restart."
    }
}
If(!$exoModule) {
    While(($exoConfirm = Read-Host "Looks like the ExchangeOnlineManagement module isn't installed. Would you like to install the module? Note: This will open a new session as an administrator. [y/n]") -notmatch '^Y$|^N$'){}
    If($exoConfirm -ne "y") {
        Write-Host "Okay, aborting..."
    } else {
        $exoArgs = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Install-Module ExchangeOnlineManagement -Force"
        Start-Process powershell -Verb RunAs -ArgumentList "-command $exoArgs" -Wait
        Write-Host "Looks like we have the ExchangeOnlineManagement module installed now."
    }
}

If($aadConfirm -eq "y" -or $exoConfirm -eq "y") {
    Clear-Host
    Write-Host "Let's restart the script. We'll verify the modules are installed. If you have an issue again, try manually installing the modules and starting the script again."
    Pause
    Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File $PSCommandPath"
    Exit
} elseif (!$adsyncModule -or !$adModule -or !$aadModule -or !$exoModule) {
    Write-Host "Exiting since dependencies were not met..."
    Pause
    Exit
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
        Disable-UserAccount -UserName $arg[0] -pswd $arg[1]
    } catch {
        $disableError = $true
    }
    $disableError
} -ArgumentList $($user.DistinguishedName), $passString
Write-Host "Disabling $($user.Name),resetting password, removing groups and syncing to Azure AD..."
Clear-Host
$disableError = Receive-Job -Job $disableResetJob
If($disableError) {
    Write-Host "Disabling $($user.Name),resetting the password, removing groups, and/or syncing to Azure AD appears to have failed. Further investigation required."
    Pause
} else {
    Write-Host "Disabled $($user.Name), reset password, removed groups and started sync to Azure AD successfully.`nMoving to the next step..."
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

Write-Host "`nConnecting to Azure AD and Exchange Online..."

Import-Module ExchangeOnlineManagement
Import-Module AzureAD
$aadConnection = Connect-AzureAD -AzureEnvironmentName $azureEnv
Connect-ExchangeOnline -ConnectionUri $connectionUri -UserPrincipalName $aadConnection.Account.Id

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
Disconnect-AzureAD