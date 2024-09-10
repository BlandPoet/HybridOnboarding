<#
.SYNOPSIS
    This script creates new users in Active Directory and Exchange Online based on a CSV file.
    It Additionally copies group memberships and manager attribute from a template user.
.DESCRIPTION
    The script reads user information from a CSV file and performs the following tasks:
    - Connects to the Active Directory and Exchange Online environments.
    - Creates new AD user accounts and Exchange mailboxes.
    - Sets various user attributes such as description, display name, email address, and office.
    - Copies group memberships and manager attributes from a template user.
    - Sets a default password for the new user accounts.
.PARAMETER CSVPath
    The path to the CSV file containing user information. Verify that the file contains accurate information before running the script.
.PARAMETER LOCATION
    The parameter that shows the location of which OU the user is associated with. CHANGE THIS FOR YOUR ORG
.PARAMETER TYPE
    The parameter that is a Type of user for the template user to reference.
.PARAMETER OUOFFICEMAPPING
    This can be used if you have and OU that is named differently to what you want to apply onto a users information
        ex. the OU is named FO for front office but needing the actual name to be spelt out. 

.NOTES
    File Name      : HybridOnBoarding.ps1
    Author         : Gaines Snodderly
    Prerequisite   : Active Directory Module, Exchange Online Management Module
    Version        : 1.3
#>

Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement 

# Exchange connection
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<EXCHANGE_SERVER>/PowerShell/ -Authentication Kerberos 
Import-PSSession $Session -DisableNameChecking -CommandName New-RemoteMailbox


$CsvPath = "FILE PATH HERE FOR CSV"
$EmailDomain = "<DOMAIN>"

# Function to get OU path. This is used to add the user into the correct OU
function Get-OU {
    param (
        [string]$LOCATION,
        [string]$Type
    )
    $baseOU = switch ($Location) {
        "OU Name" { return "OU=Company,DC=domain,DC=local" }
        #add the rest of your OUs here
        default { return "OU=Company,DC=domain,DC=local" }
    }

}

# Function to copy group memberships and manager attribute
function Copy-GroupMemberships {
    param (
        [string]$TargetUsername,
        [string]$LOCATION,
        [string]$Type
    )

    # Format for template user
    $TemplateUPN = "$LOCATION-$Type@$EmailDomain" 
    
    $TemplateUserObj = Get-ADUser -Filter { UserPrincipalName -eq $TemplateUPN } -Properties MemberOf, Manager -ErrorAction SilentlyContinue
    if (-not $TemplateUserObj) {
        Write-Warning "Template user with UPN '$TemplateUPN' not found."
        return
    }

    # Copy group memberships
    $TemplateUserObj.MemberOf | ForEach-Object {
        try {
            Add-ADGroupMember -Identity $_ -Members $TargetUsername -ErrorAction Stop
        } catch {
            Write-Warning "Failed to add user '$TargetUsername' to group '$_'."
        }
    }

    # Copy manager attribute
    if ($TemplateUserObj.Manager) {
        try {
            Set-ADUser -Identity $TargetUsername -Manager $TemplateUserObj.Manager -ErrorAction Stop
        } catch {
            Write-Warning "Failed to set manager for user '$TargetUsername'."
        }
    }
}

# Mapping of OUs to office descriptions 
$ouToOfficeMapping = @{
  "EXAMPLE" = "What will be shown on account"
}

# Function to get the office description based on the OU
function Get-OfficeDescription {
    param (
        [string]$LOCATION
    )
    if ($ouToOfficeMapping.ContainsKey($LOCATION)) {
        return $ouToOfficeMapping[$LOCATION]
    } else {
        return $LOCATION  # Default to the OU name if no mapping is found
    }
}

# Function to create a new AD user
function New-User {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$Type,
        [string]$Title,
        [string]$Fullname,
        [string]$EmployeeID,
        [string]$LOCATION
    )

    # Check if LastName is provided
    if ([string]::IsNullOrEmpty($LastName)) {
        Write-Warning "Skipping user creation for ${Fullname}: LastName is null or empty."
        return
    }

    # Generate a unique SamAccountName
    $SamAccountName = if ($Type -eq "UNIQUE TITLE IF THERE USERNAME NEEDS TO BE FORMATTED DIFFERENTLY") { "$($FirstName.Substring(0,1))$LastName" } else { "$FirstName$($LastName.Substring(0,1))" }
    
    # Initialize the length of the substring to be used from the last name
    $lastNameLength = 1
    
    while (Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue) {
        $lastNameLength++
        if ($Type -eq "UNIQUE TITLE IF THERE USERNAME NEEDS TO BE FORMATTED DIFFERENTLY") {
            $SamAccountName = "$($FirstName.Substring(0,1))$($LastName.Substring(0, $lastNameLength))"
        } else {
            $SamAccountName = "$FirstName$($LastName.Substring(0, $lastNameLength))"
        }
    }
    
    # Output the final SamAccountName for verification
    Write-Host "Generated SamAccountName: $SamAccountName"

    # Adjust DisplayName if Unique type needs to be applied to user
    # $DisplayName = if ($Type -eq "UNIQUE TITLE IF THERE USERNAME NEEDS TO BE FORMATTED DIFFERENTLY" -and $Title -ne "NP") { "Dr. $Fullname" } else { $Fullname }

    # user properties table
    $UserParams = @{
        Name                        = $Fullname
        FirstName                   = $FirstName
        LastName                    = $LastName
        SamAccountName              = $SamAccountName
        UserPrincipalName           = "$SamAccountName@$EmailDomain"
        OnPremisesOrganizationalUnit = Get-OU -Location $LOCATION -Type $Type
        DisplayName                 = $DisplayName
        PrimarySmtpAddress          = "$FirstName$LastName@$EmailDomain"
    }
    
    try {
        New-RemoteMailbox @UserParams -ErrorAction Stop 
        Write-Host "Created Exchange mailbox for $Fullname"
    }
    catch {
        Write-Warning "Failed to create Exchange mailbox for user $Fullname. $_"
    }   

        try {
            Set-ADUser -Identity $SamAccountName `
                -Description "$LOCATION - $Title" `
                -DisplayName $DisplayName `
                -EmailAddress "$FirstName$LastName@$EmailDomain" `
                -ChangePasswordAtLogon $true `
                -Replace @{ 'extensionAttribute1' = $EmployeeID } # This can be changed if needed
                -Office "$office" `
                
    
        Write-Host "Set additional attributes for $SamAccountName"
    }
        catch {
        Write-Warning "Failed to set additional attributes for user $SamAccountName. $_"
    }

    # copy group memberships and manager attribute from template
    Copy-GroupMemberships -TargetUsername $SamAccountName -Location $LOCATION -Type $Type
    
    $password = ConvertTo-SecureString "Default Password Here" -AsPlainText -Force
    Set-ADAccountPassword -Identity $SamAccountName -NewPassword $password -Reset -ErrorAction Stop
}
# Import CSV and process users
$Users = Import-Csv -Path $CsvPath
foreach ($User in $Users) {
    New-User -FirstName $User.FirstName `
             -LastName $User.LastName `
             -Type $User.Type `
             -Title $User.Title `
             -Fullname $User.Fullname `
             -EmployeeID $User.EmployeeID `
             -LOCATION $User.LOCATION
}

Write-Host "User creation process completed."

Remove-PSSession $Session
