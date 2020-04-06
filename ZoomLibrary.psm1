Enum ZoomLicenseType {
    Basic = 1
    Licensed = 2
    OnPrem = 3
}

Enum ZoomMeetingType {
    Scheduled = 1
    Live = 2
    Upcoming = 3
}

<#
    .Synopsis
    Returns a Zoom user object

    .Description
    Returns a Zoom user object based on email address

    .Parameter Email
    The email address of the user

    .Example
    Get-ZoomUser -email foo@cactus.email
#>
function Get-ZoomUser() {
    Param([Parameter(Mandatory=$true)][System.String]$email)
    $user = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $headers
    return $user
}

<#
    .Synopsis
    Checks if a user is present in the account

    .Description
    Performs a query to determine if a particular user is present in the Zoom account based on email address

    .Parameter Email
    The email address of the user

    .Example
    Get-ZoomUserExists -email foo@cactus.email
#>
function Get-ZoomUserExists() {
    Param([Parameter(Mandatory=$true)][System.String]$email)
    $check = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/email?email=$email" -Headers $headers
    return $check.existed_email
}

<#
    .Synopsis
    Returns a Zoom group object

    .Description
    Returns a Zoom group object based upon group name

    .Parameter GroupName
    The name of the group

    .Example
    Get-ZoomGroup -groupName "Marketing Execs"
#>
function Get-ZoomGroup() {
    Param([Parameter(Mandatory=$true)][System.String]$groupName)
    $groups = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers $headers
    Write-Debug "Get-ZoomGroup: Retrieved $($groups.total_records) groups."
    $thisGroup = $groups.groups | Where-object { $_.name -eq $groupName }
    return $thisGroup
}

<#
    .Synopsis
    Adds a new Zoom group

    .Description
    Creates and returns a new Zoom group

    .Parameter GroupName
    The name of the group

    .Example
    Add-ZoomGroup -groupName "Marketing Execs"
#>
function Add-ZoomGroup() {
    Param([Parameter(Mandatory=$true)][System.String]$groupName)
    $group = Get-ZoomGroup -groupName $groupName
    if ($group -ne $null) {
        Write-Error "Add-ZoomGroup: Group name ""$groupName"" already exists."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers $headers -Body (@{"name" = $groupName} | ConvertTo-Json) -Method POST
    $thisGroup = Get-ZoomGroup -groupName $groupName
    return $thisGroup
}

<#
    .Synopsis
    Removes a Zoom group

    .Description
    Removes a Zoom group

    .Parameter GroupName
    The name of the group

    .Example
    Remove-ZoomGroup -groupName "Marketing Execs"
#>
function Remove-ZoomGroup() {
    Param([Parameter(Mandatory=$true)][System.String]$groupName)
    $group = Get-ZoomGroup -groupName $groupName
    if ($group -eq $null) {
        Write-Error "Remove-ZoomGroup: Group name ""$groupName"" could not be found."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($group.id)" -Headers $headers -Method DELETE
}

<#
    .Synopsis
    Returns all Zoom users in the account

    .Description
    Returns an array of all Zoom users in the account.

    .Example
    Get-ZoomUsers
#>

<#
    .Synopsis
    Returns all Zoom users in the account

    .Description
    Returns an array of all Zoom users in the account.

    .Example
    Get-ZoomUsers
#>
function Get-ZoomUsers() {
    $zoomusers = @()
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100" -Headers $headers
    Write-Debug "Get-ZoomUsers: $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.users.count) records."
    $zoomusers += $page1.users
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100&page_number=$count" -Headers $headers
            Write-Debug "Get-ZoomUsers: Adding page $count containing $($page1.users.count) records."
            $zoomusers += $page.users
            Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
        }
    }
    return $zoomusers
}

<#
    .Synopsis
    Returns all Zoom users in a group

    .Description
    Returns an array of all Zoom users in a particular group, based upon group name.

    .Parameter GroupName
    The name of the group

    .Example
    Get-ZoomGroupUsers -groupName "Marketing Execs"
#>
function Get-ZoomGroupUsers() {
    Param([Parameter(Mandatory=$true)][System.String]$groupName)
    $group = Get-ZoomGroup -groupName $groupName
    if ($group -eq $null) {
        Write-Error "Get-ZoomGroupUsers: Group name ""$groupName"" could not be found."
        return $null
    }
    Write-Debug "Get-ZoomGroupUsers: Group name ""$groupName"" resolved to group ID: $($group.id)"
    $zoomgroupusers = @()
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($group.id)/members?page_size=100" -Headers $headers
    Write-Debug "Get-ZoomGroupUsers: $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.members.count) records."
    $zoomgroupusers += $page1.members
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($group.id)/members?page_size=100&page_number=$count" -Headers $headers
            Write-Debug "Get-ZoomGroupUsers: Adding page $count containing $($page1.members.count) records."
            $zoomgroupusers += $page.members
            Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
        }
    }
    return $zoomgroupusers
}

<#
    .Synopsis
    Adds Zoom users to a group

    .Description
    Adds Zoom users into a group, based upon group name and user's email address

    .Parameter Emails
    The email addresses of the users to add as an array of strings. Will also accept a string containing a single address.

    .Parameter GroupName
    The name of the group

    .Example
    Add-ZoomUsersToGroup -groupName "Marketing Execs" -emails @("jerome@cactus.email","percy@cactus.email")
#>
function Add-ZoomUsersToGroup() {
    Param([Parameter(Mandatory=$true)]$emails, [Parameter(Mandatory=$true)][System.String]$groupName)
    $group = Get-ZoomGroup -groupName $groupName
    if ($group -eq $null) {
        Write-Error "Add-ZoomUsersToGroup: Group name ""$groupName"" could not be found."
        return $null
    }
    $members = @()
    $emails | ForEach-Object {
        $member = New-Object -TypeName psobject
        $member | Add-Member -Name email -Value $_ -MemberType NoteProperty
        $members += $member
    }
    $body = New-Object -TypeName psobject
    $body | Add-Member -Name members -Value $members -MemberType NoteProperty
    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($group.id)/members" -Headers $headers -Body ($body | ConvertTo-Json) -Method POST

    return $result
}

<#
    .Synopsis
    Sets the licensing state of a user

    .Description
    Sets the licensing state of a Zoom user, based on email address

    .Parameter Email
    The email address of the user to modify.

    .Parameter License
    The type of license to apply. Valid values are: [ZoomLicenseType]::Basic, [ZoomLicenseType]::Licensed and [ZoomLicenseType]::OnPrem

    .Example
    Set-ZoomUserLicenseState -email foo@cactus.email -license Licensed
#>
function Set-ZoomUserLicenseState() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][ZoomLicenseType]$license)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Set-ZoomUserLicenseState: User ""$email"" could not be found."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $headers -Body (@{"type" = $license} | ConvertTo-Json) -Method PATCH
}

<#
    .Synopsis
    Modifies a user's profile details

    .Description
    Modifies a Zoom user's name, licensed state, job title, timezone, company name, location and phone number

    .Parameter Email
    The email address of the user to modify.

    .Parameter FirstName
    The Zoom user's first name
    
    .Parameter LastName
    The Zoom user's last name

    .Parameter License
    The type of license to apply. Valid values are: [ZoomLicenseType]::Basic, [ZoomLicenseType]::Licensed and [ZoomLicenseType]::OnPrem

    .Parameter TimeZone
    The time zone the Zoom user is in. Time zone to be specific in TZD format. For more information, see: https://marketplace.zoom.us/docs/api-reference/other-references/abbreviation-lists#timezones

    .Parameter JobTitle
    The Zoom user's job title

    .Parameter Company
    The Zoom user's company

    .Parameter Location
    The Zoom user's location. This is a non-validated field and may be used to represent any location information: office name, city, country, physical address, etc.

    .Parameter PhoneNumber
    The Zoom user's phone number, to be specified with country code, e.g. +44 207 782 5000

    .Example
    Set-ZoomUserDetails -email jerome@cactus.email -firstName Jerome -lastName Ramirez -license Licensed -timezone "Europe/London" -jobTitle "Head of Channel Marketing" -company "Cactus Industries (Europe)" -Location Basildon -phoneNumber "+44 1268 533333"
#>
function Set-ZoomUserDetails() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [System.String]$firstName, [System.String]$lastName, [ZoomLicenseType]$license, [System.String]$timezone, [System.String]$jobTitle, [System.String]$company, [System.String]$location, [System.String]$phoneNumber)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Set-ZoomUserDetails: User ""$email"" could not be found."
        return $null
    }
    $params = @{}
    if ($firstName -ne [System.String]::Empty ) { $params.Add("first_name", $firstName) }
    if ($lastName -ne [System.String]::Empty) { $params.Add("last_name", $lastName) }
    if ($license -ne $null) { $params.Add("type", $license) }
    if ($timezone -ne [System.String]::Empty) { $params.Add("timezone", $timezone) }
    if ($jobTitle -ne [System.String]::Empty) { $params.Add("job_title", $jobTitle) }
    if ($company -ne [System.String]::Empty) { $params.Add("company", $company) }
    if ($location -ne [System.String]::Empty) { $params.Add("location", $location) }
    if ($phoneNumber -ne [System.String]::Empty) { $params.Add("phone_number", $phoneNumber) }
    
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $headers -Body ($params | ConvertTo-Json) -Method PATCH
}

<#
    .Synopsis
    Sets the user's password

    .Description
    Sets the password of the Zoom user

    .Parameter Email
    The email address of the user

    .Parameter Password
    A SecureString value of the password to set

    .Example
    Set-ZoomUserPassword -email "jerome@cactus.email" -password (ConvertTo-SecureString -String "Marketing101!" -Force -AsPlainText)
#>
function Set-ZoomUserPassword() {
    #' Performs no validation on password length and complexity
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][SecureString]$password)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Set-ZoomUserPassword: User ""$email"" could not be found."
        return $null
    }

    Try {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
        $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $result = Invoke-WebRequest -Uri "https://api.zoom.us/v2/users/$email/password" -Headers $headers -Body (@{"password" = $UnsecurePassword} | ConvertTo-Json) -Method PUT -ErrorAction Stop
    } Catch {
        $response = $_
        $message = ($response.ErrorDetails.Message | ConvertFrom-Json).message
        Write-Error "Set-ZoomUserPassword: Could not set password. Error: $message"
    }
}

<#
    .Synopsis
    Activates or deactivates a user

    .Description
    Sets the activation state of the Zoom user

    .Parameter Email
    The email address of the user

    .Parameter Enabled
    A boolean value; $true for active, $false for inactive

    .Example
    Set-ZoomUserStatus -email "jerome@cactus.email" -enabled $false
#>
function Set-ZoomUserStatus() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][boolean]$enabled)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Set-ZoomUserPassword: User ""$email"" could not be found."
        return $null
    }
    if ($enabled) { $action = 'activate' } else { $action = 'deactivate' } 
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/status" -Headers $headers -Body (@{"action" = $action} | ConvertTo-Json) -Method PUT
}

<#
    .Synopsis
    Adds a new user

    .Description
    Adds a new Zoom user to the account

    .Parameter Email
    The email address of the user

    .Parameter FirstName
    The first name of the user

    .Parameter LastName
    The last name of the user

    .Parameter License
    The type of license to apply. Valid values are: [ZoomLicenseType]::Basic, [ZoomLicenseType]::Licensed and [ZoomLicenseType]::OnPrem

    .Example
    Add-ZoomUser -email "francisco@cactus.email" -firstName Francisco -lastName Hunter -license Basic
#>
function Add-ZoomUser() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$firstName, [Parameter(Mandatory=$true)][System.String]$lastName, [Parameter(Mandatory=$true)][ZoomLicenseType]$license)
    if ( (Get-ZoomUserExists -email $email) -eq $true ) {
        Write-Error "Add-ZoomUser: User ""$email"" already exists."
        return $null
    }
    $userInfo = @{}
    $userInfo.Add("first_name", $firstName)
    $userInfo.Add("last_name", $lastName)
    $userInfo.Add("email", $email)
    $userInfo.Add("type", $license)
    $body = @{}
    $body.Add("action", "create")
    $body.Add("user_info", $userinfo)
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users" -Headers $headers -Body ( $body | ConvertTo-Json) -Method POST
}

<#
    .Synopsis
    Removes a user

    .Description
    Removes a Zoom user from the account

    .Parameter Email
    The email address of the user

    .Example
    Remove-ZoomUser -email "francisco@cactus.email"
#>
function Remove-ZoomUser() {
    Param([Parameter(Mandatory=$true)][System.String]$email)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Remove-ZoomUser: User ""$email"" could not be found."
        return $null
    }

    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($email)?action=delete" -Headers $headers -Method DELETE
}

<#
    .Synopsis
    Returns the assistants / delegates for a user

    .Description
    Returns an array of Zoom users configured as the assistants / delegates for another user, based on email address

    .Parameter Email
    The email address of the user

    .Example
    Get-ZoomUserAssistants -email "jerome@cactus.email"
#>
function Get-ZoomUserAssistants() {
    Param([Parameter(Mandatory=$true)][System.String]$email)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Get-ZoomUserAssistants: User ""$email"" could not be found."
        return $null
    }

    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($email)/assistants" -Headers $headers -Method GET
    return $result.assistants
}

<#
    .Synopsis
    Adds a new assistant / delegate to a user

    .Description
    Adds a new assistant / delegate to a user, based upon email addresses. Caveat: both the host and assistnt must be fully-licenses users.

    .Parameter Email
    The email address of the user to add the assistant to.

    .Parameter Assistant
    The email address of the assistant.

    .Example
    Add-ZoomUserAssistant -email "jerome@cactus.email" -assistant "francisco@cactus.email"
#>
function Add-ZoomUserAssistant() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$assistant)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Add-ZoomUserAssistants: User ""$email"" could not be found."
        return $null
    }
    if ( (Get-ZoomUserExists -email $assistant) -eq $false ) {
        Write-Error "Add-ZoomUserAssistants: Assistant ""$assistant"" could not be found."
        return $null
    }
    $body = New-Object -TypeName psobject
    $body | Add-Member -Name assistants -Value @(@{"email" = $assistant}) -MemberType NoteProperty
    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/assistants" -Headers $headers -Body ($body | ConvertTo-Json) -Method POST
}

<#
    .Synopsis
    Removes an assistant / delegate from a user

    .Description
    Removes an assistant / delegate from a user, based upon email addresses. Warning: although API returns a success code, this API call does not appear to function.

    .Parameter Email
    The email address of the user to remove the assistant from.

    .Parameter Assistant
    The email address of the assistant.

    .Example
    Remove-ZoomUserAssistant -email "jerome@cactus.email" -assistant "francisco@cactus.email"
#>
function Remove-ZoomUserAssistant() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$assistant)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Remove-ZoomUserAssistants: User ""$email"" could not be found."
        return $null
    }
    if ( (Get-ZoomUserExists -email $assistant) -eq $false ) {
        Write-Error "Remove-ZoomUserAssistants: Assistant ""$assistant"" could not be found."
        return $null
    }
    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/assistants/$assistant" -Headers $headers -Method DELETE
    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/schedulers/$assistant" -Headers $headers -Method DELETE
}

<#
    .Synopsis
    Sets the Zoom JSON Web Token

    .Description
    Sets the Zoom JSON Web Token required for access. Must be called before any other function.
    For more information, see: https://marketplace.zoom.us/docs/guides/auth/jwt 

    .Parameter Token
    The JWT defined for your application

    .Example
    Set-ZoomAuthToken -token "775914458cf8498080d96c8355774f4f75c8ab71da894900946eeeb2f25b315c"
#>
function Set-ZoomAuthToken() {
    Param([Parameter(Mandatory=$true)][System.String]$token)
    $authtoken = $token
    $headers.Add('Authorization',"Bearer $authtoken")
}

#`------------------------------------------------------------------------------------------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add('Content-Type','application/json')

Write-Debug "ZoomLibrary module loaded"

#`------------------------------------------------------------------------------------------------------------------------------------------------------------

Export-ModuleMember -Function Set-ZoomAuthToken
Export-ModuleMember -Function Get-ZoomUser
Export-ModuleMember -Function Get-ZoomUserExists
Export-ModuleMember -Function Get-ZoomGroup
Export-ModuleMember -Function Add-ZoomGroup
Export-ModuleMember -Function Remove-ZoomGroup
Export-ModuleMember -Function Get-ZoomUsers
Export-ModuleMember -Function Get-ZoomGroupUsers
Export-ModuleMember -Function Add-ZoomUsersToGroup
Export-ModuleMember -Function Set-ZoomUserLicenseState
Export-ModuleMember -Function Set-ZoomUserDetails
Export-ModuleMember -Function Set-ZoomUserPassword
Export-ModuleMember -Function Set-ZoomUserStatus
Export-ModuleMember -Function Add-ZoomUser
Export-ModuleMember -Function Remove-ZoomUser
Export-ModuleMember -Function Add-ZoomUserAssistant
Export-ModuleMember -Function Remove-ZoomUserAssistant

