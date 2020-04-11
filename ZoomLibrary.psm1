#region Enums
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
#endregion Enums

#region MeetingClasses

#endregion MeetingClasses

#region ReportClasses
Class ZoomUsageReportDay {
    [System.Datetime]$date
    [System.Int32]$newUsers
    [System.Int32]$meetings
    [System.Int32]$participants
    [System.Int32]$meetingMinutes

    ZoomUsageReportDay($ZoomUsageReportDayObject) {
        $this.date = [System.DateTime]::Parse($ZoomUsageReportDayObject.date)
        $this.newUsers = $ZoomUsageReportDayObject.new_users
        $this.meetings = $ZoomUsageReportDayObject.meetings
        $this.participants = $ZoomUsageReportDayObject.participants
        $this.meetingMinutes = $ZoomUsageReportDayObject.meeting_minutes
    }
}

Class ZoomMeetingParticipant {
    [System.String]$id
    [System.String]$userId
    [System.String]$name
    [System.String]$email
    [System.DateTime]$joinTime
    [System.DateTime]$leaveTime
    [System.Int32]$duration

    ZoomMeetingParticipant($ZoomMeetingParticipantObject, [System.Boolean]$resolveNames) {
        $this.id = $ZoomMeetingParticipantObject.id
        $this.userId = $ZoomMeetingParticipantObject.user_id
        $this.joinTime = $ZoomMeetingParticipantObject.join_time
        $this.leaveTime = $ZoomMeetingParticipantObject.leave_time
        $this.duration = $ZoomMeetingParticipantObject.duration
        if ($resolveNames) {
            $this.ResolveName()
        }
    }

    ResolveName() {
        if ($null -ne $this.id -or $this.id -eq [System.String]::Empty) {
            $thisUser = Get-ZoomUser -email $this.id
            $this.name = "$($thisUser.firstName) $($thisUser.lastname)"
            $this.email = $thisUser.email
        }
    }
}

Class ZoomMeetingInstance {
    [System.String]$uuid
    [System.String]$id
    [System.String]$hostId
    [System.Int16]$meetingType
    [System.String]$topic
    [System.String]$hostName
    [System.String]$hostEmail
    [System.Datetime]$startTime
    [System.Datetime]$endTime
    [System.Timespan]$duration
    [System.Int32]$meetingMinutes
    [System.Int32]$participantCount

    ZoomMeetingInstance($ZoomMeetingInstanceObject) {
        $this.uuid = $ZoomMeetingInstanceObject.uuid
        $this.id = $ZoomMeetingInstanceObject.id
        $this.hostId = $ZoomMeetingInstanceObject.host_id
        $this.meetingType = $ZoomMeetingInstanceObject.type
        $this.topic = $ZoomMeetingInstanceObject.topic
        $this.hostName = $ZoomMeetingInstanceObject.user_name
        $this.hostEmail = $ZoomMeetingInstanceObject.user_email
        $this.startTime = $ZoomMeetingInstanceObject.start_time
        $this.endTime = $ZoomMeetingInstanceObject.end_time
        $this.duration = [System.TimeSpan]::FromMinutes($ZoomMeetingInstanceObject.duration)
        $this.meetingMinutes = $ZoomMeetingInstanceObject.total_minutes
        $this.participantCount = $ZoomMeetingInstanceObject.participants_count
    }
}

Class ZoomUsageReport {
    [System.Int16]$year
    [System.Int16]$month
    [ZoomUsageReportDay[]]$dates

    ZoomUsageReport([System.Int16]$year, [System.Int16]$month) {
        if ($null -ne $month -and $null -ne $year) {
            $uri = "https://api.zoom.us/v2/report/daily?year=$year&month=$month"
        } else {
            $uri = "https://api.zoom.us/v2/report/daily"
        }
        $report = Invoke-RestMethod -Uri $uri -Headers $global:headers
        $this.year = $report.year
        $this.month = $report.month
        $report.dates | ForEach-Object {
            $thisDate = [ZoomUsageReportDay]::new($_)
            $this.Dates += $thisDate
        }
    }
}

Class ZoomMeetingParticipantReport {
    [System.Int32]$totalRecords
    [ZoomMeetingParticipant[]]$participants

    ZoomMeetingParticipantReport([System.String]$meetingID, [System.Boolean]$resolveNames) {
        $report = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/meetings/$meetingID/participants?page-size=300" -Headers $global:headers
        $report.participants | ForEach-Object {
            Write-Debug "[ZoomMeetingParticipantReport]::new - Adding particpant with user_id: $($_.user_id)"
            $participant = [ZoomMeetingParticipant]::new($_, $resolveNames)
            $this.participants += $participant
        }
        $this.totalRecords = $report.total_records
    }
}

Class ZoomMeetingsReport {
    [ZoomMeetingInstance[]]$meetings

    ZoomMeetingsReport([System.String]$email, [System.DateTime]$from, [System.DateTime]$to) {
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/users/$email/meetings?from=$($from.ToString("yyyy-MM-dd"))&to=$($to.ToString("yyyy-MM-dd"))&page_size=100" -Headers $global:headers
        Write-Debug "[ZoomMeetingsReport]::new - $($page1.total_records) meetings in $($page1.page_count) pages. Adding page 1 containing $($page1.meetings.count) records."
        $this.meetings += $page1.meetings
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/users/$email/meetings?from=$($from.ToString('yyyy-MM-dd'))&to=$($to.ToString('yyyy-MM-dd'))&page_size=100&page_number=$count" -Headers $global:headers
                Write-Debug "ZoomMeetingsReport: Adding page $count containing $($page.meetings.count) records."
                $this.meetings += $page.meetings
                Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
            }
        }
    }
}
#endregion ReportClasses

#region UserClasses
Class ZoomUser {
    [System.String]$id
    [System.string]$firstName
    [System.String]$lastName
    [System.String]$email
    [ZoomLicenseType]$licenseType
    [System.String]$rolename
    [System.Int64]$PMI
    [System.Boolean]$usePMI
    [System.String]$vanityURL
    [System.String]$personalMeetingURL
    [System.String]$timezone
    [System.String]$department
    [DateTime]$created
    [DateTime]$lastLoginTime
    [System.String]$lastClientVersion
    [System.Int32]$hostKey
    [System.String]$JID
    [System.String[]]$groupIDs
    [ZoomGroup[]]$groups
    [System.String[]]$IMgroupIDs
    [System.String]$accountID
    [System.String]$language
    [System.String]$phoneCountry
    [System.String]$phoneNumber
    [System.String]$status
    [System.String]$jobTitle
    [System.String]$location

    ZoomUser([System.String]$email) {
        $this.email = $email
        $this.Load()
    }

    Load() {
        if ( $null -eq $this.email) {
            Write-Error "[ZoomUser]::Load No email address defined."
            Break
        }
        $user = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)" -Headers $global:headers
        $this.id = $user.id
        $this.firstName = $user.first_name
        $this.lastName = $user.last_name
        $this.licenseType = $user.type
        $this.rolename = $user.role_name
        $this.PMI = $user.pmi
        $this.usePMI = $user.use_pmi
        $this.vanityURL = $user.vanity_url
        $this.personalMeetingURL = $user.personal_meeting_url
        $this.timezone = $user.timezone
        $this.department = $user.department
        $this.created = $user.created_at
        if ($null -ne $user.last_login_time) { $this.lastLoginTime = $user.last_login_time } else { $this.lastLoginTime = [System.Datetime]::MinValue }
        $this.lastClientVersion = $user.last_client_version
        $this.hostKey = $user.host_key
        $this.JID = $user.jid
        $this.groupIDs = $user.group_ids
        $this.IMgroupIDs = $user.im_group_ids
        $this.accountID = $user.account_id
        $this.language = $user.language
        $this.phoneCountry = $user.phone_country
        $this.phoneNumber = $user.phone_number
        $this.status = $user.status
        $this.jobTitle = $user.job_title
        $this.location = $user.location

        $this.groupIDs | ForEach-Object {
            [ZoomGroup]$thisGroup = [ZoomGroup]::new($_)
            $this.groups += $thisGroup
        }
    }

    Update([System.String]$email, [System.String]$firstName, [System.String]$lastName, [ZoomLicenseType]$license, [System.String]$timezone, [System.String]$jobTitle, [System.String]$company, [System.String]$location, [System.String]$phoneNumber, [System.String]$groupName) {
        $params = @{}
        if ($firstName -ne [System.String]::Empty ) { $params.Add("first_name", $firstName) }
        if ($lastName -ne [System.String]::Empty) { $params.Add("last_name", $lastName) }
        if ($null -ne $license) { $params.Add("type", $license) }
        if ($timezone -ne [System.String]::Empty) { $params.Add("timezone", $timezone) }
        if ($jobTitle -ne [System.String]::Empty) { $params.Add("job_title", $jobTitle) }
        if ($company -ne [System.String]::Empty) { $params.Add("company", $company) }
        if ($location -ne [System.String]::Empty) { $params.Add("location", $location) }
        if ($phoneNumber -ne [System.String]::Empty) { $params.Add("phone_number", $phoneNumber) }
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $global:headers -Body ($params | ConvertTo-Json) -Method PATCH
    }

    static [ZoomUser] Create([System.String]$email, [System.String]$firstName, [System.String]$lastName, [ZoomLicenseType]$license, [System.String]$timezone, [System.String]$jobTitle, [System.String]$company, [System.String]$location, [System.String]$phoneNumber, [System.String]$groupName) {
        $userInfo = @{}
        $userInfo.Add("first_name", $firstName)
        $userInfo.Add("last_name", $lastName)
        $userInfo.Add("email", $email)
        $userInfo.Add("type", $license)
        $body = @{}
        $body.Add("action", "create")
        $body.Add("user_info", $userinfo)
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users" -Headers $global:headers -Body ( $body | ConvertTo-Json) -Method POST

        [ZoomUser]$thisUser = [ZoomUser]::new($email)
        $thisUser.Update($email, $timezone, $jobTitle, $company, $location, $phoneNumber)
        $group = [ZoomGroup]::GetByName($groupName)
        $group.AddMembers(@($email))

        return $thisUser
    }
}

Class ZoomGroup {
    [System.String]$id
    [System.String]$name
    [System.Int32]$totalMembers
    [ZoomUser[]]$members

    ZoomGroup([System.String]$id) {
        $this.id = $id
        $group = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$id" -Headers $global:headers
        $this.name = $group.name
        $this.totalMembers = $group.total_members
    }

    GetMembers() {
        $this.members = @()
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members?page_size=100" -Headers $global:headers
        Write-Debug "[ZoomGroup]::GetMembers - $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.members.count) records."
        $page1.members | ForEach-Object {
            $thisUser = [ZoomUser]::new($_.email)
            $this.members += $thisUser
            Write-Debug "[ZoomGroup]::GetMembers - adding $($_.email)"
        }
        $this.members += $page1.members
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members?page_size=100&page_number=$count" -Headers $global:headers
                Write-Debug "[ZoomGroup]::GetMembers: Adding page $count containing $($page1.members.count) records."
                $page.members | ForEach-Object {
                    $thisUser = [ZoomUser]::new($_.email)
                    $this.members += $thisUser
                    Write-Debug "[ZoomGroup]::GetMembers - adding $($_.email)"
                }
                Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
            }
        }
    }

    AddMembers([System.String[]]$emails) {
        $groupMembers = @()
        $emails | ForEach-Object {
            $member = New-Object -TypeName psobject
            $member | Add-Member -Name email -Value $_ -MemberType NoteProperty
            $groupMembers += $member
        }
        $body = New-Object -TypeName psobject
        $body | Add-Member -Name members -Value $groupMembers -MemberType NoteProperty
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members" -Headers $global:headers -Body ($body | ConvertTo-Json) -Method POST
    }

    RemoveMember([System.String]$email) {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members/$email" -Headers $global:headers -Method DELETE
    }

    static [ZoomGroup] GetByName([System.String]$name) {
        $groups = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers $global:headers
        Write-Debug "[ZoomGroup]::GetByName - Retrieved $($groups.total_records) groups."
        $thisGroup = $groups.groups | Where-object { $_.name -eq $groupName }
        return [ZoomGroup]::new($thisGroup.id)
    }
}
#endregion UserClasses

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
    return [ZoomUser]::new($email)
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
    $check = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/email?email=$email" -Headers $global:headers
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
    return [ZoomGroup]::GetByName($groupName)
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
    if ($null -ne $group) {
        Write-Error "Add-ZoomGroup: Group name ""$groupName"" already exists."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers $global:headers -Body (@{"name" = $groupName} | ConvertTo-Json) -Method POST
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
    if ($null -eq $group) {
        Write-Error "Remove-ZoomGroup: Group name ""$groupName"" could not be found."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($group.id)" -Headers $global:headers -Method DELETE
}

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
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100" -Headers $global:headers
    Write-Debug "Get-ZoomUsers: $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.users.count) records."
    $zoomusers += $page1.users
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100&page_number=$count" -Headers $global:headers
            Write-Debug "Get-ZoomUsers: Adding page $count containing $($page.users.count) records."
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
    [ZoomGroup]$group = Get-ZoomGroup -groupName $groupName
    if ($null -eq $group) {
        Write-Error "Get-ZoomGroupUsers: Group name ""$groupName"" could not be found."
        return $null
    }
    $group.GetMembers()
    return $group.members
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
    [ZoomGroup]$group = Get-ZoomGroup -groupName $groupName
    if ($null -eq $group) {
        Write-Error "Add-ZoomUsersToGroup: Group name ""$groupName"" could not be found."
        return $null
    }
    $group.AddMembers($emails)
}

<#
    .Synopsis
    Removes Zoom user from a group

    .Description
    Removes a Zoom user from a group, based upon group name and user's email address

    .Parameter Emails
    The email addresses of the user

    .Parameter GroupName
    The name of the group

    .Example
    Remove-ZoomUserFromGroup -groupName "Marketing Execs" -emails "percy@cactus.email"
#>
function Remove-ZoomUserFromGroup() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$groupName)
    [ZoomGroup]$group = Get-ZoomGroup -groupName $groupName
    if ($null -eq $group) {
        Write-Error "Remove-ZoomUserFromGroup: Group name ""$groupName"" could not be found."
        return $null
    }
    $group.RemoveMember($email)
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
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $global:headers -Body (@{"type" = $license} | ConvertTo-Json) -Method PATCH
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
    if ($null -ne $license) { $params.Add("type", $license) }
    if ($timezone -ne [System.String]::Empty) { $params.Add("timezone", $timezone) }
    if ($jobTitle -ne [System.String]::Empty) { $params.Add("job_title", $jobTitle) }
    if ($company -ne [System.String]::Empty) { $params.Add("company", $company) }
    if ($location -ne [System.String]::Empty) { $params.Add("location", $location) }
    if ($phoneNumber -ne [System.String]::Empty) { $params.Add("phone_number", $phoneNumber) }
    
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email" -Headers $global:headers -Body ($params | ConvertTo-Json) -Method PATCH
}

<#
    .Synopsis
    Sets the user's password

    .Description
    Sets the password of the Zoom user

    .Parameter Email
    The email address of the user

    .Parameter Password
    The password to set

    .Example
    Set-ZoomUserPassword -email "jerome@cactus.email" -password "Marketing101!"
#>
function Set-ZoomUserPassword() {
    #' Performs no validation on password length and complexity
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$password)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Set-ZoomUserPassword: User ""$email"" could not be found."
        return $null
    }

    Try {
        Invoke-WebRequest -Uri "https://api.zoom.us/v2/users/$email/password" -Headers $global:headers -Body (@{"password" = $password} | ConvertTo-Json) -Method PUT -ErrorAction Stop
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
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/status" -Headers $global:headers -Body (@{"action" = $action} | ConvertTo-Json) -Method PUT
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

    .Parameter GroupName
    The name of the group

    .Example
    Add-ZoomUser -email "francisco@cactus.email" -firstName Francisco -lastName Hunter -license Basic timezone "Europe/London" -jobTitle "Gra[hic Artist" -company "Cactus Industries (Europe)" -Location Basildon -phoneNumber "+44 1268 533333" -groupName "Designers"
#>
function Add-ZoomUser() {
    Param(
        [Parameter(Mandatory=$true)][System.String]$email,
        [Parameter(Mandatory=$true)][System.String]$firstName,
        [Parameter(Mandatory=$true)][System.String]$lastName,
        [Parameter(Mandatory=$true)][ZoomLicenseType]$license,
        [System.String]$timezone,
        [System.String]$jobTitle,
        [System.String]$company,
        [System.String]$location,
        [System.String]$phoneNumber,
        [System.String]$groupName
        )
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
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users" -Headers $global:headers -Body ( $body | ConvertTo-Json) -Method POST
    Set-ZoomUserDetails -email $email -timezone $timezone -jobTitle $jobTitle -company $company -location $location -phoneNumber $phoneNumber
    Add-ZoomUsersToGroup -emails $email -groupName $groupName
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

    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($email)?action=delete" -Headers $global:headers -Method DELETE
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

    $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($email)/assistants" -Headers $global:headers -Method GET
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
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/assistants" -Headers $global:headers -Body ($body | ConvertTo-Json) -Method POST
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
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/assistants/$assistant" -Headers $global:headers -Method DELETE
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/schedulers/$assistant" -Headers $global:headers -Method DELETE
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
    $global:headers.Add('Authorization',"Bearer $authtoken")
}

<#
    .Synopsis
    Retrieves a user's meetings

    .Description
    Returns a list of Zoom meetings for the specified user

    .Parameter Email
    The email address of the user.

    .Parameter MeetingType
    The type of meeting to retrieve. Available options are: Live, Upcoming and Scheduled

    .Example
    Get-ZoomMeetings -email "jerome@cactus.email" -meetingType Upcoming
#>
function Get-ZoomMeetings() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [ZoomMeetingType]$meetingType)
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Get-ZoomMeetings: User ""$email"" could not be found."
        return $null
    }
    if ($null -eq $meetingType) { $meetingType = [ZoomMeetingType]::Live }
    $zoommeetings = @()
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/meetings?page_size=100&type=$($meetingType.ToString().ToLower())" -Headers $global:headers
    Write-Debug "Get-ZoomMeetings: $($page1.total_records) meetings in $($page1.page_count) pages. Adding page 1 containing $($page1.meetings.count) records."
    $zoommeetings += $page1.meetings
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/meetings?page_size=100&type=$($meetingType.ToString().ToLower())&page_number=$page" -Headers $global:headers
            Write-Debug "Get-ZoomMeetings: Adding page $count containing $($page.meetings.count) records."
            $zoommeetings += $page.meetings
            Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
        }
    }
    return $zoommeetings
}


<#
    .Synopsis
    Retrieves the details of a meeting

    .Description
    Retrieves the details of a Zoom meeting based on meeting ID

    .Parameter MeetingID
    The numeric meeting ID

    .Example
    Get-ZoomMeetingDetails -meetingID 123456789
#>
function Get-ZoomMeetingDetails() {
    Param([Parameter(Mandatory=$true)][System.Int32]$meetingID)
    $meeting = Invoke-RestMethod -Uri "https://api.zoom.us/v2/meetings/$meetingID" -Headers $global:headers
    return $meeting
}

<#
    .Synopsis
    Terminates a meeting

    .Description
    Terminates a live Zoom meeting based on meeting ID

    .Parameter MeetingID
    The numeric meeting ID

    .Example
    Stop-ZoomMeeting -meetingID 123456789
#>
function Stop-ZoomMeeting() {
    Param([Parameter(Mandatory=$true)][System.Int32]$meetingID)
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/meetings/$meetingID/status" -Headers $global:headers -Body ( @{"action" = "end"} | ConvertTo-Json) -Method PUT
}

<#
    .Synopsis
    Lists Zoom roles

    .Description
    Returns an array of all Zoom roles

    .Example
    Get-ZoomRoles
#>
function Get-ZoomRoles() {
    $roles = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles" -Headers $global:headers -Method Get
    Write-Debug "Get-ZoomRole: Retrieved $($roles.total_records) groups."
    return $roles.roles
}

<#
    .Synopsis
    Gets a Zoom role

    .Description
    Returns a specific Zoom role based on role name

    .Parameter RoleName
    The role name

    .Example
    Get-ZoomRole -roleName "Level 2 Service Desk"
#>
function Get-ZoomRole() {
    Param([Parameter(Mandatory=$true)][System.String]$roleName)
    $roles = Get-ZoomRoles
    $thisRole = $roles | Where-object { $_.name -eq $roleName }
    return $thisRole
}

<#
    .Synopsis
    Gets Zoom users in a role

    .Description
    Returns an array of Zoom users in a role, based upon role name

    .Parameter RoleName
    The role name

    .Example
    Get-ZoomRoleUsers -roleName "Level 2 Service Desk"
#>
function Get-ZoomRoleUsers() {
    Param([Parameter(Mandatory=$true)][System.String]$roleName)
    $role = Get-ZoomRole -roleName $roleName
    if ($null -eq $role) {
        Write-Error "Get-ZoomRoleUsers: Role ""$role"" could not be found."
        return $null
    }
    $roleusers = @()
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members?page_size=100" -Headers $global:headers -Method Get
    Write-Debug "Get-ZoomRoleUsers: $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.members.count) records."
    $roleusers += $page1.members
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members?page_size=100&page_number=$count" -Headers $global:headers
            Write-Debug "Get-ZoomRoleUsers: Adding page $count containing $($page.members.count) records."
            $roleusers += $page.members
            Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
        }
    }
    return $roleusers
}

<#
    .Synopsis
    Sets a Zoom user's role

    .Description
    Sets a Zoom user's role based upon email address and role name

    .Parameter Email
    The email addresses of the users to add as an array of strings. Will also accept a string containing a single address.
    Currently, no more than 30 users can be added to a role in a single call.

    .Parameter RoleName
    The role name

    .Example
    Set-ZoomRoleUsers -email "melissa@cactus.email" -roleName "Level 2 Service Desk"
#>
function Set-ZoomRoleUsers() {
    Param([Parameter(Mandatory=$true)]$emails, [Parameter(Mandatory=$true)][System.String]$roleName)
    if ($rolename.ToLower() -eq "owner" -or $roleName.ToLower() -eq "member") {
        Write-Error "Set-ZoomRoleUsers: Users cannot be added to the Member or Owner roles."
        return $null
    }
    $role = Get-ZoomRole -roleName $roleName
    if ($null -eq $role) {
        Write-Error "Set-ZoomRoleUsers: Role ""$role"" could not be found."
        return $null
    }
    $members = @()
    $emails | ForEach-Object {
        if ( (Get-ZoomUserExists -email $_) -eq $false ) {
            Write-Error "Set-ZoomRoleUsers: User ""$_"" could not be found."
            return $null
        }
        $member = New-Object -TypeName psobject
        $member | Add-Member -Name email -Value $_ -MemberType NoteProperty
        $members += $member
    }
    $body = New-Object -TypeName psobject
    $body | Add-Member -Name members -Value $members -MemberType NoteProperty
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members" -Headers $global:headers -Body ($body | ConvertTo-Json) -Method Post
}

<#
    .Synopsis
    Removes a Zoom user from a role

    .Description
    Removes a Zoom user from a role based upon email address and role name

    .Parameter Email
    The email address of the user to remove from the role

    .Parameter RoleName
    The role name

    .Example
    Remove-ZoomRoleUser -email "melissa@cactus.email" -roleName "Level 2 Service Desk"
#>
function Remove-ZoomRoleUser() {
    Param([Parameter(Mandatory=$true)][System.String]$email, [Parameter(Mandatory=$true)][System.String]$roleName)
    $role = Get-ZoomRole -roleName $roleName
    if ($null -eq $role) {
        Write-Error "Remove-ZoomRoleUser: Role ""$role"" could not be found."
        return $null
    }
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Remove-ZoomRoleUser: User ""$email"" could not be found."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members/$email" -Headers $global:headers -Method Delete
}

<#
    .Synopsis
    Generates the Zoom daily usage report

    .Description
    Generates the Zoom daily usage report for any given month. Reports the daily number of new users, meerings, participants and meeting minutes.

    .Parameter Year
    The year of the date to report on, in yyyy format.

    .Parameter Month
    The month of the date to report on.

    .Example
    Get-ZoomUsageReport -year 2020 -month 01
#>
function Get-ZoomUsageReport() {
    Param([System.Int32]$month, [System.Int32]$year)
    $report = [ZoomUsageReport]::new($year, $month)
    return $report
}

<#
    .Synopsis
    Generates the Zoom user meetings report

    .Description
    Generates the Zoom meetings report for a specific user in a defined timespan. Note: the timespan defined by the -from and -to parameters cannot be more than one month in duration.

    .Parameter Email
    The email address of the user

    .Parameter From
    The start date of the report, expressed as a datetime object.
    .Parameter From
    The end date of the report, expressed as a datetime object.

    .Example
    Get-ZoomMeetingReport -email "jerome@cactus.email" -from "2020-01-01" -to "2020-02-01"
#>
function Get-ZoomMeetingReport() {
    Param([Parameter(Mandatory=$true)][System.String]$email,
    [Parameter(Mandatory=$true)][System.DateTime]$from,
    [Parameter(Mandatory=$true)][System.DateTime]$to)

    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Write-Error "Get-ZoomMeetingReport: User ""$email"" could not be found."
        return $null
    }
    [System.Timespan]$timeDifference = $to - $from
    if ($timeDifference.TotalDays -gt 31) {
        Write-Error "Get-ZoomMeetingReport: The date range defined by the -from and -to parameters cannot be more than one month apart."
        return $null
    }
    $zoommeetings = [ZoomMeetingsReport]::new($email, $from, $to)
    return $zoommeetings
}

<#
    .Synopsis
    Generates the Zoom meeting participant report

    .Description
    Generates the Zoom participant report for a specific meeting

    .Parameter MeetingID
    The ID of the meeting. Can be the meeting ID or the meeting instance UUID. If the meeting ID is passed, it will show details of the last instance of that meeting.

    .Parameter ResolveNames
    The ResolveNames switch attempts to retrieve each participant's name and email. THis can only be completed for participants that are part of your account.

    .Example
    Get-ZoomMeetingParticipantReport -meetingID abcefghh123456789== -resolveNames
#>
function Get-ZoomMeetingParticipantReport() {
    Param([Parameter(Mandatory=$true)][System.String]$meetingID, [switch]$resolveNames)
    $report = [ZoomMeetingParticipantReport]::new($meetingID, $resolveNames.IsPresent)
    return $report
}
#`------------------------------------------------------------------------------------------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$global:headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$global:headers.Add('Content-Type','application/json')

Write-Debug "ZoomLibrary module loaded"

#`------------------------------------------------------------------------------------------------------------------------------------------------------------

Export-ModuleMember -Function Set-ZoomAuthToken
Export-ModuleMember -Function Get-ZoomUser -Alias gzu
Export-ModuleMember -Function Get-ZoomUserExists -Alias gzue
Export-ModuleMember -Function Get-ZoomGroup -Alias gzg
Export-ModuleMember -Function Add-ZoomGroup -Alias azg
Export-ModuleMember -Function Remove-ZoomGroup -Alias rzg
Export-ModuleMember -Function Get-ZoomUsers -Alias gzusers
Export-ModuleMember -Function Get-ZoomGroupUsers -Alias gzgu
Export-ModuleMember -Function Add-ZoomUsersToGroup -Alias azug
Export-ModuleMember -Function Set-ZoomUserLicenseState -Alias szul
Export-ModuleMember -Function Set-ZoomUserDetails -Alias szud
Export-ModuleMember -Function Set-ZoomUserPassword -Alias szup
Export-ModuleMember -Function Set-ZoomUserStatus -Alias szus
Export-ModuleMember -Function Add-ZoomUser -Alias azu
Export-ModuleMember -Function Remove-ZoomUser -Alias rzu
Export-ModuleMember -Function Add-ZoomUserAssistant -Alias azua
Export-ModuleMember -Function Remove-ZoomUserAssistant -Alias rzua
Export-ModuleMember -Function Get-ZoomMeetings -Alias gzm
Export-ModuleMember -Function Get-ZoomMeetingDetails -Alias gzmd
Export-ModuleMember -Function Stop-ZoomMeeting -Alias szm
Export-ModuleMember -Function Get-ZoomRoles -Alias gzrs
Export-ModuleMember -Function Get-ZoomRole -Alias gzr
Export-ModuleMember -Function Get-ZoomRoleUsers -Alias gzru
Export-ModuleMember -Function Set-ZoomRoleUsers -Alias szru
Export-ModuleMember -Function Remove-ZoomRoleUser -Alias rzru
Export-ModuleMember -Function Get-ZoomUsageReport -Alias gzur
Export-ModuleMember -Function Get-ZoomMeetingReport -Alias gzmr
Export-ModuleMember -Function Get-ZoomMeetingParticipantReport -Alias gzmpr