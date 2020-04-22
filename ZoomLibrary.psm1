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

Enum ZoomMeetingTypeDetail {
    Instant = 1
    Scheduled = 2
    RecurringWithNoFixedTime = 3
    PMIMeeting = 4
    RecurringWithFixedTime = 8
}
#endregion Enums

#region MeetingClasses

Class ZoomMeeting {
    [System.String]$uuid
    [System.Int64]$id
    [System.String]$hostId
    [ZoomUser]$host
    [System.String]$topic
    [ZoomMeetingTypeDetail]$meetingType
    [System.String]$status
    [System.Datetime]$startTime
    [System.TimeSpan]$duration
    [System.String]$timezone
    [System.String]$agenda
    [System.Datetime]$createdAt
    [System.Uri]$startUrl
    [System.Uri]$joinUrl
    [System.String]$meetingPassword
    [System.String]$h323Password
    [System.String]$PSTNPassword
    [System.String]$encryptedPassword
    
    StopMeeting() {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/meetings/$($this.uuid)/status" -Headers (Get-ZoomAuthHeader) -Body ( @{"action" = "end"} | ConvertTo-Json) -Method PUT
    }

    static [ZoomMeeting] GetMeeting([System.String]$id) {
        write-debug "[ZoomMeeting]::GetMeeting - Retrieving meeting details for meeting ID: $id"
        $thisMeeting = [ZoomMeeting]::new()
        $meetingDetail = Invoke-RestMethod -Uri "https://api.zoom.us/v2/meetings/$id" -Headers (Get-ZoomAuthHeader) -Method Get
        $thisMeeting.uuid = $meetingDetail.uuid
        $thisMeeting.id = $meetingDetail.id
        $thisMeeting.hostId = $meetingDetail.host_id
        $thisMeeting.host = Get-ZoomUser -email $thisMeeting.hostId
        $thisMeeting.topic = $meetingDetail.topic
        $thisMeeting.meetingType = $meetingDetail.type
        $thisMeeting.status = $meetingDetail.status
        Try {
            $thisMeeting.startTime = $meetingDetail.start_time
        } Catch {
            $thisMeeting.startTime = [System.DateTime]::MinValue
        }
        if ($meetingDetail.duration) { $thisMeeting.duration = [System.TimeSpan]::FromMinutes($meetingDetail.duration) }
        $thisMeeting.timezone = $meetingDetail.timezone
        $thisMeeting.agenda = $meetingDetail.agenda
        $thisMeeting.createdAt = $meetingDetail.created_at
        $thisMeeting.startUrl = [System.Uri]::new($meetingDetail.start_url)
        $thisMeeting.joinUrl = [System.Uri]::new($meetingDetail.join_url)
        $thisMeeting.meetingPassword = $meetingDetail.password
        $thisMeeting.h323Password = $meetingDetail.h323_password
        $thisMeeting.PSTNPassword = $meetingDetail.pstn_password
        $thisMeeting.encryptedPassword = $meetingDetail.encrypted_password
        return $thisMeeting
    }

    static [ZoomMeeting] GetMeetingStub($meetingDetail) {
        write-debug "[ZoomMeeting]::GetMeetingStub - Creating meeting stub for meeting ID: $($meetingDetail.id)"
        $thisMeeting = [ZoomMeeting]::new()

        $thisMeeting.uuid = $meetingDetail.uuid
        $thisMeeting.id = $meetingDetail.id
        $thisMeeting.hostId = $meetingDetail.host_id
        $thisMeeting.topic = $meetingDetail.topic
        $thisMeeting.meetingType = $meetingDetail.type
        Try {
            $thisMeeting.startTime = $meetingDetail.start_time
        } Catch {
            $thisMeeting.startTime = [System.DateTime]::MinValue
        }
        if ($meetingDetail.duration) { $thisMeeting.duration = [System.TimeSpan]::FromMinutes($meetingDetail.duration) }
        $thisMeeting.timezone = $meetingDetail.timezone
        $thisMeeting.createdAt = $meetingDetail.created_at
        $thisMeeting.joinUrl = [System.Uri]::new($meetingDetail.join_url)
        return $thisMeeting       
    }

    static [ZoomMeeting[]] GetMeetings([System.String]$email, [ZoomMeetingType]$meetingType, [Boolean]$detailed) {
        $zoommeetings = @()
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/meetings?page_size=100&type=$($meetingType.ToString().ToLower())" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomMeeting]::GetMeetings - $($page1.total_records) meetings in $($page1.page_count) pages. Adding page 1 containing $($page1.meetings.count) records."
        $page1.meetings | ForEach-Object {
            if ($detailed) {
                $zoommeetings += [ZoomMeeting]::GetMeeting($_.id)
            } else {
                $zoommeetings += [ZoomMeeting]::GetMeetingStub($_)
            }
        }
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$email/meetings?page_size=100&type=$($meetingType.ToString().ToLower())&page_number=$count" -Headers (Get-ZoomAuthHeader)
                Write-Debug "[ZoomMeeting]::GetMeetings - Adding page $count containing $($page.meetings.count) records."
                $page.meetings | ForEach-Object {
                    if ($detailed) {
                        $zoommeetings += [ZoomMeeting]::GetMeeting($_.id)
                    } else {
                        $zoommeetings += [ZoomMeeting]::GetMeetingStub($_)
                    }
                }
                Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
            }
        }
        return $zoommeetings
    }

}

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
        $report = Invoke-RestMethod -Uri $uri -Headers (Get-ZoomAuthHeader)
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
        $report = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/meetings/$meetingID/participants?page-size=300" -Headers (Get-ZoomAuthHeader)
        $report.participants | ForEach-Object {
            Write-Debug "[ZoomMeetingParticipantReport]::new - Adding particpant with user_id: $($_.user_id)"
            $participant = [ZoomMeetingParticipant]::new($_, $resolveNames)
            $this.participants += $participant
        }
        $this.totalRecords = $report.total_records
    }
}

Class ZoomWebinarParticipantReport {
    [System.Int32]$totalRecords
    [ZoomMeetingParticipant[]]$participants

    ZoomWebinarParticipantReport([System.String]$webinarID, [System.Boolean]$resolveNames) {
        $report = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/webinars/$webinarID/participants?page-size=300" -Headers (Get-ZoomAuthHeader)
        $report.participants | ForEach-Object {
            Write-Debug "[ZoomWebinarParticipantReport]::new - Adding particpant with user_id: $($_.user_id)"
            $participant = [ZoomMeetingParticipant]::new($_, $resolveNames)
            $this.participants += $participant
        }
        $this.totalRecords = $report.total_records
    }
}

Class ZoomMeetingsReport {
    [ZoomMeetingInstance[]]$meetings

    ZoomMeetingsReport([System.String]$email, [System.DateTime]$from, [System.DateTime]$to) {
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/users/$email/meetings?from=$($from.ToString("yyyy-MM-dd"))&to=$($to.ToString("yyyy-MM-dd"))&page_size=100" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomMeetingsReport]::new - $($page1.total_records) meetings in $($page1.page_count) pages. Adding page 1 containing $($page1.meetings.count) records."
        $this.meetings += $page1.meetings
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/users/$email/meetings?from=$($from.ToString('yyyy-MM-dd'))&to=$($to.ToString('yyyy-MM-dd'))&page_size=100&page_number=$count" -Headers (Get-ZoomAuthHeader)
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
    [ZoomUser[]]$assistants
    [ZoomUser[]]$schedulers

    ZoomUser([System.String]$email) {
        $this.email = $email
    }

    ZoomUser([PSCustomObject]$user) {
        $this.id = $user.id
        $this.firstName = $user.first_name
        $this.lastName = $user.last_name
        $this.email = $user.email
        $this.licenseType = $user.type
        $this.PMI = $user.pmi
        $this.timezone = $user.timezone
        $this.department = $user.dept
        $this.created = $user.created_at
        if ($null -ne $user.last_login_time) { $this.lastLoginTime = $user.last_login_time } else { $this.lastLoginTime = [System.Datetime]::MinValue }
        $this.lastClientVersion = $user.last_client_version
        $this.groupIDs = $user.group_ids
        $this.language = $user.language
        $this.phoneNumber = $user.phone_number
        $this.status = $user.status
        $this.jobTitle = $user.job_title
        $this.location = $user.location
    }

    Load() {
        if ( $null -eq $this.email) {
            Throw "[ZoomUser]::Load No email address defined."
            Break
        }
        $user = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)" -Headers (Get-ZoomAuthHeader)
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
        $this.department = $user.dept
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

        if ($null -ne $this.groups) { $this.groups.Clear() }
        $this.groupIDs | ForEach-Object {
            [ZoomGroup]$thisGroup = [ZoomGroup]::new($_)
            $this.groups += $thisGroup
        }
    }

    Update([System.String]$firstName, [System.String]$lastName, [ZoomLicenseType]$license, [System.String]$timezone, [System.String]$jobTitle, [System.String]$company, [System.String]$location, [System.String]$phoneNumber) {
        $params = @{}
        if ($firstName -ne [System.String]::Empty ) { $params.Add("first_name", $firstName) }
        if ($lastName -ne [System.String]::Empty) { $params.Add("last_name", $lastName) }
        if ($null -ne $license) { $params.Add("type", $license) }
        if ($timezone -ne [System.String]::Empty) { $params.Add("timezone", $timezone) }
        if ($jobTitle -ne [System.String]::Empty) { $params.Add("job_title", $jobTitle) }
        if ($company -ne [System.String]::Empty) { $params.Add("company", $company) }
        if ($location -ne [System.String]::Empty) { $params.Add("location", $location) }
        if ($phoneNumber -ne [System.String]::Empty) { $params.Add("phone_number", $phoneNumber) }
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)" -Headers (Get-ZoomAuthHeader) -Body ($params | ConvertTo-Json) -Method PATCH
        $this.Load()
    }

    Save() {
        $params = @{}
        $params.Add("first_name", $this.firstName)
        $params.Add("last_name", $this.lastName)
        $params.Add("type", $this.licenseType)
        $params.Add("timezone", $this.timezone)
        $params.Add("job_title", $this.jobTitle)
        $params.Add("company", $this.company)
        $params.Add("location", $this.location)
        $params.Add("phone_number", $this.phoneNumber)
        $params.Add("department", $this.department)
        $params.Add("vanity_name", $this.vanityURL)
        $params.Add("use_pmi", $this.usePMI)
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)" -Headers (Get-ZoomAuthHeader) -Body ($params | ConvertTo-Json) -Method PATCH
    }

    SetPassword([SecureString]$password) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
        $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        Try {
            Invoke-WebRequest -Uri "https://api.zoom.us/v2/users/$($this.email)/password" -Headers (Get-ZoomAuthHeader) -Body (@{"password" = $UnsecurePassword} | ConvertTo-Json) -Method PUT -ErrorAction Stop
        } Catch {
            $response = $_
            $message = ($response.ErrorDetails.Message | ConvertFrom-Json).message
            Throw "[ZoomUser]::SetPassword - Could not set password. Error: $message"
        }
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    }

    SetStatus([System.Boolean]$enabled) {
        if ($enabled) { $action = 'activate' } else { $action = 'deactivate' } 
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/status" -Headers (Get-ZoomAuthHeader) -Body (@{"action" = $action} | ConvertTo-Json) -Method PUT
        $this.Load()
    }

    SetLicenseStatus([ZoomLicenseType]$license) {
        if ($this.licenseType -ne $license) {
            Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)" -Headers (Get-ZoomAuthHeader) -Body (@{"type" = $license} | ConvertTo-Json) -Method PATCH
            $this.Load()
        }
    }

    SetFeatureStatus([System.Boolean]$webinar, [System.Boolean]$largeMeeting, [System.Int32]$webinarCapacity, [System.Int32]$largeMeetingCapacity) {
        $params = @{}
        $features = @{}
        $features.Add("webinar", $webinar)
        $features.Add("webinar_capacity", $webinarCapacity)
        $features.Add("large_meeting", $largeMeeting)
        $features.Add("large_meeting_capacity", $largeMeetingCapacity)
        $params.Add("feature", $features)
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/settings" -Headers (Get-ZoomAuthHeader) -Body ($params | ConvertTo-Json) -Method PATCH
        $this.Load()
    }

    Delete() {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)?action=delete" -Headers (Get-ZoomAuthHeader) -Method DELETE
    }

     GetAssistants() {
        $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/assistants" -Headers (Get-ZoomAuthHeader) -Method GET
        if ($null -ne $this.assistants) { [ZoomUser[]]$this.assistants = @() }
        $result.assistants | ForEach-Object {
            $thisAssistant = [ZoomUser]::GetUserStub($_.id, $_.email)
            $this.assistants += $thisAssistant
        }
    }

    GetSchedulers() {
        $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/schedulers" -Headers (Get-ZoomAuthHeader) -Method GET
        if ($null -ne $this.schedulers) { [ZoomUser[]]$this.schedulers = @() }
        $result.schedulers | ForEach-Object {
            $thisScheduler = [ZoomUser]::GetUserStub($_.id, $_.email)
            $this.schedulers += $thisScheduler
        }
    }

    AddAssistant([System.String]$assistant) {
        $body = New-Object -TypeName psobject
        $body | Add-Member -Name assistants -Value @(@{"email" = $assistant}) -MemberType NoteProperty
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/assistants" -Headers (Get-ZoomAuthHeader) -Body ($body | ConvertTo-Json) -Method POST
    }

    RemoveAssistant([System.String]$assistant) {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/assistants/$assistant" -Headers (Get-ZoomAuthHeader) -Method DELETE
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/schedulers/$assistant" -Headers (Get-ZoomAuthHeader) -Method DELETE
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
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users" -Headers (Get-ZoomAuthHeader) -Body ( $body | ConvertTo-Json) -Method POST

        [ZoomUser]$thisUser = [ZoomUser]::new($email)
        if ($timezone -or $jobTitle -or $company -or $location -or $phoneNumber) {
            $thisUser.Update($firstName, $lastName, $license, $timezone, $jobTitle, $company, $location, $phoneNumber)
        }
        if ($groupName) {
            $group = [ZoomGroup]::GetByName($groupName)
            $group.AddMembers(@($email))
        }
        return $thisUser
    }

    static [ZoomUser] GetUserStub([System.String]$id, [System.String]$email) {
        [ZoomUser]$thisUser = [ZoomUser]::new($email)
        $thisUser.id = $id
        return $thisUser
    }

    static [ZoomUser] GetUserStub([System.String]$id, [System.String]$email, [System.String]$firstName, [System.String]$lastName, [ZoomLicenseType]$license) {
        [ZoomUser]$thisUser = [ZoomUser]::new($email)
        $thisUser.id = $id
        $thisUser.firstName = $firstName
        $thisUser.lastName = $lastName
        $thisUser.licenseType = $license
        return $thisUser
    }

    static [ZoomUser] GetUserDetails([System.String]$email) {
        [ZoomUser]$thisUser = [ZoomUser]::new($email)
        $thisUser.Load()
        return $thisUser
    }

    static [System.Boolean] CheckExists([System.String]$email) {
        $check = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/email?email=$email" -Headers (Get-ZoomAuthHeader)
        return $check.existed_email
    }

    static [ZoomUser[]] GetUsers() {
        $zoomusers = @()
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomUser]::GetUsers - $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.users.count) records."
        $page1.users | ForEach-Object {
            $thisUser = [ZoomUser]::new($_)
            $zoomusers += $thisUser
        }
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/users?status=active&page_size=100&page_number=$count" -Headers (Get-ZoomAuthHeader)
                Write-Debug "[ZoomUser]::GetUsers: Adding page $count containing $($page1.users.count) records."
                $page.users | ForEach-Object {
                    $thisUser = [ZoomUser]::new($_)
                    $zoomusers  += $thisUser
                }
                Start-Sleep -Milliseconds 100 #` To keep under Zoom's API rate limit of 10 per second.
            }
        }
        return $zoomusers
    }
}

Class ZoomGroup {
    [System.String]$id
    [System.String]$name
    [System.Int32]$totalMembers
    [ZoomUser[]]$members

    ZoomGroup() {
    }

    ZoomGroup([System.String]$id) {
        $this.id = $id
        $group = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$id" -Headers (Get-ZoomAuthHeader)
        $this.name = $group.name
        $this.totalMembers = $group.total_members
    }

    GetMembers() {
        $this.members = @()
        $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members?page_size=100" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomGroup]::GetMembers - $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.members.count) records."
        $page1.members | ForEach-Object {
            $thisUser = [ZoomUser]::GetUserStub($_.id, $_.email, $_.first_name, $_.last_name, $_.type)
            $this.members += $thisUser
        }
        if ($page1.page_count -gt 1) {
            for ($count=2; $count -le $page1.page_count; $count++) {
                $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members?page_size=100&page_number=$count" -Headers (Get-ZoomAuthHeader)
                Write-Debug "[ZoomGroup]::GetMembers: Adding page $count containing $($page1.members.count) records."
                $page.members | ForEach-Object {
                    $thisUser = $thisUser = [ZoomUser]::GetUserStub($_.id, $_.email, $_.first_name, $_.last_name, $_.type)
                    $this.members += $thisUser
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
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members" -Headers (Get-ZoomAuthHeader) -Body ($body | ConvertTo-Json) -Method POST
    }

    RemoveMember([ZoomUser]$user) {        
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)/members/$($user.id)" -Headers (Get-ZoomAuthHeader) -Method DELETE
    }

    Delete() {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups/$($this.id)" -Headers (Get-ZoomAuthHeader) -Method DELETE
    }

    static [ZoomGroup[]] GetAllGroups() {
        $zoomgroups = @()
        $groups = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomGroup]::GetAllGroups - Retrieved $($groups.total_records) groups."
        $groups.groups | ForEach-Object {
            $thisGroup = [ZoomGroup]::new()
            $thisGroup.id = $_.id
            $thisGroup.name = $_.name
            $thisGroup.totalMembers = $_.total_members
            $zoomgroups += $thisGroup
        }
        return $zoomgroups
    }

    static [ZoomGroup] GetByName([System.String]$groupName) {
        $groups = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers (Get-ZoomAuthHeader)
        Write-Debug "[ZoomGroup]::GetByName - Retrieved $($groups.total_records) groups."
        $thisGroup = $groups.groups | Where-object { $_.name -eq $groupName }
        if ($null -ne $thisGroup) {
            return [ZoomGroup]::new($thisGroup.id)
        } else {
            return $null
        }
    }

    static [ZoomGroup] Create([System.String]$groupName) {
        $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/groups" -Headers (Get-ZoomAuthHeader) -Body (@{"name" = $groupName} | ConvertTo-Json) -Method POST
        $thisGroup = [ZoomGroup]::new($result.id)
        return $thisGroup
    }
}
#endregion UserClasses

<#
    .Synopsis
    Returns a Zoom user object

    .Description
    Returns a Zoom user object based on email address
    Note: to reduce the number of API calls required, a zoom user's assistants and schedulers are not populated automatically. To populate, call [ZoomUser].GetAssistants or [ZoomUser].GetSchedulers as required.

    .Parameter Email
    The email address of the user

    .Example
    Get-ZoomUser -email foo@cactus.email
#>
function Get-ZoomUser() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.String]$email
    )
    return [ZoomUser]::GetUserDetails($email)
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
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.String]$email
    )
    return [ZoomUser]::CheckExists($email)
}

<#
    .Synopsis
    Returns details of a single Zoom group or all groups on the account

    .Description
    Returns a [ZoomGroup] object if the -groupName parameter is specified. If the -groupName parameter is omitted, returns an array of [ZoomGroup] objects corresponding to all of the Zoom groups on the account.

    .Parameter GroupName
    The name of the group

    .Example
    Get-ZoomGroup -groupName "Marketing Execs"
#>
function Get-ZoomGroup() {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0, ValueFromPipeline=$true)]
        [System.String]$groupName
    )
    if ([System.String]::IsNullOrEmpty($groupName)) {
        return [ZoomGroup]::GetAllGroups()
    } else {
        return [ZoomGroup]::GetByName($groupName)
    } 
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

    .Outputs
    A [ZoomGroup] object containing the newly-created group.
#>
function Add-ZoomGroup() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.String]$groupName
    )
    $group = Get-ZoomGroup -groupName $groupName
    if ($null -ne $group) {
        Throw "Add-ZoomGroup: Group name ""$groupName"" already exists."
        return $null
    }
    $thisGroup = [ZoomGroup]::Create($groupName)
    return $thisGroup
}

<#
    .Synopsis
    Removes a Zoom group

    .Description
    Removes a Zoom group

    .Parameter GroupName
    The name of the group

    .Parameter Group
    A [ZoomGroup] object representing the group to be deleted

    .Example
    Remove-ZoomGroup -groupName "Marketing Execs"
#>
function Remove-ZoomGroup() {
    [CmdletBinding(DefaultParameterSetName="name")]
    Param(
        [Parameter(ParameterSetName="name", Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.String]$groupName,
        [Parameter(ParameterSetName="group", Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [ZoomGroup]$group
    )

    process {
        if ($group) { $thisGroup = $group }
        if ($groupName) {
            $thisGroup = Get-ZoomGroup -groupName $groupName
            if ($null -eq $thisGroup) {
                Throw "Remove-ZoomGroup: Group name ""$groupName"" could not be found."
                return $null
            }
        }
        $thisGroup.Delete()
    }
}

<#
    .Synopsis
    Returns all Zoom users in the account

    .Description
    Returns an array of all Zoom users in the account.

    .Example
    Get-ZoomUsers

    .Outputs
    An array of [ZoomUser] objects.
#>
function Get-ZoomUsers() {
    return [ZoomUser]::GetUsers()
}

<#
    .Synopsis
    Returns all Zoom users in a group

    .Description
    Returns an array of all Zoom users in a particular group, based upon group name.
    Users are returned as user 'stub' objects to avoid making too many API calls. To load all user details, call $user.Load() on each stub.

    .Parameter GroupName
    The name of the group

    .Parameter Group
    A [ZoomGroup] object representing the group to be enumerated

    .Example
    Get-ZoomGroupUsers -groupName "Marketing Execs"
#>
function Get-ZoomGroupUsers() {
    [CmdletBinding(DefaultParameterSetName="name")]
    Param(
        [Parameter(ParameterSetName="name", Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [System.String]$groupName,
        [Parameter(ParameterSetName="group", Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [ZoomGroup]$group
    )

    process {
        if ($group) { $thisGroup = $group }
        if ($groupName) {
            [ZoomGroup]$thisGroup = Get-ZoomGroup -groupName $groupName
            if ($null -eq $thisGroup) {
                Throw "Get-ZoomGroupUsers: Group name ""$groupName"" could not be found."
                return $null
            }
        }
        
        $thisGroup.GetMembers()
        return $thisGroup.members
    }
}

<#
    .Synopsis
    Adds Zoom users to a group

    .Description
    Adds one or more Zoom users to a group

    .Parameter Emails
    The email addresses of the users, either singular or as an array

    .Parameter Group
    The group to add users to, either as a ZoomGroup object or specified by name

    .Example
    Add-ZoomUsersToGroup -groupName "Marketing Execs" -emails @("percy@cactus.email", "michelle@cactus.email")
#>
function Add-ZoomUsersToGroup() {
    [CmdletBinding(DefaultParameterSetName="email")]
    Param(
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline=$true, Position=1)]
        [System.String[]]$emails, 
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline=$true, Position=1)]
        [ZoomUser[]]$users, 
        [Parameter(ParameterSetName="user", Mandatory=$true, Position=0)]
        [Parameter(ParameterSetName="email", Mandatory=$true, Position=0)]
        $group
    )

    Begin {
        if ($group -is [System.String]) {
            [ZoomGroup]$thisGroup = Get-ZoomGroup -groupName $group
            if ($null -eq $thisGroup) {
                Throw "Add-ZoomUsersToGroup: Group name ""$group"" could not be found."
                return $null
            }
        } elseif ($group -is [ZoomGroup]) {
            $thisGroup = $group
        } else {
            Throw "Add-ZoomUsersToGroup: The -group parameter should either be the name of a group or a ZoomGroup object."
            return $null
        }
    }

    Process {
        if ($PSCmdlet.ParameterSetName -eq "email") {
            $thisGroup.AddMembers($emails)
        } elseif ($PSCmdlet.ParameterSetName -eq "user") {
            $thisGroup.AddMembers($users.email)
        }
    }
}


<#
    .Synopsis
    Removes Zoom user from a group

    .Description
    Removes one or more Zoom users from a group

    .Parameter Email
    The email address of the user to remove

    .Parameter User
    The ZoomUser user object

    .Parameter Group
    The group to remove users from, either as a ZoomGroup object or specified by name

    .Example
    Remove-ZoomUsersFromGroup -group "Marketing Execs" -emails @("percy@cactus.email", "michelle@cactus.email")
#>
function Remove-ZoomUsersFromGroup() {
    [CmdletBinding(DefaultParameterSetName="email")]
    Param(
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline=$true, Position=1)]
        [System.String]$email, 
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline=$true, Position=1)]
        [ZoomUser]$user, 
        [Parameter(ParameterSetName="user", Mandatory=$true, Position=0)]
        [Parameter(ParameterSetName="email", Mandatory=$true, Position=0)]
        $group
    )

    Begin {
        if ($group -is [System.String]) {
            [ZoomGroup]$thisGroup = Get-ZoomGroup -groupName $group
            if ($null -eq $thisGroup) {
                Throw "Remove-ZoomUsersFromGroup: Group name ""$group"" could not be found."
                return $null
            }
        } elseif ($group.GetType().Name -eq 'ZoomGroup') {
            $thisGroup = $group
        } else {
            Throw "Remove-ZoomUsersFromGroup: The -group parameter should either be the name of a group or a ZoomGroup object."
            return $null
        }
    }

    Process {
        if ($PSCmdlet.ParameterSetName -eq "email") {
            $thisUser = Get-ZoomUser -email $email
            if ($thisUser.groupIDs -contains $thisGroup.id) {
                $thisGroup.RemoveMember($thisUser)
            } else {
                Write-Error "$email is not a member of $($thisGroup.name)"
            }
        } elseif ($PSCmdlet.ParameterSetName -eq "user") {
            $thisGroup.RemoveMember($user)
        }
    }
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
    [CmdletBinding(DefaultParameterSetName = 'email')]
    Param(
        [Parameter(ParameterSetName="email", Position=0, Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Position=0, Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="email", Position=1, Mandatory=$true)]
        [Parameter(ParameterSetName="user", Position=1, Mandatory=$true)]
        [ZoomLicenseType]$license
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Set-ZoomUserLicenseState: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.SetLicenseStatus($license)
    }
}

<#
    .Synopsis
    Sets the feature state of a user

    .Description
    Assigns or removes webinar and large meeting add-ons

    .Parameter Email
    The email address of the user to modify.

    .Parameter Webinar
    A boolean value indicating if the user should be assigned the webinar feature

    .Parameter WebinarCapacity
    The webinar capacity to assign. Must be one of 100, 500, 1000, 3000, 5000, 10000

    .Parameter LargeMeeting
    A boolean value indicating if the user should be assigned the large meeting feature

    .Parameter LargeMeetingCapacity
    The large meeting capacity to assign. Must be either 500 or 1000.

    .Example
    Set-ZoomUserFeatureState -email jerome@cactus.email -webinar -webinarCapacity 1000 -largeMeeting -largeMeetingCapacity 500
#>
function Set-ZoomUserFeatureState() {
    [CmdletBinding(DefaultParameterSetName = 'email')]
    Param(
        [Parameter(ParameterSetName="email", Position=0, Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Position=0, Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(Mandatory=$true)]
        [System.Boolean]$webinar,
        [ValidateSet(100, 500, 1000, 3000, 5000, 10000)]
        [System.Int32]$webinarCapacity,
        [Parameter(Mandatory=$true)]
        [System.Boolean]$largeMeeting,
        [ValidateSet(500, 1000)]
        [System.Int32]$largeMeetingCapacity
    )

    process {
        if ( $webinar -eq $true -and ($webinarCapacity -ne 100 -and $webinarCapacity -ne 500 -and $webinarCapacity -ne 1000 -and $webinarCapacity -ne 3000 -and $webinarCapacity -ne 5000 -and $webinarCapacity -ne 10000)) {
            Throw "Set-ZoomUserFeatureState: Invalid webinarCapacity value."
            return $null
        }
        if ( $largeMeeting -eq $true -and ($largeMeetingCapacity -ne 500 -and $largeMeetingCapacity -ne 1000)) {
            Throw "Set-ZoomUserFeatureState: Invalid largeMeetingCapacity value."
            return $null
        }
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Set-ZoomUserFeatureState: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.SetFeatureStatus($webinar, $largeMeeting, $webinarCapacity, $largeMeetingCapacity)
    }
}

<#
    .Synopsis
    Modifies a user's profile details

    .Description
    Modifies a Zoom user's name, licensed state, job title, timezone, company name, location and phone number.
    An alternative to calling Set-ZoomUserDetails is to modify the ZoomUser object directly, then call [ZoomUser].Save().

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
    Param(
        [Parameter(Mandatory=$true)][System.String]$email, 
        [System.String]$firstName, 
        [System.String]$lastName, 
        [ZoomLicenseType]$license, 
        [System.String]$timezone, 
        [System.String]$jobTitle, 
        [System.String]$company, 
        [System.String]$location, 
        [System.String]$phoneNumber
    )
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Throw "Set-ZoomUserDetails: User ""$email"" could not be found."
        return $null
    }
    [ZoomUser]$thisUser = [ZoomUser]::new($email)
    $thisUser.Update($firstName, $lastName, $license, $timezone, $jobTitle, $company, $location, $phoneNumber)
}

<#
    .Synopsis
    Sets the user's password

    .Description
    Sets the password of the Zoom user.
    Note: This function performs no validation on password length or complexity.

    .Parameter Email
    The email address of the user

    .Parameter User
    The user's ZoomUser object

    .Parameter Password
    The password to set, exressed as a SecureString

    .Example
    Set-ZoomUserPassword -email "jerome@cactus.email" -password (ConvertTo-SecureString -String "Marketing101!" -AsPlainText -Force)
#>
function Set-ZoomUserPassword() {
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="user", Position=0, Mandatory=$true)]
        [Parameter(ParameterSetName="email", Position=0, Mandatory=$true)]
        [securestring]$password
    )

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Set-ZoomUserPassword: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.SetPassword($password)
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
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="user", Position=0, Mandatory=$true)]
        [Parameter(ParameterSetName="email", Position=0, Mandatory=$true)]
        [boolean]$enabled
    )

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Set-ZoomUserPassword: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.SetStatus($enabled)
    }
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
    Add-ZoomUser -email "francisco@cactus.email" -firstName Francisco -lastName Hunter -license Basic timezone "Europe/London" -jobTitle "Graphic Artist" -company "Cactus Industries (Europe)" -Location Basildon -phoneNumber "+44 1268 533333" -groupName "Designers"
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
        Throw "Add-ZoomUser: User ""$email"" already exists."
        return $null
    }
    $thisUser = [ZoomUser]::Create($email, $firstName, $lastName, $license, $tiemzone, $jobTitle, $company, $location, $phoneNumber, $groupName)
    return $thisUser
}

<#
    .Synopsis
    Removes a user

    .Description
    Removes a Zoom user from the account

    .Parameter Email
    The email address of the user

    .Parameter User
    The ZoomUser object of the user to remove

    .Example
    Remove-ZoomUser -email "francisco@cactus.email"
#>
function Remove-ZoomUser() {
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user
    )

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Remove-ZoomUser: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.Delete()
    }
}

<#
    .Synopsis
    Adds a new assistant / delegate to a user

    .Description
    Adds a new assistant / delegate to a user, based upon email addresses. Caveat: Zoom requires that both the host and assistant must be fully-licensed users.

    .Parameter Email
    The email address of the user to add the assistant to.

    .Parameter User
    The ZoomUser object of the user to add the assistant to.

    .Parameter Assistant
    The assistant to add. Can either be passed as an email address or a ZoomUser object.

    .Example
    Add-ZoomUserAssistant -email "jerome@cactus.email" -assistant "francisco@cactus.email"
#>
function Add-ZoomUserAssistant() {
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="email", Mandatory=$true, Position=0)]
        [Parameter(ParameterSetName="user", Mandatory=$true, Position=0)]
        $assistant
    )
    Begin {
        if ($assistant -is [System.String]) {
            if ( (Get-ZoomUserExists -email $assistant) -eq $false ) {
                Throw "Add-ZoomUserAssistant: Assistant ""$assistant"" could not be found."
                return $null
            }
            $thisAssistant = $assistant
        } elseIf ( $assistant -is [ZoomUser] ) {
            $thisAssistant = $assistant.Email
        }
    }

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Add-ZoomUserAssistant: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.AddAssistant($thisAssistant)
    }
}

<#
    .Synopsis
    Removes an assistant / delegate from a user

    .Description
    Removes an assistant / delegate from a user, based upon email addresses.

    .Parameter Email
    The email address of the user to remove the assistant from.

    .Parameter User
    The ZoomUser object of the user to remove the assistant from.

    .Parameter Assistant
    The assistant to remove. Can either be passed as an email address or a ZoomUser object.

    .Example
    Remove-ZoomUserAssistant -email "jerome@cactus.email" -assistant "francisco@cactus.email"
#>
function Remove-ZoomUserAssistant() {
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="email", Mandatory=$true, Position=0)]
        [Parameter(ParameterSetName="user", Mandatory=$true, Position=0)]
        $assistant
    )
    Begin {
        if ($assistant -is [System.String]) {
            if ( (Get-ZoomUserExists -email $assistant) -eq $false ) {
                Throw "Add-ZoomUserAssistant: Assistant ""$assistant"" could not be found."
                return $null
            }
            $thisAssistant = get-zoomuser -email $assistant
        } elseIf ( $assistant -is [ZoomUser] ) {
            $thisAssistant = $assistant
        }
    }

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Remove-ZoomUserAssistant: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.RemoveAssistant($thisAssistant.id)
    }
}

<#
    .Synopsis
    Retrieves a user's meetings

    .Description
    Returns an array of ZoomMeeting objects representing the Zoom meetings for the specified user

    .Parameter Email
    The email address of the user to retrieve meetings for.
    
    .Parameter User
    The ZoomUser object of the user to retrieve meetings for.

    .Parameter MeetingType
    The type of meeting to retrieve. Available options are: Live, Upcoming and Scheduled

    .Parameter Detailed
    If specified, the -Detailed switch retrieves the full meeting details for each meeting. This requires an API call for each meeting instance, so can be time consuming.
    For most uses, the -Detailed switch is not required.

    .Example
    Get-ZoomMeetings -email "jerome@cactus.email" -meetingType Upcoming
#>
function Get-ZoomMeetings() {
    Param(
        [CmdletBinding(DefaultParameterSetName="email")]
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="email", Position=0)]
        [Parameter(ParameterSetName="name", Position=0)]
        [ZoomMeetingType]$meetingType,
        [Parameter(ParameterSetName="email")]
        [Parameter(ParameterSetName="name")]
        [switch]$detailed
    )

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisEmail = $user.email
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Get-ZoomMeetings: User ""$email"" could not be found."
                return $null
            }
            $thisEmail = $email
        }
        if ($null -eq $meetingType) { $meetingType = [ZoomMeetingType]::Live }
        $zoommeetings = [ZoomMeeting]::GetMeetings($thisEmail, $meetingType, $detailed.IsPresent)
        return $zoommeetings
    }
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
    Param(
        [Parameter(Position=0, Mandatory=$true)]
        [System.String]$meetingID
    )
    $meeting = [ZoomMeeting]::GetMeeting($meetingID)
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
    Param(
        [CmdletBinding(DefaultParameterSetName="id")]
        [Parameter(ParameterSetName="id", Mandatory=$true, ValueFromPipeline, position=0)]
        [System.String]$meetingID,
        [Parameter(ParameterSetName="meeting", Mandatory=$true, ValueFromPipeline, position=0)]
        [ZoomUser]$meeting
    )

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'meeting') {
            $thisMeeting = $meeting
        }
        if ($PSCmdlet.ParameterSetName -eq 'id') {
            $thisMeeting = [ZoomMeeting]::GetMeetingStub($meetingID)
        }
        $thisMeeting.StopMeeting()
    }
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
    $roles = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles" -Headers (Get-ZoomAuthHeader) -Method Get
    Write-Debug "Get-ZoomRole: Retrieved $($roles.total_records) roles."
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
        Throw "Get-ZoomRoleUsers: Role ""$role"" could not be found."
        return $null
    }
    $roleusers = @()
    $page1 = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members?page_size=100" -Headers (Get-ZoomAuthHeader) -Method Get
    Write-Debug "Get-ZoomRoleUsers: $($page1.total_records) users in $($page1.page_count) pages. Adding page 1 containing $($page1.members.count) records."
    $roleusers += $page1.members
    if ($page1.page_count -gt 1) {
        for ($count=2; $count -le $page1.page_count; $count++) {
            $page = Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members?page_size=100&page_number=$count" -Headers (Get-ZoomAuthHeader)
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
    Note: roles are not cumulative. A user can only be assigned one role.

    .Parameter RoleName
    The role name

    .Example
    Set-ZoomRoleUsers -email "melissa@cactus.email" -roleName "Level 2 Service Desk"
#>
function Set-ZoomRoleUsers() {
    Param([Parameter(Mandatory=$true)]$emails, [Parameter(Mandatory=$true)][System.String]$roleName)
    if ($rolename.ToLower() -eq "owner" -or $roleName.ToLower() -eq "member") {
        Throw "Set-ZoomRoleUsers: Users cannot be added to the Member or Owner roles."
        return $null
    }
    $role = Get-ZoomRole -roleName $roleName
    if ($null -eq $role) {
        Throw "Set-ZoomRoleUsers: Role ""$role"" could not be found."
        return $null
    }
    $members = @()
    $emails | ForEach-Object {
        if ( (Get-ZoomUserExists -email $_) -eq $false ) {
            Throw "Set-ZoomRoleUsers: User ""$_"" could not be found."
            return $null
        }
        $member = New-Object -TypeName psobject
        $member | Add-Member -Name email -Value $_ -MemberType NoteProperty
        $members += $member
    }
    $body = New-Object -TypeName psobject
    $body | Add-Member -Name members -Value $members -MemberType NoteProperty
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members" -Headers (Get-ZoomAuthHeader) -Body ($body | ConvertTo-Json) -Method Post
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
        Throw "Remove-ZoomRoleUser: Role ""$role"" could not be found."
        return $null
    }
    if ( (Get-ZoomUserExists -email $email) -eq $false ) {
        Throw "Remove-ZoomRoleUser: User ""$email"" could not be found."
        return $null
    }
    Invoke-RestMethod -Uri "https://api.zoom.us/v2/roles/$($role.id)/members/$email" -Headers (Get-ZoomAuthHeader) -Method Delete
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
        Throw "Get-ZoomMeetingReport: User ""$email"" could not be found."
        return $null
    }
    [System.Timespan]$timeDifference = $to - $from
    if ($timeDifference.TotalDays -gt 31) {
        Throw "Get-ZoomMeetingReport: The date range defined by the -from and -to parameters cannot be more than one month apart."
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
    The ResolveNames switch attempts to retrieve each participant's name and email. This can only be completed for participants that are part of your account.
    Note: Enabling this feature adds an extra API call for each participant, and can heavily impact performance.

    .Example
    Get-ZoomMeetingParticipantReport -meetingID abcefghh123456789== -resolveNames
#>
function Get-ZoomMeetingParticipantReport() {
    Param([Parameter(Mandatory=$true)][System.String]$meetingID, [switch]$resolveNames)
    $report = [ZoomMeetingParticipantReport]::new($meetingID, $resolveNames.IsPresent)
    return $report
}

<#
    .Synopsis
    Generates the Zoom webinar participant report

    .Description
    Generates the Zoom participant report for a specific webinar

    .Parameter WebinarID
    The ID of the webinar.

    .Parameter ResolveNames
    The ResolveNames switch attempts to retrieve each participant's name and email. This can only be completed for participants that are part of your account.
    Note: Enabling this feature adds an extra API call for each participant, and can heavily impact performance.

    .Example
    Get-ZoomWebinarParticipantReport -meetingID 98765432109 -resolveNames
#>
function Get-ZoomWebinarParticipantReport() {
    Param([Parameter(Mandatory=$true)][System.String]$webinarID, [switch]$resolveNames)
    $report = [ZoomWebinarParticipantReport]::new($webinarID, $resolveNames.IsPresent)
    return $report
}
#`------------------------------------------------------------------------------------------------------------------------------------------------------------

#region Utility functions

<#
    .Synopsis
    Generates a JWT

    .Description
    Generates a short-lived JSON Web Token to authenticate with Zoom

    .Parameter apiKey
    Your application's API Key

    .Parameter apiSecret
    Your application's API Secret

    .Parameter validForSeconds
    The period of time for which the token should be valid.
    For optimum security, you should generate a new JWT for each and every API call and the token validity should be no more than 30 seconds.

    .Example
    Get-JSONWebToken -apiKey abc123 -apiSecret 789xyz -validForSeconds 10

    .Notes
    This function was generated from a source script found here: https://www.reddit.com/r/PowerShell/comments/8bc3rb/generate_jwt_json_web_token_in_powershell/
#>
function Get-JSONWebToken (){
    Param (
        [Parameter(Mandatory = $True)][string]$apiKey,
        [Parameter(Mandatory = $True)]$apiSecret,
        [int]$validforSeconds = $null
    )

    $exp = [int][double]::parse((Get-Date -Date $((Get-Date).addseconds($ValidforSeconds).ToUniversalTime()) -UFormat %s)) # Grab Unix Epoch Timestamp and add desired expiration.

    [hashtable]$header = @{alg = "HS256"; typ = "JWT"}
    [hashtable]$payload = @{iss = $apiKey; exp = $exp}

    $headerjson = $header | ConvertTo-Json -Compress
    $payloadjson = $payload | ConvertTo-Json -Compress
    
    $headerjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')
    $payloadjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')

    $ToBeSigned = $headerjsonbase64 + "." + $payloadjsonbase64

    $SigningAlgorithm = New-Object System.Security.Cryptography.HMACSHA256

    $SigningAlgorithm.Key = [System.Text.Encoding]::UTF8.GetBytes($apiSecret)
    $Signature = [Convert]::ToBase64String($SigningAlgorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($ToBeSigned))).Split('=')[0].Replace('+', '-').Replace('/', '_')
    
    $token = "$headerjsonbase64.$payloadjsonbase64.$Signature"
    return $token
}

<#
    .Synopsis
    Gets the Zoom authentication HTTP header

    .Description
    Gets the header object required to authenticate with the Zoom API
#>
function Get-ZoomAuthHeader() {
    $token = Get-JSONWebToken -apiKey $global:api_key -apiSecret $global:api_secret -ValidforSeconds 10

    $header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $header.Add('Content-Type','application/json')
    $header.Add('Authorization',"Bearer $token")

    return $header
}

<#
    .Synopsis
    Sets the Zoom API Key and Secret

    .Description
    Sets the Zoom API Key and Secret required for the generation of JSON Web Tokens for authentication. Must be called before any other function.
    For more information, see: https://marketplace.zoom.us/docs/guides/auth/jwt 

    .Parameter apiKey
    Your application's API key

    .Parameter apiKey
    Your application's API secret

    .Example
    Set-ZoomAuthToken -apiKey abc123 -apiSecret -xyz789
#>
function Set-ZoomAuthToken() {
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [System.String]$apiKey,
        [Parameter(Mandatory=$true, Position=0)]
        [System.String]$apiSecret
    )
    $global:api_key = $apiKey
    $global:api_secret = $apiSecret
}

#endregion Utility functions

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Debug "ZoomLibrary module loaded"

#`------------------------------------------------------------------------------------------------------------------------------------------------------------

Export-ModuleMember -Function Get-ZoomAuthHeader
Export-ModuleMember -Function Set-ZoomAuthToken
Export-ModuleMember -Function Get-ZoomUser -Alias gzu
Export-ModuleMember -Function Get-ZoomUserExists -Alias gzue
Export-ModuleMember -Function Get-ZoomGroup -Alias gzg
Export-ModuleMember -Function Add-ZoomGroup -Alias azg
Export-ModuleMember -Function Remove-ZoomGroup -Alias rzg
Export-ModuleMember -Function Get-ZoomUsers -Alias gzusers
Export-ModuleMember -Function Get-ZoomGroupUsers -Alias gzgu
Export-ModuleMember -Function Add-ZoomUsersToGroup -Alias azug
Export-ModuleMember -Function Remove-ZoomUsersFromGroup -Alias rzug
Export-ModuleMember -Function Set-ZoomUserLicenseState -Alias szul
Export-ModuleMember -Function Set-ZoomUserDetails -Alias szud
Export-ModuleMember -Function Set-ZoomUserPassword -Alias szup
Export-ModuleMember -Function Set-ZoomUserStatus -Alias szus
Export-ModuleMember -Function Add-ZoomUser -Alias azu
Export-ModuleMember -Function Remove-ZoomUser -Alias rzu
Export-ModuleMember -Function Add-ZoomUserAssistant -Alias azua
Export-ModuleMember -Function Remove-ZoomUserAssistant -Alias rzua
Export-ModuleMember -Function Set-ZoomUserFeatureState -Alias szufs
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
Export-ModuleMember -Function Get-ZoomWebinarParticipantReport -Alias gzwpr