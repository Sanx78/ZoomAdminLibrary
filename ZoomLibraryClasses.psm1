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

Class ZoomOperationsLogEntry {
    [System.DateTime]$time
    [System.String]$operator
    [System.String]$category
    [System.String]$action
    [System.String]$detail

    ZoomOperationsLogEntry($logEntry) {
        $this.time = $logEntry.time
        $this.operator = $logEntry.operator
        $this.category = $logEntry.category_type
        $this.action = $logEntry.action
        $this.detail = $logEntry.operation_detail
    }
}

Class ZoomOperationsReport {
    [ZoomOperationsLogEntry[]]$entries

    ZoomOperationsReport([System.DateTime]$from, [System.DateTime]$to) {
        $fromString = $from.ToString("yyyy-MM-dd")
        $toString = $to.ToString("yyyy-MM-dd")

        $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/operationlogs?page_size=100&from=$fromString&to=$toString" -Headers (Get-ZoomAuthHeader)
        do {
            $result.operation_logs | ForEach-Object {
                $this.entries += [ZoomOperationsLogEntry]::new($_)
            }
            Write-Debug "[ZoomOperationsReport]::new - Next page token: $($result.next_page_token)"
            $result = Invoke-RestMethod -Uri "https://api.zoom.us/v2/report/operationlogs?page_size=100&next_page_token=$($result.next_page_token)&from=$fromString&to=$toString" -Headers (Get-ZoomAuthHeader)
        } while ([System.String]::IsNullOrWhiteSpace($result.next_page_token) -eq $false)
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
    [System.Boolean]$isStub

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
        $this.isStub = $false
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

        $this.isStub = $false
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
        if ([System.String]::IsNullOrWhiteSpace($this.vanityURL) -eq $false ) { $params.Add("vanity_name", $this.vanityURL) }
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
        $this.GetAssistants()
    }

    RemoveAssistant([System.String]$assistant) {
        Invoke-RestMethod -Uri "https://api.zoom.us/v2/users/$($this.email)/assistants/$assistant" -Headers (Get-ZoomAuthHeader) -Method DELETE
        $this.GetAssistants()
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
        $thisUser.isStub = $true
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

Write-Debug "ZoomLibraryClasses loaded."