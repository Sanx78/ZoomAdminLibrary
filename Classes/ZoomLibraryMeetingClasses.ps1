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

Write-Debug "ZoomLibraryMeetingClasses loaded"