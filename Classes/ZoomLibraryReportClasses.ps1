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

Write-Debug "ZoomLibraryReportClasses loaded"