
Using Module .\ZoomLibraryClasses.psm1

<#
    .Synopsis
    Returns a Zoom user object

    .Description
    Returns a Zoom user object based on email address
    Note: to reduce the number of API calls required, a zoom user's assistants and schedulers are not populated automatically. To populate, call [ZoomUser].GetAssistants or [ZoomUser].GetSchedulers as required.

    .Parameter Email
    The email address of the user

    .Example
    Get-ZoomUser -email jerome@cactus.email

    .Outputs
    A [ZoomUser] object for the specified user
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
    Get-ZoomUserExists -email jerome@cactus.email

    .Outputs
    A boolean value indicating existence
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

    .Outputs
    A [ZoomGroup] object
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

    .Outputs
    An array of [ZoomUser] objects.
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
    The ZoomUser user object to remove

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

    .Parameter User
    The ZoomUser object to modify.

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

    .Parameter User
    The ZoomUser object to modify.

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

    .Parameter User
    The ZoomUser object to modify.

    .Parameter FirstName
    The Zoom user's first name
    
    .Parameter LastName
    The Zoom user's last name

    .Parameter License
    The type of license to apply. Valid values are: [ZoomLicenseType]::Basic, [ZoomLicenseType]::Licensed and [ZoomLicenseType]::OnPrem

    .Parameter TimeZone
    The time zone the Zoom user is in. Time zone to be in specific TZD format. For more information, see: https://marketplace.zoom.us/docs/api-reference/other-references/abbreviation-lists#timezones

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
    [CmdletBinding(DefaultParameterSetName = 'email')]
    Param(
        [Parameter(ParameterSetName="email", Mandatory=$true, ValueFromPipeline)]
        [System.String]$email,
        [Parameter(ParameterSetName="user", Mandatory=$true, ValueFromPipeline)]
        [ZoomUser]$user,
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$firstName, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$lastName, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [ZoomLicenseType]$license, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$timezone, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$jobTitle, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$company, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$location, 
        [Parameter(ParameterSetName="email")][Parameter(ParameterSetName="user")]
        [System.String]$phoneNumber
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'user') {
            $thisUser = $user
        }
        if ($PSCmdlet.ParameterSetName -eq 'email') {
            if ( (Get-ZoomUserExists -email $email) -eq $false ) {
                Throw "Set-ZoomUserDetails: User ""$email"" could not be found."
                return $null
            }
            $thisUser = Get-ZoomUser -email $email
        }
        $thisUser.Update($firstName, $lastName, $license, $timezone, $jobTitle, $company, $location, $phoneNumber)
    }
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

    .Outputs
    A [ZoomUser] object.
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
    $thisUser = [ZoomUser]::Create($email, $firstName, $lastName, $license, $timezone, $jobTitle, $company, $location, $phoneNumber, $groupName)
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

<#
    .Synopsis
    Generates the Zoom operations log report

    .Description
    Generates the Zoom meetings report for a specific user in a defined timespan. Note: the timespan defined by the -from and -to parameters cannot be more than one month in duration.

    .Parameter From
    The start date of the report, expressed as a datetime object.

    .Parameter From
    The end date of the report, expressed as a datetime object.

    .Example
    Get-ZoomOperationsReport -from "2020-01-01" -to "2020-02-01"
#>
function Get-ZoomOperationsReport() {
    Param(
        [Parameter(Mandatory=$true)][System.DateTime]$from,
        [Parameter(Mandatory=$true)][System.DateTime]$to
    )

    [System.Timespan]$timeDifference = $to - $from
    if ($timeDifference.TotalDays -gt 31) {
        Throw "Get-ZoomOperationsReport: The date range defined by the -from and -to parameters cannot be more than one month apart."
        return $null
    }
    $zoomopslog = [ZoomOperationsReport]::new($from, $to)
    return $zoomopslog
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
Export-ModuleMember -Function Get-ZoomOperationsReport -Alias gzor