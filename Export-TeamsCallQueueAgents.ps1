<#

.SYNOPSIS
 
    Export-TeamsCallQueueAgents.ps1
    This script will display all Teams Call Queue agents and group members.
 
.DESCRIPTION
    
    Author: csatswo
    This script outputs to the terminal a table of all Teams Call Queue agents.  The table will show agents directly assigned to a queue as well as agents that are assigned via group membership.
    
.LINK

    https://github.com/csatswo/Export-TeamsCallQueueAgents.ps1
 
.EXAMPLE 
    
    .\Export-TeamsCallQueueAgents.ps1 -Path C:\Temp\Agents.csv -UserName admin@domain.com -OverrideAdminDomain domain.onmicrosoft.com

.NOTES
    
    If Call Queues have 'ConferenceMode' enabled, the script will display many warningsget-c

#>

Param(
    [Parameter(mandatory=$true)][String]$Path,
    [Parameter(mandatory=$true)][String]$UserName,
    [Parameter(mandatory=$false)][string]$OverrideAdminDomain
)

# Check for MSOnline module and install if missing

if (Get-Module -ListAvailable -Name MSOnline) {
    
    Write-Host "`nMSOnline module is installed" -ForegroundColor Cyan
    Import-Module MSOnline

} else {

    Write-Host "`nMSOnline module is not installed" -ForegroundColor Red
    Write-Host "`nInstalling module..." -ForegroundColor Cyan
    Install-Module MSOnline

}

# Check for SfBO module

if (Get-Module -ListAvailable -Name SkypeOnlineConnector) {
    
    Write-Host "`nSkype Online Module installed" -ForegroundColor Cyan
    Import-Module SkypeOnlineConnector

} else {

    Write-Error -Message "Skype Online Module not installed, please download and install then try again."
        
    break

}

# Connect to MSOnline and SfBO

Write-Host "`nConnecting to MSOnline and Skype Online" -ForegroundColor Cyan
Write-Host "You may be prompted more than once to authenticate" -ForegroundColor Yellow
Write-Host `n

if((Get-PSSession | Where-Object {$_.ComputerName -like "*.online.lync.com"}).State -ne "Opened") {

    if ($OverrideAdminDomain) {

        $global:PSSession = New-CsOnlineSession -UserName $UserName -OverrideAdminDomain $OverrideAdminDomain

        } else {

        $global:PSSession = New-CsOnlineSession -UserName $UserName

        }
    
    Import-PSSession $global:PSSession -AllowClobber | Out-Null
    Enable-CsOnlineSessionForReconnection
    Connect-MsolService | Out-Null

    }

# Start script loops

$Queues = @()
$Queues = Get-CsCallQueue -WarningAction SilentlyContinue
$CustomObject = @()

foreach ($Queue in $Queues) {
    
    $Users = @()
    $Groups = @()
    $Users = (Get-CsCallQueue -WarningAction SilentlyContinue -Identity $Queue.Identity).Users
    $Groups = (Get-CsCallQueue -WarningAction SilentlyContinue -Identity $Queue.Identity).DistributionLists

    foreach ($User in $Users) {
        
        $QueueUser = @()
        $QueueMsolUser = @()
        $QueueUser = Get-CsOnlineUser -Identity $User.Guid
        $QueueMsolUser = Get-MsolUser -ObjectId $User.Guid
        $QueueUserProperties = @{
            UserDisplayName = $QueueMsolUser.DisplayName
            UserPrincipalName = $QueueMsolUser.UserPrincipalName
            UserSipAddress = $QueueUser.SipAddress
            QueueAssignment = "Direct"
            GroupName = "N/A"
            CallQueue = $Queue.Name
            CallQueueId = $Queue.Identity
            ConferenceMode = $Queue.ConferenceMode
            }

        $ObjectProperties = New-Object -TypeName PSObject -Property $QueueUserProperties
        $CustomObject += $ObjectProperties

        }

    foreach ($Group in $Groups) {
        
        $QueueGroup = @()
        $QueueGroupMembers = @()
        $QueueGroup = Get-MsolGroup -ObjectId $Group.Guid
        $QueueGroupMembers = Get-MsolGroupMember -GroupObjectId $Group.Guid

        foreach ($QueueGroupMember in $QueueGroupMembers) {

            if((Get-MsolUser -ObjectId $QueueGroupMember.ObjectId).isLicensed -like "False") {

                Write-Host "`nWarning: " -ForegroundColor Yellow -NoNewline
                Write-Host $QueueGroupMember.DisplayName -ForegroundColor White -NoNewline
                Write-Host " is not licensed" -ForegroundColor Yellow -NoNewline
                Write-Host `n
                
                $Member = @()
                $MsolMember = @()
                $MsolMember = Get-MsolUser -ObjectId $QueueGroupMember.ObjectId
                $MemberProperties = @{
                    UserDisplayName = $MsolMember.DisplayName
                    UserPrincipalName = $MsolMember.UserPrincipalName
                    UserSipAddress = "Not Licensed"
                    QueueAssignment = $QueueGroup.GroupType
                    GroupName = $QueueGroup.DisplayName
                    CallQueue = $Queue.Name
                    CallQueueId = $Queue.Identity
                    ConferenceMode = $Queue.ConferenceMode
                    }
                
                $ObjectProperties = New-Object -TypeName PSObject -Property $MemberProperties
                $CustomObject += $ObjectProperties
                
                } else {
                
                $Member = @()
                $MsolMember = @()
                $Member = Get-CsOnlineUser -Identity $QueueGroupMember.ObjectId

                if(($Member).SipAddress -notlike "sip:*") {

                    $MsolMember = Get-MsolUser -ObjectId $QueueGroupMember.ObjectId
                    $MemberProperties = @{
                        UserDisplayName = $MsolMember.DisplayName
                        UserPrincipalName = $MsolMember.UserPrincipalName
                        UserSipAddress = "Not Teams Enabled"
                        QueueAssignment = $QueueGroup.GroupType
                        GroupName = $QueueGroup.DisplayName
                        CallQueue = $Queue.Name
                        CallQueueId = $Queue.Identity
                        ConferenceMode = $Queue.ConferenceMode
                        }
                    
                    $ObjectProperties = New-Object -TypeName PSObject -Property $MemberProperties
                    $CustomObject += $ObjectProperties
                    
                    } else {
                    
                    $MsolMember = Get-MsolUser -ObjectId $QueueGroupMember.ObjectId
                    $MemberProperties = @{
                        UserDisplayName = $MsolMember.DisplayName
                        UserPrincipalName = $MsolMember.UserPrincipalName
                        UserSipAddress = $Member.SipAddress
                        QueueAssignment = $QueueGroup.GroupType
                        GroupName = $QueueGroup.DisplayName
                        CallQueue = $Queue.Name
                        CallQueueId = $Queue.Identity
                        ConferenceMode = $Queue.ConferenceMode
                        }

                    $ObjectProperties = New-Object -TypeName PSObject -Property $MemberProperties
                    $CustomObject += $ObjectProperties
                    
                    }

                }

            }

        }

    }

Write-Output $CustomObject | Select-Object CallQueue,QueueAssignment,GroupName,UserDisplayName,UserSipAddress | Sort-Object -Property @{Expression="CallQueue"},@{Expression="GroupName"},@{Expression="UserDisplayName"} | Format-Table

$CustomObject | Select-Object CallQueue,CallQueueId,ConferenceMode,QueueAssignment,GroupName,UserDisplayName,UserPrincipalName,UserSipAddress | Sort-Object -Property @{Expression="CallQueue"},@{Expression="GroupName"},@{Expression="UserDisplayName"} | Export-Csv -Path $Path -NoTypeInformation

Write-Host "Export saved to " -ForegroundColor Cyan -NoNewline
Write-Host $Path -ForegroundColor Yellow -NoNewline
Write-Host `n
