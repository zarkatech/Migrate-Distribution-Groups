$ScriptInfo = @"
================================================================================
Create-TargetShadowGroups.ps1 | v3.2.5 DRAFT
by Roman Zarka | Microsoft Services
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$MailGroupsAuditFile = "11270338_MASTER_AuditMailGroups.csv"
$MailGroupMembersAuditFile = "11270338_MASTER_AuditMailGroupMembers.csv"
$SetRbacScope = $true
    $RbacAttribute = 'CustomAttribute1'
    $RbacValue = 'USA'
$NewGroupNamePrefix = ''
$NewGroupNameSuffix = ' (USA)'
$ShadowGroupPrefix = 'shadow_'
    $CreateShadowGroups = $true
    $AddShadowGroupMembers = $true
    $CutoverShadowGroups = $false
    $PurgeShadowGroups = $false
$MicroDelay = 0
$Debug = $true

# --- Initialize log file

$Timestamp = Get-Date -Format MMddhhmm
If ($BatchName -ne "" -and $BatchName -ne $null) { $Timestamp = $Timestamp + "_$BatchName" }
$RunLog = $Timestamp + "_CreateTargetShadowGroups.log"
If ($Debug) { $RunLog = "DEBUG_" + $RunLog }
Function Write-Log ($LogString) {
    $LogStatus = $LogString.Split(":")[0]
    If ($LogStatus -eq "ALERT") {
        Write-Host $LogString -BackgroundColor Yellow -ForegroundColor Black
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "INFO") {
        Write-Host $LogString -ForegroundColor Cyan
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ERROR") {
        Write-Host $LogString -BackgroundColor Red -ForegroundColor White
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "SUCCESS") {
        Write-Host $LogString -ForegroundColor Green
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "DEBUG") {
        Write-Host $LogString -ForegroundColor DarkGray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "") {
        Write-Host ""
        $LogString | Out-File $RunLog -Append }
}

# --- Initialize script environment

If ([int]$PsVersionTable.PSVersion.Major -lt 3) { Write-Log "ERROR: Script requires PowerShell v3 or later and must be run from an upgraded console."; Break }
If ((Get-PSSession) -eq $null -or ((Get-PSSession).ConfigurationName) -ne "Microsoft.Exchange") { Write-Log "ERROR: Script must be run from an Exchange session."; Break }
If (Test-Path $MailGroupsAuditFile) { Write-Log "INFO: Import mail groups audit data. [$MailGroupsAuditFile]"; $MailGroups = Import-Csv $MailGroupsAuditFile | Sort Name }
Else { Write-Log "ERROR: Mail groups audit file not found. [$MailGroupsAuditFile]"; Break }
If ($AddShadowGroupMembers) { 
    If (Test-Path $MailGroupMembersAuditFile) { Write-Log "INFO: Import mail group members audit data. [$MailGroupMembersAuditFile]"; $MailGroupMembers = (Import-Csv $MailGroupMembersAuditFile | Sort ParentGroupEmail) }
    Else { Write-Log "ERROR: Mail group members audit file not found. [$MailGroupMembersAuditFile]"; Break } }
If ($Debug) { Write-Log ""; Write-Log "ALERT: Script is in debug mode and changes will NOT be committed." }
Else { Write-Log ""; Write-Log "ALERT: Script is NOT in debug mode and changes will be commited." }
$Continue = Read-Host "Continue? [Y]es or [N]o"; If ($Continue -ne "Y") { Break }

# --- Stage new shadow groups

If ($CreateShadowGroups) {
    ForEach ($Group in $MailGroups) {
        Start-Sleep -Milliseconds $MicroDelay
        $ShadowGroupName = "$ShadowGroupPrefix" + "$NewGroupNamePrefix" + "$($Group.Name)" + "$NewGroupNameSuffix"
        If (($ShadowGroupName | Measure -Character).Characters -gt 64) { Write-Log "ERROR: Group name exceeds the maximum 64 character limit. [$ShadowGroupName]"; Continue }
        $ShadowGroupAlias = "$ShadowGroupPrefix" + "$($Group.Alias)"
        If (($ShadowGroupAlias | Measure -Character).Characters -gt 64) { Write-Log "ERROR: Group alias exceeds the maximum 64 character limit. [$ShadowGroupAlias]"; Continue }
        $ShadowGroupEmail = "$ShadowGroupPrefix" + "$($Group.PrimarySmtpAddress)"
        Write-Log ""; Write-Log "INFO: Stage new shadow group. [$ShadowGroupEmail]"
        $RunCmd = 'New-DistributionGroup -Name ' + "`"$ShadowGroupName`"" + ' -DisplayName ' + "`"$ShadowGroupName`"" + ' -Alias ' + "`"$ShadowGroupAlias`"" + ' -PrimarySmtpAddress ' + "`"$ShadowGroupEmail`""
        If ($Group.RecipientTypeDetails -match "security") { $RunCmd = $RunCmd + ' -Type Security' } Else { $RunCmd = $RunCmd + ' -Type Distribution' }
        If ($Group.RecipientTypeDetails -eq "RoomList") { $RunCmd = $RunCmd + ' -RoomList' }
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
        If ($Debug -eq $false) { Write-Log "INFO: Hide shadow group from address lists."; Set-DistributionGroup $ShadowGroupEmail -HiddenFromAddressListsEnabled $true -WarningAction SilentlyContinue }
    }
}

# -- Add shadow group members

If ($AddShadowGroupMembers) {
    ForEach ($Member in $MailGroupMembers) {
        Start-Sleep -Milliseconds $MicroDelay
        If ($Member.PrimarySmtpAddress -eq "") { Write-Log "ERROR: Account not found or is not mail-enabled. [$($Member.Name)]"; Continue }
        $ShadowGroupEmail = "$ShadowGroupPrefix" + "$($Member.ParentGroupEmail)"
        If ($LastGroup -ne $Member.ParentGroupEmail) { Write-Log ""; Write-Log "INFO: Add mail group members. [$ShadowGroupEmail]"; $LastGroup = $Member.ParentGroupEmail }
        If ($Member.RecipientTypeDetails -match "group") { $MemberEmail = "$ShadowGroupPrefix" + "$($Member.PrimarySmtpAddress)" } Else { $MemberEmail = $($Member.PrimarySmtpAddress) }
        $RunCmd = 'Add-DistributionGroupMember ' + "`"$ShadowGroupEmail`"" + ' -Member ' + "`"$MemberEmail`"" + ' -WarningAction SilentlyContinue'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
    }
}

# --- Set extended shadow group properties

If (($CreateShadowGroups) -and ($Debug -eq $false)) {
    $GroupParameters = @()
    $GroupParameters += ("AceeptMessagesOnlyFrom", "AcceptedMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers", "BypassModerationFromSendersOrMembers", "BypassNestedModerationEnabled")
    $GroupParameters += ("CustomAttribute1", "CustomAttribute2", "CustomAttribute3", "CustomAttribute4", "CustomAttribute5", "CustomAttribute6", "CustomAttribute7", "CustomAttribute8", "CustomAttribute9", "CustomAttribute10", "CustomAttribute11", "CustomAttribute12", "CustomAttribute13", "CustomAttribute14", "CustomAttribute15")
    $GroupParameters += ("ExtensionCustomAttribute1", "ExtensionCustomAttribute2", "ExtensionCustomAttribute3", "ExtensionCustomAttribute4", "ExtensionCustomAttribute5")
    $GroupParameters += ("MailTip", "ManagedBy", "MemberDepartRestriction", "MemberJoinRestriction", "ModeratedBy", "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers")
    $GroupParameters += ("ReportToManagerEnabled", "ReportToOriginatorEnabled", "RequireSenderAuthenticationEnabled", "SendModerationNotifications", "SendOOFMessageToOriginatorEnabled" )
    ForEach ($Group in $MailGroups) {
        $ShadowGroupEmail = "$ShadowGroupPrefix" + "$($Group.PrimarySmtpAddress)"
        Write-Log ""; Write-Log "INFO: Set extended mail group properties. [$ShadowGroupEmail]"
        ForEach ($Property in $GroupParameters) {
            Start-Sleep -Milliseconds $MicroDelay
            $PropertyValue = "$($Group.$Property)"; If ($PropertyValue -eq "") { Continue }
            $SplitValue = $PropertyValue -split ";"; $PropertyValue = @()
            ForEach ($Value in $SplitValue) {
                If ($Value -eq "TRUE" -or $Value -eq "FALSE") { $PropertyValue += '$'+"$Value" }
                ElseIf ($Value -match "/") { $PropertyValue += "`"$(($Value -split "/")[-1])`"" }
                Else { $PropertyValue += "`"$Value`"" } }
            If ($PropertyValue.Count -eq 1) {
                $RunCmd = 'Set-DistributionGroup ' + "`"$ShadowGroupEmail`"" + ' -' + "$Property $PropertyValue" + ' -WarningAction SilentlyContinue'
                Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
                Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: Account not found or is not mail-enabled. [$Value]" } }
            If ($PropertyValue.Count -gt 1) {
                ForEach ($Value in $PropertyValue) {
                    $RunCmd = 'Set-DistributionGroup ' + "`"$ShadowGroupEmail`"" + ' -' + "$Property" + ' @{Add=' + "$Value" + '} -WarningAction SilentlyContinue'
                    Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
                    Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: Account not found or is not mail-enabled [$Value]" } } } }
    }
}

# --- Cutover shadow groups

If ($CutoverShadowGroups) {
    ForEach ($Group in $MailGroups) {
        Start-Sleep -Milliseconds $MicroDelay
        $GroupName = "$NewGroupNamePrefix" + "$($Group.Name)" + "$NewGroupNameSuffix"
        $ShadowGroupEmail = "$ShadowGroupPrefix" + "$($Group.PrimarySmtpAddress)"
        Write-Log ""; Write-Log "INFO: Cutover shadow group. [$ShadowGroupEmail]"
        If ((Get-MailContact $Group.PrimarySmtpAddress -ErrorAction SilentlyContinue) -ne $null) {
            Write-Log "INFO: Delete existing contact."
            $RunCmd = 'Remove-MailContact ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -Confirm:$false'
            Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
            Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
        $RunCmd = 'Set-DistributionGroup ' + "`"$ShadowGroupEmail`"" + ' -Name ' + "`"$GroupName`"" + ' -DisplayName ' + "`"$GroupName`"" + ' -Alias ' + "`"$($Group.Alias)`"" + ' -PrimarySmtpAddress ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -HiddenFromAddressListsEnabled $' + "$($Group.HiddenFromAddressListsEnabled)" 
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
    }
}

# --- Set alternate proxy addresses

If (($CutoverShadowGroups) -and ($Debug -eq $false)) {
    $AcceptedDomains = (Get-AcceptedDomain).Name
    ForEach ($Group in $MailGroups) {
        Start-Sleep -Milliseconds $MicroDelay
        Write-Log ""; Write-Log "INFO: Set alternate proxy addresses. [$($Group.PrimarySmtpAddress)]"
        $SourceAddresses = @(); $SourceAddresses = ($Group.EmailAddresses -split ";")
        $TargetAddresses = @(); $TargetAddresses = (Get-DistributionGroup $Group.PrimarySmtpAddress).EmailAddresses
        ForEach ($Proxy in $SourceAddresses) {
            If ($Proxy -notlike "smtp:*" -and $Proxy -notlike "x500:*") { Continue }
            If ($Proxy -like "smtp:*") {
                $CheckDomain = ($Proxy.Split("@"))[1]
                If ($AcceptedDomains -notcontains $CheckDomain) { Continue } }            If ($TargetAddresses -notcontains $Proxy) {
                $RunCmd = 'Set-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -EmailAddresses @{Add="' + "$Proxy" + '"}'
                Write-Log "DEBUG: $RunCmd"; Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }
        ForEach ($Proxy in $TargetAddresses) {
            If ($Proxy -like "$ShadowGroupPrefix*") {
                $RunCmd = 'Set-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -EmailAddresses @{Remove="' + "$Proxy" + '"}'
                Write-Log "DEBUG: $RunCmd"; Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }
    }
}

# --- Audit cutover shadow groups

If (($CutoverShadowGroups) -and  ($Debug -eq $false)) {
    Write-Log ""; Write-Log "INFO: Audit cutover mail groups..."
    $MailGroupsExport = $Timestamp + "_MigratedShadowGroups.csv"
    $Select = @(); Get-DistributionGroup -ResultSize 1 -WarningAction SilentlyContinue | Get-Member -MemberType Property | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
    ForEach ($Group in $MailGroups) {     
        $RunCmd = 'Get-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -WarningAction SilentlyContinue | Select ' + $($Select -join ",")
        Invoke-Expression $RunCmd | Export-Csv "$MailGroupsExport" -NoTypeInformation -Append }
}

# --- Purge shadow groups

If ($PurgeShadowGroups) {
    ForEach ($Group in $MailGroups) {
        Start-Sleep -Milliseconds $MicroDelay
        $ShadowGroupEmail = "$ShadowGroupPrefix" + "$($Group.PrimarySmtpAddress)"
        If ((Get-DistributionGroup $ShadowGroupEmail -ErrorAction SilentlyContinue) -ne $null) {
            Write-Log "INFO: Purge shadow group. [$ShadowGroupEmail]"
            $RunCmd = 'Remove-DistributionGroup ' + "`"$ShadowGroupEmail`"" + ' -Confirm:$false'
            Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
            Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
    }
}

# --- Script complete

Write-Log ""; Write-Log "SUCCESS: Script complete."