$ScriptInfo = @"
================================================================================
Create-TargetMailGroups.ps1 | v3.2.2
by Roman Zarka | Microsoft Services
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$MailGroupsAuditFile = "10060103_DistrictABC_AuditMailGroups.csv"
$MailGroupMembersAuditFile = "10060103_DistrictABC_AuditMailGroupMembers.csv"
$SetRbacScope = $true
    $RbacAttribute = 'CustomAttribute1'
    $RbacValue = 'USA'
$GroupNamePrefix = ''
$GroupNameSuffix = ' (USA)'
$Debug = $true

# --- Initialize log file

$Timestamp = Get-Date -Format MMddhhmm
If ($BatchName -ne "" -and $BatchName -ne $null) { $Timestamp = $Timestamp + "_$BatchName" }
$RunLog = $Timestamp + "_CreateTargetMailGroups.log"
If ($Debug) { $RunLog = "DEBUG_" + $RunLog }
Function Write-Log ($LogString) {
    $LogStatus = $LogString.Split(":")[0]
    If ($LogStatus -eq "ALERT") {
        Write-Host $LogString -ForegroundColor Yellow
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
        Write-Output "`n" | Out-File $RunLog -Append }
}

# --- Initialize script environment

If ([int]$PsVersionTable.PSVersion.Major -lt 3) { Write-Log "ERROR: Script requires PowerShell v3 or later and must be run from an upgraded console."; Break }
If ((Get-PSSession) -eq $null -or ((Get-PSSession).ConfigurationName) -ne "Microsoft.Exchange") { Write-Log "ERROR: Script must be run from an Exchange session."; Break }
Write-Log ""; Write-Log "INFO: Verify mail groups audit data..."
If (Test-Path $MailGroupsAuditFile) { $MailGroups = Import-Csv $MailGroupsAuditFile | Sort Name }
Else { Write-Log "ERROR: Mail groups audit file not found. [$MailGroupsAuditFile]"; Break }
If (Test-Path $MailGroupMembersAuditFile) { $MailGroupMembers = (Import-Csv $MailGroupMembersAuditFile | Sort ParentGroupEmail) }
Else { Write-Log "ERROR: Mail group members audit file not found. [$MailGroupMembersAuditFile]"; Break }
$AcceptedDomains = (Get-AcceptedDomain).Name
If ($Debug) { Write-Log "ALERT: Script is in debug mode and changes will NOT be committed." }
Else { Write-Log "ALERT: Script is NOT in debug mode and changes will be commited." }
$Continue = Read-Host "Continue? [Y]es or [N]o"; If ($Continue -ne "Y") { Break }

# --- Stage new mail groups

ForEach ($Group in $MailGroups) {
    Write-Log ""; Write-Log "INFO: Stage new mail group. [$($Group.PrimarySmtpAddress)]"
    If ((Get-MailContact $Group.PrimarySmtpAddress -ErrorAction SilentlyContinue) -ne $null) {
        Write-Log "INFO: Delete existing contact."
        $RunCmd = 'Remove-MailContact ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -Confirm:$false'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
    $GroupName = "$GroupNamePrefix" + "$($Group.Name)" + "$GroupNameSuffix"
    If (($GroupName | Measure -Character).Characters -gt 64) { Write-Log "ERROR: Group name exceeds maximum characters. [$GroupName]"; Continue }
    $RunCmd = 'New-DistributionGroup -Name ' + "`"$GroupName`"" + ' -DisplayName ' + "`"$GroupName`"" + ' -Alias ' + "$($Group.Alias)" + ' -PrimarySmtpAddress ' + "`"$($Group.PrimarySmtpAddress)`""
    If ($Group.RecipientTypeDetails -match "security") { $RunCmd = $RunCmd + ' -Type Security' }
    Else { $RunCmd = $RunCmd + ' -Type Distribution' }
    If ($Group.RecipientTypeDetails -eq "RoomList") { $RunCmd = $RunCmd + ' -RoomList' }
    Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
    Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
}

# -- Add mail group members

ForEach ($Member in $MailGroupMembers) {
    If ($LastGroup -ne $Member.ParentGroupEmail) { Write-Log ""; Write-Log "INFO: Add mail group members. [$($Member.ParentGroupEmail)]"; $LastGroup = $Member.ParentGroupEmail }
    $RunCmd = 'Add-DistributionGroupMember ' + "`"$($Member.ParentGroupEmail)`"" + ' -Member ' + "`"$($Member.PrimarySmtpAddress)`""
    Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
    If ($Member.PrimarySmtpAddress -eq "") { Write-Log "ALERT: Account not found or is not mail-enabled. [$($Member.Name)]"; Continue }
    Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
}

# --- Set alternate proxy addresses

ForEach ($Group in $MailGroups) {
    Write-Log ""; Write-Log "INFO: Set alternate proxy addresses. [$($Group.PrimarySmtpAddress)]"
    $SourceAddresses = @(); $SourceAddresses = ($Group.EmailAddresses -split ";")
    $TargetAddresses = @(); $TargetAddresses = (Get-DistributionGroup $Group.PrimarySmtpAddress).EmailAddresses
    ForEach ($Proxy in $SourceAddresses) {
        If ($Proxy -notlike "smtp:*" -and $Proxy -notlike "x500:*") { Continue }
        If ($Proxy -like "smtp:*") {
            $CheckDomain = ($Proxy.Split("@"))[1]
            If ($AcceptedDomains -notcontains $CheckDomain) { Continue } }        If ($TargetAddresses -notcontains $Proxy) {
            $RunCmd = 'Set-DistributionGroup '+"`"$($Group.PrimarySmtpAddress)`""+' -EmailAddresses @{Add="'+"$Proxy"+'"}'
            Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
            Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }
}

# --- Set extended mail group properties

$GroupParameters = @()
$GroupParameters += ("AceeptMessagesOnlyFrom", "AcceptedMessagesOnlyFromDLMembers", "AcceptMessagesOnlyFromSendersOrMembers", "BypassModerationFromSendersOrMembers", "BypassNestedModerationEnabled")
$GroupParameters += ("CustomAttribute1", "CustomAttribute2", "CustomAttribute3", "CustomAttribute4", "CustomAttribute5", "CustomAttribute6", "CustomAttribute7", "CustomAttribute8", "CustomAttribute9", "CustomAttribute10", "CustomAttribute11", "CustomAttribute12", "CustomAttribute13", "CustomAttribute14", "CustomAttribute15")
$GroupParameters += ("ExtensionCustomAttribute1", "ExtensionCustomAttribute2", "ExtensionCustomAttribute3", "ExtensionCustomAttribute4", "ExtensionCustomAttribute5")
$GroupParameters += ("Hidden", "MailTip", "ManagedBy", "MemberDepartRestriction", "MemberJoinRestriction", "ModeratedBy", "RejectMessagesFrom", "RejectMessagesFromDLMembers", "RejectMessagesFromSendersOrMembers")
$GroupParameters += ("ReportToManagerEnabled", "ReportToOriginatorEnabled", "RequireSenderAuthenticationEnabled", "SendModerationNotifications", "SendOOFMessageToOriginatorEnabled" )
ForEach ($Group in $MailGroups) {
    Write-Log ""; Write-Log "INFO: Set extended mail group properties. [$($Group.PrimarySmtpAddress)]"
    ForEach ($Property in $GroupParameters) {
        $PropertyValue = "$($Group.$Property)"; If ($PropertyValue -eq "") { Continue }
        $SplitValue = $PropertyValue -split ";"; $PropertyValue = @()
        ForEach ($Value in $SplitValue) {
            If ($Value -eq "TRUE" -or $Value -eq "FALSE") { $PropertyValue += '$'+"$Value" }
            ElseIf ($Value -match "/") { $PropertyValue += "`"$(($Value -split "/")[-1])`"" }
            Else { $PropertyValue += "`"$Value`"" } }
        If ($PropertyValue.Count -eq 1) {
            $RunCmd = 'Set-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -' + "$Property $PropertyValue"
            Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
            Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ALERT: Account not found or is not mail-enabled. [$Value]" } }
        If ($PropertyValue.Count -gt 1) {
            ForEach ($Value in $PropertyValue) {
                $RunCmd = 'Set-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -' + "$Property" + ' @{Add='+"$Value"+'}'
                Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
                Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ALERT: Account not found or is not mail-enabled [$Value]" } } } }
}

# --- Audit migrated mail groups

If ($Debug -eq $false) {
    Write-Log ""; Write-Log "INFO: Audit migrated mail groups..."
    $MailGroupsExport = $Timestamp + "_MigratedMailGroups.csv"
    ForEach ($Group in $MailGroups) {     
        $Select = @(); Get-DistributionGroup -ResultSize 1 | Get-Member | Where { $_.MemberType -eq "Property" } | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
        $RunCmd = 'Get-DistributionGroup ' + "`"$($Group.PrimarySmtpAddress)`"" + ' -WarningAction SilentlyContinue | Select ' + $($Select -join ",")
        Invoke-Expression $RunCmd | Export-Csv "$MailGroupsExport" -NoTypeInformation -Append } }

# --- Script complete

Write-Log ""; Write-Log "SUCCESS: Script complete."
If ($Debug -eq $false) { Write-Log "ALERT: Run script to re-apply mailbox permissions using newly created groups." }
