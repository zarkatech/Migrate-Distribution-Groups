$ScriptInfo = @"
================================================================================
Deprovision-MailGroups.ps1 | v3.2
by Roman Zarka | github.com/zarkatech
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$MailGroupsAuditFile = "10060103_DistrictABC_AuditMailGroups.csv"
$UseNoSync = $true
    $NoSyncAttribute = "CustomAttribute11"
    $NoSyncValue = "NoSyncO365"
$DeleteDistributionGroups = $false
$DisableMailSecurityGroups = $false
$Debug = $true

# --- Initialize log file

$TimeStamp = Get-Date -Format MMddhhmm
If ($BatchName -ne "" -and $BatchName -ne $null) { $TimeStamp = $TimeStamp + "_$BatchName" }
$RunLog = $Timestamp + "_DeprovisionMailGroups.log"
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
        Write-Host $LogString -ForegroundColor Gray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "" -or $LogStatus -eq $null) {
        Write-Host ""
        Write-Output "`n" | Out-File $RunLog -Append }
}

# --- Verify mail groups migration data

Write-Log "INFO: Verify mail groups migration data..."
If (Test-Path $MailGroupsAuditFile) { $MailGroupsAudit = (Import-Csv $MailGroupsAuditFile | Select PrimarySmtpAddress, RecipientTypeDetails) }
Else { Write-Log "ERROR: Mail groups audit export not found. [$MailGroupsAuditFile]"; Break }
Write-Log "SUCCESS: Found $($MailGroupsAudit.Count) mail groups."
If ($Debug) { Write-Log "ALERT: Script is in debug mode and changes will NOT be committed." }
Else { Write-Log "ALERT: Script is NOT in debug mode and changes will be commited." }
$Continue = Read-Host "Continue? [Y]es or [N]o"; If ($Continue -ne "Y") { Break }

# --- Deprovision source mail groups

ForEach ($Group in $MailGroupsAudit) {
    Write-Log ""; Write-Log "INFO: Deprovision mail group. [$($Group.PrimarySmtpAddress)]"
    If ($UseNoSync -eq $true) {
        Write-Log "INFO: Set NoSync attribute."
        $RunCmd = 'Set-DistributionGroup ' + "$($Group.PrimarySmtpAddress)" + ' -' + "$NoSyncAttribute ""$NoSyncValue"""
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
    If (($DeleteDistributionGroups) -and ($Group.RecipientTypeDetails -notmatch "security")) {
        Write-Log "INFO: Delete distribution group."
        $RunCmd = 'Remove-DistributionGroup ' + "$($Group.PrimarySmtpAddress)" + ' -Confirm:$false'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }  
    If (($DisableMailSecurityGroups) -and ($Group.RecipientTypeDetails -match "security")) {
        Write-Log "INFO: Disable mail-enabled security group."
        $RunCmd = 'Disable-DistributionGroup ' + "$($Group.PrimarySmtpAddress)" + ' -Confirm:$false'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
}

# --- Script complete

Write-Log ""; Write-Log "SUCCESS: Script complete."
If ($Debug -eq $false) { Write-Log "ALERT: Manually trigger or wait for AAD Connect synchroniztion before proceeding." }
