$ScriptInfo = @"
================================================================================
Create-SourceContacts.ps1 | v3.2.1
by Roman Zarka | Microsoft Services
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$MailGroupsAuditFile = "DistrictABC_10060351_MigratedMailGroups.csv"
$TargetOU = 'OU=NoSync,DC=rjzlab,DC=local'
$UseNoSync = $true
    $NoSyncAttribute = 'CustomAttribute11'
    $NoSyncValue = 'NoSyncO365'
$DeleteDistributionGroups = $true
$DisableMailSecurityGroups = $false
$Debug = $true

# --- Initialize log file

$Timestamp = Get-Date -Format MMddhhmm
$RunLog = $Timestamp + "_CreateSourceContacts.log"
If ($Debug) { $RunLog = "DEBUG_" + $RunLog }
If ($BatchName -ne "" -and $BatchName -ne $null) { $RunLog = $BatchName + "_" + $RunLog }
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
If (Test-Path $MailGroupsAuditFile) { $MailGroupsAudit = Import-Csv $MailGroupsAuditFile }
Else { Write-Log "ERROR: Mail groups audit export not found. [$MailGroupsAuditFile]"; Break }
Write-Log "SUCCESS: Found $($MailGroupsAudit.Count) mail groups."
If ($Debug) { Write-Log "ALERT: Script is in debug mode and changes will NOT be committed." }
Else { Write-Log "ALERT: Script is NOT in debug mode and changes will be commited." }
$Continue = Read-Host "Continue? [Y]es or [N]o"; If ($Continue -ne "Y") { Break }

# --- Create mail group contacts

ForEach ($Group in $MailGroupsAudit) {
    If (($DeleteDistributionGroups) -and ($Group.RecipientTypeDetails -notmatch "security")) {
        Write-Log "INFO: Delete distribution group. [$($Group.PrimarySmtpAddress)]"
        $RunCmd = 'Remove-DistributionGroup ' + "$($Group.PrimarySmtpAddress)" + ' -Confirm:$false'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }  
    If (($DisableMailSecurityGroups) -and ($Group.RecipientTypeDetails -match "security")) {
        Write-Log "INFO: Disable mail-enabled security group. [$($Group.PrimarySmtpAddress)]"
        $RunCmd = 'Disable-DistributionGroup ' + "$($Group.PrimarySmtpAddress)" + ' -Confirm:$false'
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
    $EmailAddresses = $Group.EmailAddresses -split ";"
    ForEach ($Proxy in $EmailAddresses) {
        If ($Proxy -like "*.mail.onmicrosoft.com") { $TargetProxy = $Proxy}
        If ($Proxy -like "*.onmicrosoft.com" -and $TargetProxy -notlike "*.mail.onmicrosoft.com") { $TargetProxy = $Proxy } }
    Write-Log "INFO: Create mail group contact. [$($Group.PrimarySmtpAddress)]"
    $RunCmd = 'New-MailContact -Name ' + "`"$($Group.DisplayName)`"" + ' -Alias ' + "$($Group.Alias)" + ' -PrimarySmtpAddress ' + "$($Group.PrimarySmtpAddress)" + ' -ExternalEmailAddress ' + $TargetProxy + ' -OrganizationalUnit ' + "`"$TargetOU`""
    Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
    Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" }
    If ($Debug -eq $false) { Set-MailContact $Group.PrimarySmtpAddress -EmailAddresses @{Add="$TargetProxy"} }
}

# --- Set NoSync attribute

If ($UseNoSync -eq $true) {
    ForEach ($Group in $MailGroupsAudit) {
        Write-Log "INFO: Set NoSync attribute. [$($Group.PrimarySmtpAddress)]"
        $RunCmd = 'Set-MailContact ' + "$($Group.PrimarySmtpAddress)" + ' -' + "$NoSyncAttribute ""$NoSyncValue"""
        Write-Log "DEBUG: $RunCmd"; If ($Debug) { $RunCmd = $RunCmd + ' -WhatIf' }
        Invoke-Expression $RunCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
}

# --- Script complete
Write-Log "SUCCESS: Script complete."