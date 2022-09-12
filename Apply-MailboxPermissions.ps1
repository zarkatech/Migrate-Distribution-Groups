$ScriptInfo = @"
================================================================================
Apply-MailboxPermissions.ps1 | v3.2
by Roman Zarka
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$IncludeMailboxAccess = $false
    $MailboxAccessImportFile = "190827100349_MailboxAccess.csv"
$IncludeSendAs = $false
    $SendAsImportFile = "190827100349_MailboxSendAs.csv"
$IncludeSendOnBehalf = $false
    $SendOnBehalfImportFile = "190827100349_MailboxSendOnBehalf.csv"
$IncludeFolderDelegates = $false
    $FolderDelegatesImportFile = "190827100349_MailboxFolderDelegates.csv"
$IncludeMailboxForwarding = $true
    $MailboxForwardingImportFile = "DistrictABC_10070127_MailboxForwarding.csv"
$ApplyGroupPermissionsOnly = $true
$Debug = $true

# --- Initialize log file

$Timestamp = Get-Date -Format MMddhhmm
$RunLog = $Timestamp+"_ApplyCloudMailboxPermissions.log"
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
        Write-Host $LogString -ForegroundColor DarkGray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "") {
        Write-Host ""
        Write-Output "`n" | Out-File $RunLog -Append }
}

# --- Validate import files

If (($IncludeMailboxAccess) -and (Test-Path "$MailboxAccessImportFile") -eq $false) { Write-Log "ERROR: Mailbox Acess import file not found. [$MailboxAccessImportFile]"; Break }
If (($IncludeSendAs) -and (Test-Path "$SendAsImportFile") -eq $false) { Write-Log "ERROR: Send As import file not found. [$SendAsImportFile]"; Break }
If (($IncludeSendOnBehalf) -and (Test-Path "$SendOnBehalfImportFile") -eq $false) { Write-Log "ERROR: Send On Behalf import file not found. [$SendOnBehalfImportFile]"; Break }
If (($IncludeFolderDelegates) -and (Test-Path "$FolderDelegatesImportFile") -eq $false) { Write-Log "ERROR: Folder Delelegates import file not found. [$FolderDelegatesImportFile]"; Break }
If (($IncludeMailboxForwarding) -and (Test-Path "$MailboxForwardingImportFile") -eq $false) { Write-Log "ERROR: Folder Delelegates import file not found. [$MailboxForwardingImportFile]"; Break }
If ($Debug) { Write-Log "ALERT: Script is in debug mode and changes will NOT be committed." }
Else { Write-Log "ALERT: Script is NOT in debug mode and changes will be commited." }
$Continue = Read-Host "Continue? [Y]es or [N]o"; If ($Continue -ne "Y") { Break }

# --- Apply mailbox access permissions

If ($IncludeMailboxAccess -eq $true) {
    Write-Log ""; Write-Log "INFO: Apply mailbox access permissions..."
    If ($ApplyGroupPermissionsOnly) { $Delegates = (Import-Csv "$MailboxAccessImportFile" | Where {$_.DelegateType -like "*Group"}) }
    Else {$Delegates = Import-Csv "$MailboxAccessImportFile" }
    ForEach ($Delegate in $Delegates) {
        $SetCmd = 'Add-MailboxPermission '+"`"$($Delegate.MailboxEmail)`""+' -User '+"`"$($Delegate.DelegateEmail)`""+' -AccessRights '+"`"$($Delegate.DelegateAccess)`""
        If ($Debug) { Write-Log "DEBUG: $SetCmd"; $SetCmd = $SetCmd+' -WhatIf' }
        Invoke-Expression $SetCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }

# --- Apply SendAs permissions

If ($IncludeSendAs -eq $true) {
    Write-Log ""; Write-Log "INFO: Apply SendAs permissions..."
    If ($ApplyGroupPermissionsOnly) { $Delegates = (Import-Csv "$SendAsImportFile" | Where {$_.DelegateType -like "*Group"}) }
    Else { $Delegates = Import-Csv "$SendAsImportFile" }
    ForEach ($Delegate in $Delegates) {
        $SetCmd = 'Add-RecipientPermission '+"`"$($Delegate.MailboxEmail)`""+' -Trustee '+"`"$($Delegate.DelegateEmail)`""+' -AccessRights SendAs -Confirm:$false'
        If ($Debug) { Write-Log "DEBUG: $SetCmd"; $SetCmd = $SetCmd+' -WhatIf' }
        Invoke-Expression $SetCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }

# --- Apply SendOnBehalf permissions

If ($IncludeSendOnBehalf -eq $true) {
    Write-Log ""; Write-Log "INFO: Apply SendOnBehalf permissions..."
    If ($ApplyGroupPermissionsOnly) { $Delegates = (Import-Csv "$SendOnBehalfImportFile" | Where {$_.DelegateType -like "*Group"}) }
    Else { $Delegates = Import-Csv "$SendOnBehalfImportFile" }
    ForEach ($Delegate in $Delegates) {
        $SetCmd = 'Set-Mailbox '+"`"$($Delegate.MailboxEmail)`""+' -GrantSendOnBehalfTo @{Add='+"`"$($Delegate.DelegateEmail)`""+'}'
        If ($Debug) { Write-Log "DEBUG: $SetCmd"; $SetCmd = $SetCmd+' -WhatIf' }
        Invoke-Expression $SetCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }

# --- Apply folder delegate permissions

If ($IncludeFolderDelegates -eq $true) {
    Write-Log ""; Write-Log "INFO: Apply folder delegate permissions..."
    If ($ApplyGroupPermissionsOnly) { $Delegates = (Import-Csv "$FolderDelegatesImportFile" | Where {$_.DelegateType -like "*Group"}) }
    Else { $Delegates = Import-Csv "$FolderDelegatesImportFile" }
    ForEach ($Delegate in $Delegates) {
        $SetCmd = 'Add-MailboxFolderPermission '+"`"$($Delegate.FolderLocation)`""+' -User '+"`"$($Delegate.DelegateEmail)`""+' -AccessRights '+"`"$($Delegate.DelegateAccess)`""
        If ($Debug) { Write-Log "DEBUG: $SetCmd"; $SetCmd = $SetCmd+' -WhatIf' }
        Invoke-Expression $SetCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } } }

# --- Apply mailbox forwarding

If ($IncludeMailboxForwarding -eq $true) {
    Write-Log ""; Write-Log "INFO: Apply mailbox forwarding..."
    $Delegates = (Import-Csv "$MailboxForwardingImportFile")
    ForEach ($Delegate in $Delegates) {
        If ($Delegate.ForwardingEmail -like "smtp:*") { $SetCmd = 'Set-Mailbox '+"$($Delegate.MailboxEmail)"+' -ForwardingSmtpAddress '+"$($Delegate.ForwardingEmail)"+' -DeliverToMailboxAndForward $'+"$($Delegate.DeliverToMailbox)" }
        Else { $SetCmd = 'Set-Mailbox '+"$($Delegate.MailboxEmail)"+' -ForwardingAddress '+"$($Delegate.ForwardingEmail)"+' -DeliverToMailboxAndForward $'+"$($Delegate.DeliverToMailbox)" }
        If ($Debug) { Write-Log "DEBUG: $SetCmd"; $SetCmd = $SetCmd+' -WhatIf' }
        Invoke-Expression $SetCmd -ErrorVariable CmdError; If ($CmdError -ne "") { Write-Log "ERROR: $CmdError" } }
}

# --- Script complete

Write-Log ""; Write-Log "SUCCESS: Script complete."
