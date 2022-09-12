$ScriptInfo = @"
================================================================================
Audit-MailboxPermissions.ps1 | v3.3.1
by Roman Zarka | Microsoft Services
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$UseImportFile = $false
    $ImportFile = "UserList.csv"
$UseFilterCriteria = $false
    $FilterBatchName = "DistrictABC"
    $FilterCriteria = '(CustomAttribute1 -eq "USA") -and (CustomAttribute3 -eq "IT")'
$IncludeMailboxAccess = $true
$IncludeSendAs = $true
$IncludeSendOnBehalf = $true
$IncludeFolderDelegates = $true
    $IncludeCommonFoldersOnly = $true
$IncludeMailboxForwarding = $true
$DelegatesToSkip = "NT AUTHORITY\SELF", "SELF", "DOMAIN\BESADMIN","DOMAIN\Administrators"
$ExpandSecurityGroups = $false
$ExpandDistributionGroups = $false
$MicroDelaySleep = 0

# --- Initialize log file

$TimeStamp = Get-Date -Format MMddhhmm
If ($UseImportFile) { $TimeStamp = $TimeStamp + "_" + $($ImportFile.Split(".")[0]) }
If ($UseFilterCriteria) { $TimeStamp = $TimeStamp + "_" + $FilterBatchName }
$RunLog = $TimeStamp + "_AuditMailboxPermissions.log"
Get-PSSession | fl | Out-File $RunLog
Function Write-Log ($LogString) {
    $LogStatus = $LogString.Split(":")[0]
    If ($LogStatus -eq "SUCCESS") {
        Write-Host $LogString -ForegroundColor Green
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "INFO") {
        Write-Host "$LogString" -ForegroundColor Cyan
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ALERT") {
        Write-Host $LogString -ForegroundColor Yellow
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ERROR") {
        Write-Host $LogString -BackgroundColor Red
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "AUDIT") { Write-Host $LogString -ForegroundColor DarkGray } }

# --- Initialize script environment

If ($UseImportFile -eq $true -and (Test-Path $ImportFile) -eq $false) { Write-Log "ERROR: User list import file not found. [$ImportFile]"; Break }
If ($UseImportFile -eq $true -and $UseFilterCriteria -eq $true) { Write-Log "ERROR: Import file and filter criteria cannot be used together."; Break }
If ((Get-PSSession) -eq $null -or ((Get-PSSession).ConfigurationName) -ne "Microsoft.Exchange") { Write-Log "ERROR: Script must be run from a connected Exchange session."; Break }
If (Get-Command "Get-ADPermission" -ErrorAction SilentlyContinue) { $SourceIsEXO = $false } Else { $SourceIsEXO = $true }
If (($SourceIsExo) -and (Get-ExoMailbox -ResultSize 1 -ErrorAction SilentlyContinue) -ne $null) { $SourceIsRest = $true } Else { $SourceIsRest = $false }
If ($SourceIsEXO -eq $false) {
    $EmsVersion = Get-PSSnapin -Registered | Where { $_.Name -like "*Exchange*" }
    If ($EmsVersion -eq $null) { Write-Log "ERROR: On-premises Exchange session not detected."; Break }
    If ($EmsVersion -like "*Powershell.Admin*") {
        $AdminSessionADSettings.ViewEntireForest = $true
        If ($IncludeFolderDelegates -eq $true) { Write-Log "ALERT: Folder delegate permissions cannot be retrieved from Exchange 2007 and will not be exported."; $IncludeFolderDelegates = $false } }
    Else { Set-AdServerSettings -ViewEntireForest $true }
    $SessionVersion = (Get-ExchangeServer (Get-PSSession).ComputerName).AdminDisplayVersion.Major
    Write-Log "INFO: Script session console detected as Exchange version $SessionVersion."
    $ExchVersions = (Get-ExchangeServer | Where { $_.ServerRole -like "*Mailbox*" }).AdminDisplayVersion.Major | Select -Unique
    If ($ExchVersions -gt 1 -and $IncludeFolderDelegates -eq $true) { Write-Log "ALERT: Folder delegate permissions cannot be audited across different versions of Exchange." } }

# --- Initialize export files

If ($IncludeMailboxAccess) { $MailboxAccessExport = $Timestamp + "_MailboxAccess.csv"; "`"MailboxEmail`",`"DelegateEmail`",`"DelegateType`",`"DelegateAccess`"" | Out-File $MailboxAccessExport -Encoding ascii }
If ($IncludeSendAs) { $SendAsExport = $Timestamp + "_MailboxSendAs.csv"; "`"MailboxEmail`",`"DelegateEmail`",`"DelegateType`",`"DelegateAccess`"" | Out-File $SendAsExport -Encoding ascii }
If ($IncludeSendOnBehalf) { $SendOnBehalfExport = $Timestamp + "_MailboxSendOnBehalf.csv"; "`"MailboxEmail`",`"DelegateEmail`",`"DelegateType`",`"DelegateAccess`"" | Out-File $SendOnBehalfExport -Encoding ascii }
If ($IncludeFolderDelegates) { $FolderDelegateExport = $Timestamp + "_MailboxFolderDelegates.csv"; "`"MailboxEmail`",`"FolderLocation`",`"DelegateEmail`",`"DelegateType`",`"DelegateAccess`"" | Out-File $FolderDelegateExport -Encoding ascii }
If ($IncludeMailboxForwarding) { $MailboxForwardingExport = $Timestamp + "_MailboxForwarding.csv"; "`"MailboxEmail`",`"ForwardingEmail`",`"DeliverToMailbox`"" | Out-File $MailboxForwardingExport -Encoding ascii }

# --- Initialize Check-Delegates function

Function Check-Delegates ([string]$DelegateUser, $ExportFile) {
    Start-Sleep -Milliseconds $MicroDelaySleep
    If ($DelegateUser -like "*\*") { $DelegateUser = $DelegateUser.Split("\")[1] }
    $CheckDelegate = Get-Recipient $DelegateUser -ErrorAction SilentlyContinue
    If ($CheckDelegate.Count -gt 1) { Write-Log "ALERT: Delegate user matches multiple accounts. [$DelegateUser]" }
    If ($CheckDelegate -eq $null) {
        $CheckDelegate = Get-Group $DelegateUser -ErrorAction SilentlyContinue }
    If ($CheckDelegate -ne $null) {
        If (($CheckDelegate.RecipientType -like "Mail*" -and $ExpandDistributionGroups -eq $false) -or $CheckDelegate.RecipientType -like "*Mailbox") {
            $DelegateEmail = $CheckDelegate.PrimarySmtpAddress
            $DelegateType = $CheckDelegate.RecipientTypeDetails
            "`"$MailboxEmail`",`"$DelegateEmail`",`"$DelegateType`",`"$DelegateAccess`"" | Out-File $ExportFile -Encoding ascii -Append }
        If ($CheckDelegate.RecipientType -like "Mail*" -and $CheckDelegate.RecipientType -like "*Group" -and $ExpandDistributionGroups -eq $true) {
            Write-Log "ALERT: Expand distribution group membership. [$($CheckDelegate.Name)]"
            ForEach ($Member in Get-DistributionGroupMember $CheckDelegate.Name -ResultSize Unlimited) {
                $CheckMember = Get-Recipient $Member -ErrorAction SilentlyContinue
                If ($CheckMember -ne $null) {
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    $DelegateType = $CheckMember.RecipientTypeDetails
                    "`"$MailboxEmail`",`"$DelegateEmail`",`"$DelegateType`",`"$DelegateAccess`"" | Out-File $ExportFile -Encoding ascii -Append } } }
        If ($CheckDelegate.RecipientType -eq "Group" -and $ExpandSecurityGroups -eq $true) {
            Write-Log "ALERT: Expand security group membership. [$($CheckDelegate.Name)]"
            ForEach ($Member in (Get-Group $DelegateUser).Members) {
                $CheckMember = Get-Recipient $Member -ErrorAction SilentlyContinue
                If ($CheckMember -ne $null) {
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    $DelegateType = $CheckMember.RecipientTypeDetails
                    "`"$MailboxEmail`",`"$DelegateEmail`",`"$DelegateType`",`"$DelegateAccess`"" | Out-File $ExportFile -Encoding ascii -Append } } } }      
 }

# --- Retrieve mailboxes

If ($SourceIsRest) { Write-Log "INFO: Optimizing script for Exchange Online REST API." }
If ($SourceIsEXO) { Write-Log "INFO: Retrieving mailboxes from Exchange Online..." }
Else { Write-Log "INFO: Retrieving mailboxes from on-premises Exchange..." }
If ($UseImportFile) {
    Write-Log "INFO: Importing user list from file. [$ImportFile]"
    $Mailboxes = @()
    If ($SourceIsRest) { Import-Csv $ImportFile | ForEach { $Mailboxes += (Get-ExoMailbox -Identity $_.PrimarySmtpAddress -Properties PrimarySmtpAddress, DistinguishedName, ExchangeVersion) } }
    Else { Import-Csv $ImportFile | ForEach { $Mailboxes += (Get-Mailbox $_.PrimarySmtpAddress | Select PrimarySmtpAddress, DistinguishedName, AdminDisplayVersion, ExchangeVersion) } } }
Else {
    If ($SourceIsRest) { $RunCmd = 'Get-ExoMailbox -ResultSize Unlimited' }
    Else { $RunCmd = 'Get-Mailbox -ResultSize Unlimited' }
    If (($UseFilterCriteria) -and $FilterCriteria -ne "") { $RunCmd = $RunCmd + ' -Filter {' + "$FilterCriteria" + '}' }
    If ($SourceIsRest) { $RunCmd = $RunCmd + ' -Properties PrimarySmtpAddress, DistinguishedName, ExchangeVersion' }
    Else { $RunCmd = $RunCmd + ' | Select PrimarySmtpAddress, DistinguishedName, AdminDisplayVersion, ExchangeVersion' }
    $Mailboxes = Invoke-Expression $RunCmd }
If ($Mailboxes -eq $null) { Write-Log "ERROR: No mailboxes found."; Break }
$MailboxCount = $Mailboxes.Count; $Progress = 0
Write-Log "SUCCESS: Found $MailboxCount Mailboxes."

# --- Audit mailbox permissions

ForEach ($Mailbox in $Mailboxes) {
    [string]$MailboxEmail = $Mailbox.PrimarySmtpAddress; [string]$MailboxDN = $Mailbox.DistinguishedName
    $Progress = $Progress + 1
    Write-Log "INFO: Audit mailbox $Progress of $MailboxCount. [$MailboxEmail]"

    # --- Export mailbox access permissions

    If ($IncludeMailboxAccess -eq $true) {
        Write-Log "AUDIT: Mailbox access permissions..."
        $Delegates = @()
        If ($SourceIsRest) { $Delegates = (Get-ExoMailboxPermission -Identity $MailboxEmail | Where { $DelegatesToSkip -notcontains $_.User -and $_.IsInherited -eq $false }) }
        Else { $Delegates = (Get-MailboxPermission $MailboxEmail | Where { $DelegatesToSkip -notcontains $_.User -and $_.IsInherited -eq $false }) }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = $Delegate.AccessRights
                Check-Delegates $Delegate.User $MailboxAccessExport } } }

    # --- Export SendAs permissions

    If ($IncludeSendAs -eq $true) {
        Write-Log "AUDIT: Send As permissions..."
        $Delegates = @()
        If ($SourceIsEXO) { $Delegates = Get-RecipientPermission $MailboxEmail | Where { $DelegatesToSkip -notcontains $_.Trustee -and $_.AccessRights -like "SendAs" } }
        Else { $Delegates = Get-ADPermission -Identity $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.ExtendedRights -like "*send-as*" } }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendAs" 
                If ($SourceIsExo) { Check-Delegates $Delegate.Trustee $SendAsExport }
                Else { Check-Delegates $Delegate.User $SendAsExport } } } }

    # --- Export SendOnBehalf permissions

    If ($IncludeSendOnBehalf -eq $true) {
        Write-Log "AUDIT: Send On Behalf permissions..."
        $Delegates = @()
        If ($SourceIsRest) { $Delegates = (Get-ExoMailbox -Identity $MailboxEmail -Properties GrantSendOnBehalfTo).GrantSendOnBehalfTo }
        Else { $Delegates = (Get-Mailbox $MailboxEmail).GrantSendOnBehalfTo }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendOnBehalf"
                If ($SourceIsExo) { Check-Delegates $Delegate $SendOnBehalfExport }
                Else { Check-Delegates $Delegate.Name $SendOnBehalfExport } } } }

    # --- Export folder permissions

    If ($IncludeFolderDelegates -eq $true) {
        Write-Log "AUDIT: Folder delegate permissions..."
        If ($Mailbox.AdminDisplayVersion.Major -ne "") { $MbxVersion = $Mailbox.AdminDisplayVersion.Major }
        If ($Mailbox.ExchangeVersion.ExchangeBuild.Major -ne "") { $MbxVersion = $Mailbox.ExchangeVersion.ExchangeBuild.Major }
        If ($MbxVersion -ne $SessionVersion) { Write-Log "ERROR: Cannot audit folder delegate permissions for Exchange version $MbxVersion mailbox from version $SessionVersion console." } 
        Else {
            If ($IncludeCommonFoldersOnly -eq $true) {
                If ($SourceIsRest) { $Folders = (Get-ExoMailboxFolderStatistics -Identity $MailboxEmail | Where { $_.FolderPath -eq "/Top of Information Store" -or $_.FolderPath -eq "/Inbox" -or $_.FolderPath -eq "/Calendar" }) }
                Else { $Folders = Get-MailboxFolderStatistics $MailboxEmail | Where { $_.FolderPath -eq "/Top of Information Store" -or $_.FolderPath -eq "/Inbox" -or $_.FolderPath -eq "/Calendar" } } }
            Else {
                If ($SourceIsRest) { $Folders = (Get-ExoMailboxFolderStatistics -Identity $MailboxEmail | Select $_.FolderPath) }
                Else { $Folders = (Get-MailboxFolderStatistics $MailboxEmail | Select $_.FolderPath) } }
            ForEach ($Folder in $Folders) {
                $FolderPath = $Folder.FolderPath.Replace("/","\")
                If ($FolderPath -eq "\Top of Information Store") { $FolderPath = "\" }
                $FolderLocation = $MailboxEmail + ":" + $FolderPath
                If ($SourceIsRest) { $FolderPermissions = Get-ExoMailboxFolderPermission -Identity "$FolderLocation" -ErrorAction SilentlyContinue }
                Else { $FolderPermissions = Get-MailboxFolderPermission "$FolderLocation" -ErrorAction SilentlyContinue }
                If ($FolderPermissions -ne $null) {
                    ForEach ($Permission in $FolderPermissions) {
                        [string]$FolderDelegate = $Permission.User
                        If ($FolderDelegate -ne "Default" -and $FolderDelegate -ne "Anonymous") {
                            Start-Sleep -Milliseconds $MicroDelaySleep
                            $CheckDelegate = Get-Recipient $FolderDelegate -ErrorAction SilentlyContinue
                            If ($CheckDelegate.Count -gt 1) { Write-Log "ALERT: Delegate user matches multiple accounts. [$DelegateUser]" }
                            If ($CheckDelegate -ne $null) {
                                $DelegateEmail = $CheckDelegate.PrimarySmtpAddress
                                $DelegateType = $CheckDelegate.RecipientTypeDetails
                                $DelegateAccess = $Permission.AccessRights
                                "`"$MailboxEmail`",`"$FolderLocation`",`"$DelegateEmail`",`"$DelegateType`",`"$DelegateAccess`"" | Out-File $FolderDelegateExport -Encoding ascii -Append } } } } } } }

    # --- Export mailbox forwarding

    If ($IncludeMailboxForwarding -eq $true) {
        Write-Log "AUDIT: Mailbox forwarding..."
        $ForwardingAddress = ""; $CheckForwarding = Get-Mailbox $MailboxEmail
        If (($CheckForwarding.ForwardingAddress) -ne $null) { [string]$ForwardingAddress = (Get-Recipient $CheckForwarding.ForwardingAddress).PrimarySmtpAddress }
        If ($CheckForwarding.ForwardingSmtpAddress -ne $null) { [string]$ForwardingAddress = $CheckForwarding.ForwardingSmtpAddress }
        If ($ForwardingAddress -ne "") {
            $DeliverToMailbox = $CheckForwarding.DeliverToMailboxAndForward
            "`"$MailboxEmail`",`"$ForwardingAddress`",`"$DeliverToMailbox`"" | Out-File $MailboxForwardingExport -Encoding ascii -Append } }
}

Write-Log "SUCCESS: Script complete."