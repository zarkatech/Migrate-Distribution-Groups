$ScriptInfo = @"
================================================================================
Audit-MailGroupMembers.ps1 | v3.2.1
by Roman Zarka | github.com/zarkatech
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$BatchName = "DistrictABC"
$MailGroupsAuditFile = "c:\Scripts\12010714_MASTER_AuditMailGroups.csv"
$IncludeStaticGroupMembers = $true
$IncludeNestedGroupMembers = $true
$IncludeDynamicGroupMembers = $false
$IncludePublicFolderMembers = $false

# --- Initialize log files

$TimeStamp = Get-Date -Format MMddhhmm
If ($BatchName -ne "" -and $BatchName -ne $null) { $Timestamp = $Timestamp + "_$BatchName" }
$RunLog = $TimeStamp + "_AuditMailGroupMembers.log"
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
    If ($LogStatus -eq "AUDIT") {
        Write-Host $LogString -ForegroundColor DarkGray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "") {
        Write-Host ""
        Write-Output "`n" | Out-File $RunLog -Append }
}
# --- Initialize script environment

If ([int]$PsVersionTable.PSVersion.Major -lt 3) { Write-Log "ERROR: Script requires PowerShell v3 or later and must be run from an upgraded console."; Break }
If ((Get-PSSession) -eq $null -or ((Get-PSSession).ConfigurationName) -ne "Microsoft.Exchange") { Write-Log "ERROR: Script must be run from an Exchange session."; Break }
If (Test-Path $MailGroupsAuditFile) { Write-Log "INFO: Import mail groups audit data. [$MailGroupsAuditFile]"; $MailGroups = Import-Csv $MailGroupsAuditFile | Select Name, PrimarySmtpAddress }
Else { Write-Log "ERROR: Mail groups audit file not found. [$MailGroupsAuditFile]"; Break }

# --- Audit static group members

If ($MailGroups.Count -eq 0) { Write-Log "ALERT: No static mail groups collected." }
Else { Write-Log "SUCCESS: Found $($MailGroups.Count) groups." } 
If (($IncludeStaticGroupMembers) -and ($MailGroups.Count -ne 0)) {
    Write-Log "INFO: Audit static mail group members..."
    $MailGroupMembersExport = $Timestamp + "_AuditMailGroupMembers.csv"
    $Select = @(); $Select += '@{Name="ParentGroupName";Expression={$Group.Name}}'; $Select += '@{Name="ParentGroupEmail";Expression={$Group.PrimarySmtpAddress}}'
    ForEach ($Group in $MailGroups) {
        If ((Get-DistributionGroupMember $Group.PrimarySmtpAddress -ResultSize 1 -WarningAction SilentlyContinue) -ne $null) { $SampleGroup = $Group.PrimarySmtpAddress }
        Else { Write-Log "ALERT: Group has no members. [$($Group.PrimarySmtpAddress)]" } }
    Get-DistributionGroupMember $SampleGroup -ResultSize 1 -WarningAction SilentlyContinue | Get-Member | Where { $_.MemberType -eq "Property" } | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
    $Progress = @{Activity='Export static mail group members...';PercentComplete=0}
    Write-Progress @Progress; $Count=0
    ForEach ($Group in $MailGroups) {
        $Count++; [int]$Percentage = $Count/($MailGroups.Count)*100; $Progress.CurrentOperation = $Group.Name; $Progress.PercentComplete = $Percentage
        $RunCmd = 'Get-DistributionGroupMember $($Group.PrimarySmtpAddress) -ResultSize Unlimited | Select ' + ($Select -join ",")
        Invoke-Expression $RunCmd | Export-Csv $MailGroupMembersExport -Append -NoTypeInformation
        Write-Progress @Progress }
    Write-Progress @Progress -Completed }

# --- Audit nested groups

If ($IncludeNestedGroupMembers) {
    Write-Log "INFO: Audit nested groups..."
    $NestedGroupsExport = $Timestamp + "_AuditNestedGroups.csv"
    Function Check-NestedGroups ($ParentGroup, $ChildGroup) {
        $script:Level = $script:Level + 1; If ($script:Level -eq 1) { $TopGroup = $ParentGroup }; $Circular = $false
        #Write-Host "Child:$ChildGroup Parent:$ParentGroup Top:$TopGroup"
        If ($script:NestedGroups.ParentGroup -contains $ParentGroup) {
            Write-Log "ERROR: Circular nesting detected at level $($script:Level-1). [$ParentGroup]"; $Circular = $true }
        $script:NestedGroups +=@([pscustomobject]@{TopGroup="$TopGroup";ParentGroup="$ParentGroup";ChildGroup="$ChildGroup";Level="$($script:Level-1)";Circular="$Circular"})
        If ($Circular -eq $true) { Break }
        $CheckNesting = $script:ChildGroups | Where { $_.ParentGroupEmail -eq "$ChildGroup" } | Select ParentGroupEmail, PrimarySmtpAddress
        If ($CheckNesting -ne $null) { ForEach ($Member in $CheckNesting) { Check-NestedGroups $Member.ParentGroupEmail $Member.PrimarySmtpAddress } } }
    $script:ChildGroups = Import-Csv $MailGroupMembersExport | Where { $_.RecipientTypeDetails -like "*Group" } | Select ParentGroupEmail, PrimarySmtpAddress 
    ForEach ($Group in $script:ChildGroups) {
        If ($Group.PrimarySmtpAddress -ne "") {
            $script:NestedGroups = @(); $script:Level = 0
            Check-NestedGroups $Group.ParentGroupEmail $Group.PrimarySmtpAddress
            $script:NestedGroups | Export-Csv "$NestedGroupsExport" -NoTypeInformation -Append } }
}

# --- Script complete

Write-Log "SUCCESS: Script complete." -ForegroundColor Green
