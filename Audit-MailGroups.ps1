$ScriptInfo = @"
================================================================================
Audit-MailGroups.ps1 | v3.2.3
by Roman Zarka | Microsoft Services
================================================================================
SAMPLE SCRIPT IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
"@; cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Define preference variables

$UseFilterCriteria = $false
    $FilterBatchName = "DistrictABC"
    $FilterCriteria = '(CustomAttribute1 -eq "USA") -and (CustomAttribute3 -eq "IT") -and (CustomAttribute11 -ne "NoSyncO365")'
$IncludeStaticGroups = $false
    $IncludeGroupMembers = $false
    $IncludeNestedGroups = $false
$IncludeDynamicGroups = $true
$IncludePublicFolders = $false

# --- Initialize script environment

If ([int]$PsVersionTable.PSVersion.Major -lt 3) { Write-Log "ERROR: Script requires PowerShell v3 or later and must be run from an upgraded console."; Break }
If ((Get-PSSession) -eq $null -or ((Get-PSSession).ConfigurationName) -ne "Microsoft.Exchange") { Write-Log "ERROR: Script must be run from an Exchange session."; Break }
$TimeStamp = Get-Date -Format MMddhhmm
If ($UseFilterCriteria -eq $false) { $TimeStamp = $TimeStamp + "_MASTER" }
Else { $TimeStamp = $TimeStamp + "_$FilterBatchName" }

# --- Initialize log files

$RunLog = $TimeStamp + "_AuditMailGroups.log"
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

# --- Audit static mail groups

If ($IncludeStaticGroups) {
    Write-Log "INFO: Audit static mail groups..."
    $MailGroupsExport = $Timestamp + "_AuditMailGroups.csv"
    $Select = @(); Get-DistributionGroup -ResultSize 1 -WarningAction SilentlyContinue | Get-Member -MemberType Property | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
    $RunCmd = 'Get-DistributionGroup -ResultSize Unlimited'
    If (($UseFilterCriteria) -and ($FilterCriteria -ne "")) { $RunCmd = $RunCmd + ' -Filter {' + "$FilterCriteria" + '}' }
    $RunCmd = $RunCmd + ' | Select ' + ($Select -join ",")
    Invoke-Expression $RunCmd | Export-Csv $MailGroupsExport -NoTypeInformation }

# --- Audit static group members

If ($IncludeMailGroupMembers) {
    $MailGroups = Import-Csv $MailGroupsExport | Select Name, PrimarySmtpAddress
    If ($MailGroups.Count -eq 0) { Write-Log "ALERT: No static mail groups collected." }
    Else { Write-Log "SUCCESS: Found $($MailGroups.Count) groups." } 
    If (($IncludeGroupMembers) -and ($MailGroups.Count -ne 0)) {
        Write-Log "INFO: Audit static mail group members..."
        $MailGroupMembersExport = $Timestamp + "_AuditMailGroupMembers.csv"
        $Select = @(); $Select += '@{Name="ParentGroupName";Expression={$Group.Name}}'; $Select += '@{Name="ParentGroupEmail";Expression={$Group.PrimarySmtpAddress}}'
        ForEach ($Group in $MailGroups) {
            If ((Get-DistributionGroupMember $Group.PrimarySmtpAddress -ResultSize 1 -WarningAction SilentlyContinue) -ne $null) { $SampleGroup = $Group.PrimarySmtpAddress }
            Else { Write-Log "ALERT: Group has no members. [$($Group.PrimarySmtpAddress)]" } }
        Get-DistributionGroupMember $SampleGroup -ResultSize 1 -WarningAction SilentlyContinue | Get-Member -MemberType Property | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
        $Progress = @{Activity='Export static mail group members...';PercentComplete=0}
        Write-Progress @Progress; $Count=0
        ForEach ($Group in $MailGroups) {
            $Count++; [int]$Percentage = $Count/($MailGroups.Count)*100; $Progress.CurrentOperation = $Group.Name; $Progress.PercentComplete = $Percentage
            $RunCmd = 'Get-DistributionGroupMember $($Group.PrimarySmtpAddress) -ResultSize Unlimited | Select ' + ($Select -join ",")
            Invoke-Expression $RunCmd | Export-Csv $MailGroupMembersExport -Append -NoTypeInformation
            Write-Progress @Progress }
        Write-Progress @Progress -Completed }
}

# --- Audit dynamic groups

If ($IncludeDynamicGroups) {
    Write-Log "INFO: Audit dynamic mail groups..."
    $DynamicGroupsExport = $Timestamp + "_AuditDynamicGroups.csv"
    $Select = @(); Get-DynamicDistributionGroup -ResultSize 1 -WarningAction SilentlyContinue | Get-Member -MemberType Property | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
    $RunCmd = 'Get-DynamicDistributionGroup -ResultSize Unlimited'
    If (($UseFilterCriteria) -and ($FilterCriteria -ne "")) { $RunCmd = $RunCmd + ' -Filter {'+"$FilterCriteria"+'}' }
    $RunCmd = $RunCmd + ' | Select ' + ($Select -join ",")
    Invoke-Expression $RunCmd | Export-Csv $DynamicGroupsExport  -NoTypeInformation
    Import-Csv $DynamicGroupsExport | ForEach {
        $GroupPreview = Get-DynamicDistributionGroup $_.Guid
        $DynamicMembers = (Get-Recipient -RecipientPreviewFilter $GroupPreview.RecipientFilter -OrganizationalUnit $GroupPreview.RecipientContainer)
        If ($DynamicMembers -eq $null) { Write-Log "ALERT: Dynamic group preview found no qualifying members. [$($GroupPreview.PrimarySmtpAddress)]" } }
        
}

# --- Audit mail public folders

If ($IncludePublicFolders) {
    Write-Log "INFO: Audit mail public folders..."
    $PublicFolderExport = $Timestamp + "_AuditPublicFolders.csv"
    $Select = @(); Get-MailPublicFolder -ResultSize 1 -WarningAction SilentlyContinue | Get-Member -MemberType Property | ForEach { $Select += '@{Name="'+"$($_.Name)"+'";Expression={$_.'+"$($_.Name)"+' -join ";"}}' }
    $RunCmd = 'Get-MailPublicFolder -ResultSize Unlimited'
    If (($UseFilterCriteria) -and ($FilterCriteria -ne "")) { $RunCmd = $RunCmd + ' -Filter {'+"$FilterCriteria"+'}' }
    $RunCmd = $RunCmd + ' | Select ' + ($Select -join ",")
    Invoke-Expression $RunCmd | Export-Csv $PublicFolderExport -NoTypeInformation }

# --- Audit nested groups

If (($IncludeGroupMembers) -and ($IncludeNestedGroups)) {
    Write-Log "INFO: Audit nested groups..."
    $NestedGroupsExport = $Timestamp + "_AuditNestedGroups.csv"
    Function Check-NestedGroups ($ParentGroup, $ChildGroup) {
        $script:Level = $script:Level + 1; If ($script:Level -eq 1) { $TopGroup = $ParentGroup }; $Circular = $false
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