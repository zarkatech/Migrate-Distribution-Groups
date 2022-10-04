# Migrate-Distribution-Groups
## PowerShell scripts to audit source distribution groups, deprovision and recreate them in target environment, re-apply associated mailbox permsissions, and back-fill a contact.

INTRODUCTION

Migrating distribution groups from on-premises Exchange or between Exchange Online tenants can be categorized into four distinct types: static distribution groups, mail-enabled security groups, dynamic distribution groups, and static permission groups.  Each group type requires a slightly different approach, but it is strongly recommended that all groups be migrated at the same time to mitigate issues related to group nesting, delegate permissions, mail flow, and overall user experience. If groups must be migrated in batches, then script logic can help facilitate analysis and processing of which groups can be migrated and when they should be migrated.  It is also assumed that user mailbox migrations have already occured.

STEP 1: AUDIT TARGET MAILBOX PERMISSIONS

Run “Audit-MailboxPermissions.ps1” script in target environment to export mailbox access and delegate permissions so they can be re-applied after group migrations. Script offers the following preference variables which can be toggled or customized depending on scenario requirements:
* $UseImportFile = $true / $false
*	$ImportFile = path and filename of CSV user list which includes “PrimarySmtpAddress”
*	$UseFilterCriteria = $true / $false
*	$FilterBatchName = short string to identify batch process, log, or export content.
*	$FilterCriteria = literal filter syntax used to scope audit results
*	$IncludeMailboxAccess = $true / $false 
*	$IncludeSendAs = $true / $false
*	$IncludeSendOnBehalf = $true / $false
*	$IncludeFolderDelegates = $true / $false
*	$IncludeMailboxForwarding = $true / $false
*	$DelegatesToSkip = service or built-In accounts which should not be audited

Script can be run once and resulting datasets used throughout the entire group migration process or it can be run prior to each migration batch to re-assess any recent changes. Exported CSV files should be preserved as a "point-in-time" snapshot and can be used for auditing or any roll-back contingencies.

STEP 2: AUDIT SOURCE MAIL GROUPS

Run “Audit-MailGroups.ps1” script is source environment to audit the existing mail group information, memberships, and all associated metadata. Script offers the following preference variables which can be customized depending on scenario requirements:
* $UseFilterCriteria = $true / $false
* $FilterBatchName = short string to identify batch process, log, or export content.
* $FilterCriteria = literal filter syntax used to scope audit results.
* $IncludeStaticGroups = $true / $false
* $IncludeDynamicGroups = $true / $false
* $IncludePublicFolders = $true / $false
* $IncludeGroupMembers = $true / $false
* $IncludeNestedGroups = $true / $false

Datasets will be exported into separate CSV files and referenced throughout the mail group migration process. Exported CSV files should be preserved as a "point-in-time" snapshot and can be used for auditing or any roll-back contingencies. Setting preference variable “$IncludeGroupMembers = $false” will allow for manual analysis and/or batch filtering of the mail groups dataset before auditing the mail group members. Defined filter criteria must be enclosed in single quotes as literal syntax, but not all attributes support the filter parameter. Dynamic distribution groups do not synchronize to O365 and must be manually re-provisioned using available attribute criteria in EXO. In the meantime, dynamic distribution groups can remain on-premises so that they can continue to leverage OU as filter or recipient criteria. On-premises permission groups must be mail-enabled and synchronized to EXO to preserve mailbox or delegate permissions and can then be migrated like any other mail group.

STEP 3: AUDIT SOURCE MAIL GROUP MEMBERS

If mail group members were audited and exported along with mail groups in STEP 2 then this step can be skipped.  Otherwise, run “Audit-MailGroupMembers.ps1” script against the manually refined or filtered mail group dataset to export the associated group memberships. Script offers the following preference variables which can be toggled or customized depending on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content.
* $MailGroupsAuditFile = path and filename of refined mail group audit data
* $IncludeStaticGroupMembers = $true / $false
* $IncludeNestedGroupMembers = $true / $false
* $IncludeDynamicGroupMembers = $true / $false
* $IncludePublicFolderMembers = $true / $false

STEP 4: PRE-STAGE SHADOW GROUPS

Pre-staging shadow groups in the target environment is optional and not required for mail group migrations.  However, staging shadow groups may offer benefits under certain circumstances. Shadow groups use temporary name, alias, and email address attributes to stage hidden mail groups alongside the production groups which can be cut-over quickly by simply renaming those attributes.  However, any changes to production groups while shadow groups are staged will be lost during the cut-over.

Run “Create-TargetShadowGroups.ps1” script if pre-staging shadow groups is strategically desirable or technically warranted. Script offers the following preference variables which can be toggled or customized depending on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content.
*	$MailGroupsAuditFile = path and filename of mail group audit data
*	$MailGroupMembersAuditFile = path and filename of mail group members audit data
*	$SetRbacScope = $true / $false
*	$RbacAttribute = attribute name defined for delegated group management
*	$RbacValue = attribute value defined for delegated group management
*	$NewGroupNamePrefix = string value to prepend when creating new target group
*	$NewGroupNameSuffix = string value to append when creating new target group
*	$CreateShadowGroups = $true / $false
*	$ShadowGroupPrefix = string value used to prepend temporary shadow group attributes
*	$CutoverShadowGroups = $true / $false
*	$PurgeShadowGroups = $true / $false
*	$Debug = $true / $false

Successful cut-over of staged shadow groups can only occur after production mail groups have been deleted or deprovisioned from target environment using STEP 5. Staging shadow groups and their members can also help estimate the perceived cut-over outage windo, but should be purged before re-creating target groups using the latest or most recent mail group audit data.

STEP 5: DEPROVISION SOURCE MAIL GROUPS

In a hybrid environment, static mail groups must be de-provisioned or deleted from on-premises Exchange and recreated in Exchange Online using mail-enabled members only. Run “Deprovision-MailGroups.ps1” script in source environment to automate the mail group deprovisioning tasks. Script offers the following preference variables which can be toggled or customized based on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content.
* $UseNoSync = $true / $false
* $NoSyncAttribute = attribute name used by AAD Connect for account filtering
* $NoSyncValue = attribute value used by AAD Connect for account filtering
* $DeleteDistributionGroups = $true / $false
* $DisableMailSecurityGroups = $true / $false
* $Debug = $true / $false

The No-Sync attribute and value must be defined in AAD Connect to automatically deprovision synchronized mail groups without deleting them from on-premises. Script can delete distribution groups (if toggled) so that on-premises contacts can be back-filled for GAL synchronization or LDAP lookup. Script can also disable mail-enabled security groups (if toggled) while preserving the group and its members as an AD security group for other Windows-integrated access control lists. Script offers the following preference variables which can be toggled or customized based on scenario requirements:

STEP 6: CREATE TARGET MAIL GROUPS

If shadow groups have been staged in STEP 4 with the intention of production cut-over, then re-run “Create-TargetShadowGroups.ps1” with preference variables $CutoverShadowGroups set to $true. Otherwise, ensure shadow groups have been purged by running "Create-TargetShadowGroups.ps1" with $PurgeShadowGroups set to $true and then run “Create-TargetGroups.ps1” to re-create mail groups, re-assign members, and re-configure metadata options. Script offers the following preference variables which can be toggled or customized based on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content
* $SetRbacScope = $true / $false
* $RbacAttribute = attribute name defined for delegated group management
* $RbacValue = attribute value defined for delegated group management
* $GroupNamePrefix = string value to prepend when creating new target group
* $GroupNameSuffix = string value to append when creating new target group
* $Debug = $true / $false

First pass will delete any contacts representing source mail groups and creates new mail groups in target environment using the priority parameters collected in STEP 2. Second pass will re-assign group members including mailboxes, nested groups, dynamic groups, and contacts based on information collected in STEP 3. Third pass attempts to re-apply extended group settings based on information collected in STEP 2 and exports relevant data for back-filling contacts in STEP 8.

STEP 7: RE-APPLY TARGET MAILBOX PERMISSIONS

Run “Apply-MailboxPermissions.ps1” in target environment to re-apply mailbox access and delegate permission for new groups based on data collected in STEP 1. Script offers the following preference variables which can be toggled or customized based on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content
* $IncludeMailboxAccess = $true / $false
* $IncludeSendAs = $true / $false
* $IncludeSendOnBehalf = $true / $false
* $IncludeFolderDelegates = $true / $false
* $IncludeMailboxForwarding = $true / $false
* $ApplyGroupPermissionsOnly = $true / $false
* $Debug = $true / $false

STEP 8: BACK-FILL MAIL GROUP CONTACTS

If or when source mail groups must be deleted, run "Create-SourceContants.ps1" to backfill mail contacts based on audit data collected during target group creation or cut-over. Script offers the following preference variables which can be toggled or customized based on scenario requirements:
* $BatchName = short string to identify batch process, log, or export content
* $TargetOU = organizational unit to create mail contacts
* $NoSyncAttribute = attribute name used by AAD Connect for account filtering
* $NoSyncValue = attribute value used by AAD Connect for account filtering 
*	$DeleteDistributionGroups = $true / $false
*	$DisableMailSecurityGroups = $true / $false
*	$Debug = $true / $false

Script will delete conflicting distribution groups and disable conflicting mail-enabled security groups so that contact can be created using the primary email address. Contacts should be created in an OU that is excluded from AAD Connect or set as "NoSync" using the preference variables.

