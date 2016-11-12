<#PSScriptInfo

.VERSION 1.2.0

.GUID 348975cd-3d3b-4eaf-bf6b-3fbe39c462a2

.AUTHOR Mike Hendrickson

.COMPANYNAME Microsoft Corporation

.COPYRIGHT (C) Microsoft Corporation. All rights reserved.

#>

<# 
.Synopsis
 Gets an inventory of the sizes of all mailboxes in the Exchange 2010/2013/2016 environment, and saves the results to .CSV.

.DESCRIPTION 
 Gets an inventory of the sizes of all mailboxes in the Exchange 2010/2013/2016 environment, and saves the results to .CSV.
 Works in all environments, but specifically designed for larger environments, where a one liner of 
 'Get-Mailbox | Get-MailboxStatistics | Export-Csv' may fail due to memory limitations. Gets the initial list
 of mailboxes directly from Active Directory to save both time and memory.

.EXAMPLE
PS> .\Get-MailboxSizeInventory.ps1 -GlobalCatalog GCNAME

.EXAMPLE
PS> .\Get-MailboxSizeInventory.ps1 -GlobalCatalog GCNAME -Verbose

#>

[CmdletBinding()]
param
(
    #Specifies the name of the Global Catalog to retrieve the list of mailboxes from
    [parameter(Mandatory = $true)]
    [string]
    $GlobalCatalog,

    #The name of the CSV file to output results to
    [string]
    $OutputFile = "MailboxSizes.csv",

    [string]
    $CustomLDAPFilter = ""
)

if ((Get-Module ActiveDirectory -ErrorAction SilentlyContinue) -eq $null)
{
    Import-Module ActiveDirectory -ErrorAction Stop
}

Write-Verbose "$([DateTime]::Now) Retrieving list of mailboxes from Active Directory"

if ([string]::IsNullOrEmpty($CustomLDAPFilter) -eq $false)
{
    $ldapFilter = $CustomLDAPFilter
}
else
{
    $ldapFilter = "(&(msExchHomeServerName=*)(mail=*)(objectClass=user)(!mail=healthmailbox*)(!mail=extest*)(!mail=systemmailbox*))"
}

$users = Get-ADObject -Server "$($GlobalCatalog):3268" -LDAPFilter $ldapFilter -Properties distinguishedName, mDBUseDefaults, mDBStorageQuota, mDBOverQuotaLimit, mDBOverHardQuotaLimit, msExchArchiveDatabaseLink, msExchArchiveWarnQuota, msExchArchiveQuota, msExchMailboxTemplateLink

Write-Verbose "$([DateTime]::Now) Retrieving statistics of $($users.Count) users"

[PSObject[]]$allStats = @()
[Hashtable]$retentionPolicyDNtoNameMap = @{}

for ($i = 0; $i -lt $users.Count; $i++)
{
    Write-Progress -Activity "Getting Mailbox Statistics" -Status "Processing mailbox $($i + 1) / $($users.Count)" -PercentComplete (($i + 1) * 100 / $users.Count)

    $user = $users[$i]

    $stats = $null
    $stats = Get-MailboxStatistics -Identity $user.distinguishedName

    if ($stats -ne $null)
    {
        if ($user.msExchMailboxTemplateLink -ne $null)
        {
            if ($retentionPolicyDNtoNameMap.ContainsKey($user.msExchMailboxTemplateLink) -eq $true)
            {
                $retentionPolicy = $retentionPolicyDNtoNameMap[$user.msExchMailboxTemplateLink]
            }
            else
            {
                $retentionPolicyValue = $null
                $retentionPolicyValue = Get-RetentionPolicy -Identity $user.msExchMailboxTemplateLink -ErrorAction SilentlyContinue

                if ($retentionPolicyValue -ne $null)
                {
                    $retentionPolicy = $retentionPolicyValue.Name
                    $retentionPolicyDNtoNameMap.Add($user.msExchMailboxTemplateLink, $retentionPolicy)
                }
            }
        }
        else
        {
            $retentionPolicy = ""
        }

        $allStats += ($stats | Select-Object DisplayName,StorageLimitStatus,@{Label=”TotalItemSize(Bytes)”;Expression={$_.TotalItemSize.Value.ToBytes()}},@{Label=”TotalDeletedItemSize(Bytes)”;Expression={$_.TotalDeletedItemSize.Value.ToBytes()}},@{Label="UseDatabaseQuotaDefaults";Expression={$user.mDBUseDefaults}},@{Label="IssueWarningQuota";Expression={$user.mDBStorageQuota}},@{Label="ProhibitSendQuota";Expression={$user.mDBOverQuotaLimit}},@{Label="ProhibitSendReceiveQuota";Expression={$user.mDBOverHardQuotaLimit}},@{Label="IsArchive";Expression={$false}},@{Label="ArchiveWarningQuota";Expression={""}},@{Label="ArchiveQuota";Expression={""}},@{Label="RetentionPolicy";Expression={$retentionPolicy}})
    }

    if ($user.msExchArchiveDatabaseLink -ne $null)
    {
        $archiveStats = $null
        $archiveStats = Get-MailboxStatistics -Identity $user.distinguishedName -Archive

        if ($archiveStats -ne $null)
        {
            $allStats += ($archiveStats | Select-Object DisplayName,StorageLimitStatus,@{Label=”TotalItemSize(Bytes)”;Expression={$_.TotalItemSize.Value.ToBytes()}},@{Label=”TotalDeletedItemSize(Bytes)”;Expression={$_.TotalDeletedItemSize.Value.ToBytes()}},@{Label="UseDatabaseQuotaDefaults";Expression={""}},@{Label="IssueWarningQuota";Expression={""}},@{Label="ProhibitSendQuota";Expression={""}},@{Label="ProhibitSendReceiveQuota";Expression={""}},@{Label="IsArchive";Expression={$true}},@{Label="ArchiveWarningQuota";Expression={$user.msExchArchiveWarnQuota}},@{Label="ArchiveQuota";Expression={$user.msExchArchiveQuota}},@{Label="RetentionPolicy";Expression={""}})
        }
    }
}

Write-Progress -Activity "Getting Mailbox Statistics" -Status "Complete" -PercentComplete 100
Write-Verbose "$([DateTime]::Now) Saving results to CSV"

$allStats | Export-Csv $OutputFile -NoTypeInformation