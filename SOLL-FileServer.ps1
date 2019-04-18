<#
VERSION   DATE          AUTHOR
1.0       2019-01-16    Nicola Marini <n.marini@novigo-consulting.it>
    - Initial Release
1.1       2019-03-18    Nicola Marini <n.marini@novigo-consulting.it>
    - Fix error in check if network share ending with a $ exists in Active Directory as user

#> # Revision history

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="FQDN of server for which you need to produce SOLL")]
    [String]
    $ServerName = "",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Folder in which to save CSV files")]
    [String]
    $OutputFolder = "SOLL-OUTPUT",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV containing all AD groups that should be flagged as admin")]
    [String]
    $AdminGroupsFile = "admin-groups.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing associations between shares and its business owner")]
    [String]
    $ShareBusinessOwnerFile = "bo.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV containing all AD groups that should be flagged as admin")]
    [String]
    $FoldersToSkip = "skip-folders.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing ULM export with username and user details")]
    [String]
    $ULMFile = "ULM.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing non personal accounts")]
    [String]
    $NPAFile = "NPA.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Enable or disable Active Directory groups resolution")]
    [bool]
    $ResolveADGroups = $false,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="If set to true it uses ShareBusinessOwnerFile as an exclusion list of folders which should not be elaborated")]
    [bool]
    $ExcludeBOFolders = $false
) #Accept input parameters

$CSVSeparator = ";"

<#
$ROPermissions = @(
    "ListDirectory",
    "ReadData",
    "ReadExtendedAttributes",
    "Traverse",
    "ExecuteFile",
    "ReadAttributes",
    "ReadPermissions",
    "Read",
    "ReadAndExecute",
    "Synchronize"
)
#>

$RWPermissions = @(
    "WriteData",
    "CreateFiles",
    "CreateDirectories",
    "AppendData",
    "WriteExtendedAttributes",
    "DeleteSubdirectoriesAndFiles",
    "WriteAttributes",
    "Write",
    "Delete",
    "Modify",
    "ChangePermissions",
    "TakeOwnership",
    "FullControl"
)

$NoResolveADGroups = @(
    "Domain Users",
    "Authenticated Users",
    "Users",
    "Everyone"
)

# Declare main variables
[System.Collections.ArrayList]$BusinessOwners = @();
[System.Collections.ArrayList]$Acl = @();
[System.Collections.ArrayList]$ULM = @();
[System.Collections.ArrayList]$PersonalAccounts = @();
$BOPath = @();
$OldOwner = "";

# Main function
Function Main {
    Test-InputParameters
    
    [System.Collections.ArrayList]$BusinessOwners = @(Import-Csv $ShareBusinessOwnerFile -Delimiter $CSVSeparator | Select-Object -Unique Share, Owner | Sort-Object -Property Owner)

    $AdminGroups = @()
    Import-Csv $AdminGroupsFile -Delimiter $CSVSeparator | ForEach-Object { $AdminGroups += $_.AdminGroup }

    $NonPersonalAccounts = @()
    Import-Csv $NPAFile -Delimiter $CSVSeparator | ForEach-Object { $NonPersonalAccounts += $_.UserID }

    [System.Collections.ArrayList]$ULM = Import-Csv $ULMFile -Delimiter $CSVSeparator | Select-Object -Unique LoginId, LeaseUser, MasterUser, RootUser, SupportUser, LastName, FirstName, Area, Department, JobTitle, EmailAddress, Model
    
    $fToSkip = @()
    if ($FoldersToSkip -ne "" -and (Test-Path -Path $FoldersToSkip)) {
        Import-Csv $FoldersToSkip -Delimiter $CSVSeparator | ForEach-Object { $fToSkip += $_.Folder }
    }

    If ($ExcludeBOFolders) {
        $BusinessOwners.ForEach({
            $Owner = $_.Owner
            if ($Owner -ne "") {
                $BOPath += $_.Share
            }
        })
    }
    Else {
        $BusinessOwners.ForEach({
            $Owner = $_.Owner
            if ($Owner -ne "") {
                $BOPath += $_.Share
                if ($OldOwner -ne "") {
                    if ($null -eq $BODetails.EmailAddress) {
                        Write-Host $Owner - $OldOwner - $CurAcl.ShareFullPath
                        Write-Progress "Elaborating share $($CurAcl.Share)..." -Status "Save to file NO-OWNER_$($OldOwner)_ShareFunctionalMatrix.csv..."
                        Export-SollCsv -Content $Acl -FileName "NO-OWNER_$($OldOwner)_ShareFunctionalMatrix.csv" -Append $true


                    }
                    else {
                        Write-Progress "Elaborating share $($CurAcl.Share)..." -Status  "Save to file $($OldOwner)_ShareFunctionalMatrix.csv..."
                        Export-SollCsv -Content $Acl -FileName "$($OldOwner)_ShareFunctionalMatrix.csv" -Append $true
                    }

                    # Reset Acl
                    $Acl.Clear()
                    
                    [gc]::collect()
                    [gc]::WaitForPendingFinalizers()
                }

                if ($Owner -ne $OldOwner){
                    
                    $BODetails= $ULM.Where({$_.LoginId -eq $Owner});
                    
                }

                $CurShare = $_.Share
                try {
                    Write-Progress "Elaborating share $($_.Share)..."
                    (Get-ShareAcl -Path (Get-Item $_.Share -ErrorAction Stop) -BusinessOwner $BODetails -Recursive $true -IncludeInherited $true).ForEach({
                        $CurAcl = $_;
                        $ExistingAcl = $Acl.Where({$_.ShareFullPath -eq $CurAcl.ShareFullPath -and $_.ShareAccessGroup -eq $CurAcl.ShareAccessGroup}) | Select-Object -First 1
                        if ($null -ne $ExistingAcl.ShareFullPath) {
                            Merge-AclPermissions -ExistingAcl $ExistingAcl -NewAcl $CurAcl
                        }
                        else {
                            $Acl.Add($_)
                        }
                    });
                }
                catch {
                    $Acl.Add([PSCustomObject]@{
                        "Entity" = $CurShare.TrimStart("\\").Substring($CurShare.TrimStart("\\").IndexOf("\")+1).Replace("\", ":");
                        "BusinessOwnerName" = $BODetails.LastName + " " + $BODetails.FirstName;
                        "BusinessOwnerEmail" = $BODetails.EmailAddress;
                        "ShareFullPath" = $CurShare;
                        "ShareAccessGroup" = "Err";
                        "ShareUserRW" = "Err";
                        "ShareUserRO" = "Err";
                        "ShareUserDeny" = "Err";
                        "ShareAppRW" = "Err";
                        "ShareAppRO" = "Err";
                        "ShareAppDeny" = "Err";
                        "ShareAdmin" = "Err";
                    })
                }
                $OldOwner = $Owner
            }

        });

        # Export last item not exported in the loop
        if ($null -eq $BODetails.EmailAddress) {
            Write-Host "Last item: $OldOwner" -ForegroundColor Green
            Write-host "Elaborating share $($CurAcl.Share)..." -Status -Verbose "Save to file NO-OWNER_$($OldOwner)_ShareFunctionalMatrix.csv..."
            Export-SollCsv -Content $Acl -FileName "NO-OWNER_$($OldOwner)_ShareFunctionalMatrix.csv" -Append $true
        }
        else {
            Write-host "Elaborating share $($CurAcl.Share)..." -Status -Verbose "Save to file $($OldOwner)_ShareFunctionalMatrix.csv..."
            Export-SollCsv -Content $Acl -FileName "$($OldOwner)_ShareFunctionalMatrix.csv" -Append $true
        }
    }

    If ($ServerName -ne "1") {

        # TODO
        # Get ACL for all shared folders not contained in input file
        #$NetworkShares = net view \\prdsshtp /all | Select-Object -Skip 7 | Where-Object {$_ -match 'disk*'} | ForEach-Object {$_ -match '^(.+?)\s+Disk*'|out-null;$matches[1]};
        $NetworkShares = ls -Directory -LiteralPath G:\Data | Select-Object Name -ExpandProperty name
        $Acl = @()
    
        $Root = [ADSI]''
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher($Root)

        ForEach ($NetworkShare in $NetworkShares) {
            $toElaborate = $true
            # Trim ending $ and check if share name is a user
            If ($NetworkShare.EndsWith("$")) {
                $Searcher.filter = "(&(objectClass=user)(sAMAccountName= $($NetworkShare.TrimEnd("$"))))"

                $res = $Searcher.FindOne()

                If ($res.Count -gt 0 -or $NetworkShare.length -le 2) {
                    $toElaborate = $false
                }
            }

            # If share is not a user then elaborate

            If ($toElaborate) {
                $FullSharePath = "G:\Data\$NetworkShare";
                If ($BOPath -notcontains $FullSharePath) {
                    try {
                        Write-host "Elaborating share $($FullSharePath)..."
                        (Get-ShareAcl -Path (Get-Item $FullSharePath -ErrorAction Stop) -BusinessOwner $null -Recursive $true -IncludeInherited $true).ForEach({
                            $CurAcl = $_;
                            $ExistingAcl = $Acl.Where({$_.ShareFullPath -eq $CurAcl.ShareFullPath -and $_.ShareAccessGroup -eq $CurAcl.ShareAccessGroup}) | Select-Object -First 1
                            if ($null -ne $ExistingAcl.ShareFullPath) {
                                Merge-AclPermissions -ExistingAcl $ExistingAcl -NewAcl $CurAcl
                            }
                            else {
                                $Acl.Add($_)
                            }
                        });
                    }
                    catch {
                        $Acl.Add([PSCustomObject]@{
                            "Entity" = $FullSharePath;
                            "BusinessOwnerName" = $null;
                            "BusinessOwnerEmail" = $null;
                            "ShareFullPath" = $FullSharePath;
                            "ShareAccessGroup" = "Err";
                            "ShareUserRW" = "Err";
                            "ShareUserRO" = "Err";
                            "ShareUserDeny" = "Err";
                            "ShareAppRW" = "Err";
                            "ShareAppRO" = "Err";
                            "ShareAppDeny" = "Err";
                            "ShareAdmin" = "Err";
                        })
                    }

                
                    #Export-SollCsv -Content $Acl -FileName "$($NetworkShare.Replace("\","-"))-NO-OWNER_ShareFunctionalMatrix.csv"
                    Write-Progress "Elaborating share $($FullSharePath)..." -Status "Save to file NO-OWNER_ShareFunctionalMatrix.csv..."
                    Export-SollCsv -Content $Acl -FileName "NO-OWNER_ShareFunctionalMatrix.csv" -Append $true
                    $Acl.Clear()
                    
                    Clear-GC
                }
            }
        }
        
        Write-Progress "Elaborating share $($FullSharePath)..." -Status "Save to file NO-OWNER_ShareFunctionalMatrix.csv..."
        Export-SollCsv -Content $Acl -FileName "NO-OWNER_ShareFunctionalMatrix.csv" -Append $true
    }

    if ($ResolveADGroups -eq $true) {
        Export-PersonalAccountCsv -Content $PersonalAccounts -FileName "GLOBAL_SubjectMatrixPA.csv"
    }
}

Function Get-ShareAcl ($Path, $BusinessOwner, $Recursive=$false, $IncludeInherited=$false) {
    if ($fToSkip -notcontains $Path.FullName) {
        $lAcl = $null
        Write-Progress "Elaborating share $($Path.FullName)..."
        try {
            If ($IncludeInherited -eq $true) {
                $lAcl = Get-Acl -Path $Path
            }
            Else {
                $lAcl = Get-Acl -Path $Path | Where-Object {$_.Access.IsInherited -eq $false}

                # If the folder has explicit permissions (not inherited) include also inherited permissions
                If ($null -ne $lAcl) {
                    $lAcl = Get-Acl -Path $Path
                }
            }
            $lAcl.ForEach({
                $_.Access.ForEach({
                    $cIsReadOnly = $true;
                    $cIsNPA = $false;
                    $cIsAdmin = $false;
                    $cIsDeny = $false;

                    $LocalPA = Add-PersonalAccount -Account $_.IdentityReference

                    If ($AdminGroups -contains $_.IdentityReference) {
                        $cIsAdmin = $true;
                    }
            
                    #If ($NonPersonalAccounts -contains $_.IdentityReference -or $NonPersonalAccounts -contains $_.IdentityReference.ToString().SubString($_.IdentityReference.ToString().IndexOf("\")+1)) {
                    If ($LocalPA.IsNPA -or $LocalPa.ContainsNPA) {
                        $cIsNPA = $true;
                    }
                    If ($_.AccessControlType -eq "Deny" -and $_.FileSystemRights.ToString() -ilike "*FullControl*") {
                        $cIsDeny = $true;
                    }
                    If ($_.AccessControlType -eq "Allow") {
                        ForEach ($right in ($_.FileSystemRights.ToString().Split(",").Trim())) {
                            If ($RWPermissions -contains $right) {
                                $cIsReadOnly = $false;
                                break;
                            }
                        }
                    }
            
                    [PSCustomObject]@{
                        "Entity" = $Path.FullName.TrimStart("\\").Substring($Path.FullName.TrimStart("\\").IndexOf("\")+1).Replace("\", ":");
                        "BusinessOwnerName" = $BusinessOwner.LastName + " " + $BusinessOwner.FirstName;
                        "BusinessOwnerEmail" = $BusinessOwner.EmailAddress;
                        "ShareFullPath" = $Path.FullName;
                        "ShareAccessGroup" = $_.IdentityReference;
                        "ShareUserRW" = If (-not $cIsAdmin -and -not $cIsNPA -and $_.AccessControlType -eq "Allow" -and -not $cIsReadOnly) { "X" } Else { $null };
                        "ShareUserRO" = If (-not $cIsAdmin -and -not $cIsNPA -and $_.AccessControlType -eq "Allow" -and $cIsReadOnly) { "X" } Else { $null };
                        "ShareUserDeny" = If (-not $cIsAdmin -and -not $cIsNPA -and $cIsDeny) { "X" } Else { $null };
                        "ShareAppRW" = If ($cIsNPA -and $_.AccessControlType -eq "Allow" -and -not $cIsReadOnly) { "X" } Else { $null };
                        "ShareAppRO" = If ($cIsNPA -and $_.AccessControlType -eq "Allow" -and $cIsReadOnly) { "X" } Else { $null };
                        "ShareAppDeny" = If ($cIsNPA -and $cIsDeny) { "X" } Else { $null };
                        "ShareAdmin" = If ($cIsAdmin) { "X" } Else { $null };
                    }
                });

                if ($Recursive -eq $true) {

                    (Get-ChildItem -Path $Path -Directory -ErrorAction Stop).ForEach({
                        If ($BOPath -notcontains $_.FullName) {
                            Get-ShareAcl -Path (Get-Item $_.FullName -ErrorAction Stop) -BusinessOwner $BusinessOwner -Recursive $true
                        }
                    });
                }
        
            });
        }
        catch {
            [PSCustomObject]@{
                "Entity" = $Path.FullName.TrimStart("\\").Substring($Path.FullName.TrimStart("\\").IndexOf("\")+1).Replace("\", ":");
                "BusinessOwnerName" = $BusinessOwner.LastName + " " + $BusinessOwner.FirstName;
                "BusinessOwnerEmail" = $BusinessOwner.EmailAddress;
                "ShareFullPath" = $Path.FullName;
                "ShareAccessGroup" = "Err";
                "ShareUserRW" = "Err";
                "ShareUserRO" = "Err";
                "ShareUserDeny" = "Err";
                "ShareAppRW" = "Err";
                "ShareAppRO" = "Err";
                "ShareAppDeny" = "Err";
                "ShareAdmin" = "Err";
            }
        }
    }
    else {
        [PSCustomObject]@{
            "Entity" = $Path.FullName.TrimStart("\\").Substring($Path.FullName.TrimStart("\\").IndexOf("\")+1).Replace("\", ":");
            "BusinessOwnerName" = $BusinessOwner.LastName + " " + $BusinessOwner.FirstName;
            "BusinessOwnerEmail" = $BusinessOwner.EmailAddress;
            "ShareFullPath" = $Path.FullName;
            "ShareAccessGroup" = "Err";
            "ShareUserRW" = "Err";
            "ShareUserRO" = "Err";
            "ShareUserDeny" = "Err";
            "ShareAppRW" = "Err";
            "ShareAppRO" = "Err";
            "ShareAppDeny" = "Err";
            "ShareAdmin" = "Err";
        }
    }
}

Function Add-PersonalAccount($Account, $ParentGroup = $null) {
    If (-not $Account.ToString().StartsWith("S-1-5-21")) {
        $PA = $PersonalAccounts.Where({ $_.AccountName -eq $Account })
        If (-not $PA) {
            $accWithoutDomain = $Account.ToString().Substring($Account.ToString().IndexOf("\")+1)
            
            if ($NoResolveADGroups -contains $accWithoutDomain) {
                $TempPA = [PSCustomObject]@{
                    "AccountName" = $Account;
                    "Type" = $group;
                    "IsNPA" = $false;
                    "ContainsNPA" = $false;
                    "ParentGroup" = $ParentGroup;
                    "Members" = $null;
                }
            }
            else {
                $lAccSid = $Account.Translate([System.Security.Principal.SecurityIdentifier])
                $lAccount = [ADSI]"LDAP://<SID=$lAccSid>"

                $lType = $lAccount.SchemaClassName

                $lMembers = $null
                
                $lContainsNpa = $false

                If ( $lType -eq "group" -and $ResolveADGroups -eq $true ) {
                    $lMembers = @()

                    ForEach ($lMember in $lAccount.Properties["member"]) {
                        $lGroupMember = [ADSI]"LDAP://$lMember"
                        $lMembers += $lGroupMember
                        
                        If ($NonPersonalAccounts -contains $lGroupMember.Properties["sAMAccountName"]) {
                            $lContainsNpa = $true
                        }

                        If (@('group', 'user') -contains $lGroupMember.SchemaClassName) {
                            Add-PersonalAccount -Account (New-Object System.Security.Principal.SecurityIdentifier $lGroupMember.objectSid[0], 0).Translate([System.Security.Principal.NTAccount]) -ParentGroup $Account.ToString() | Out-Null
                        }
                    }
                }
                
                $TempPA = [PSCustomObject]@{
                    "AccountName" = $Account;
                    "Type" = $lType;
                    "IsNPA" = If ( $lType -eq "user" -and ($NonPersonalAccounts -contains $Account -or $NonPersonalAccounts -contains $accWithoutDomain) ) { $true } Else { $false };
                    "ContainsNPA" = $lContainsNpa;
                    "ParentGroup" = $ParentGroup;
                    "Members" = $lMembers.Count;
                }
            }

            $PersonalAccounts.Add($TempPA)
            $TempPA
        }
        Else {
            $tempNewPA = $null
            $grPA = $PA.Where({$_.ParentGroup -eq $ParentGroup}) | Select-Object -First 1
            If (-not $grPA) {
                $obj = $PA | Select-Object -First 1
                $tempNewPA = [PSCustomObject]@{
                    "AccountName" = $obj.AccountName;
                    "Type" = $obj.Type;
                    "IsNPA" = $obj.IsNPA;
                    "ContainsNPA" = $obj.ContainsNpa;
                    "ParentGroup" = $ParentGroup;
                    "Members" = $obj.Members;
                }

                $PersonalAccounts.Add($tempNewPA);
                $tempNewPA
            }
            Else {
                $grPA
            }
        }
    }
}

Function Merge-AclPermissions($ExistingAcl, $NewAcl) {
    if ($null -ne $NewAcl.ShareUserDeny) {
        $ExistingAcl.ShareUserDeny = $NewAcl.ShareUserDeny;
        $ExistingAcl.ShareUserRW = $null;
        $ExistingAcl.ShareUserRO = $null;
    }
    elseif ($null -ne $NewAcl.ShareUserRW) {
        $ExistingAcl.ShareUserRW = $NewAcl.ShareUserRW
        $ExistingAcl.ShareUserRO = $null;
    }
    elseif ($null -ne $NewAcl.ShareUserRO) {
        $ExistingAcl.ShareUserRO = $NewAcl.ShareUserRO
    }

    if ($null -ne $NewAcl.ShareAppDeny) {
        $ExistingAcl.ShareAppDeny = $NewAcl.ShareAppDeny;
        $ExistingAcl.ShareAppRW = $null;
        $ExistingAcl.ShareAppRO = $null;
    }
    elseif ($null -ne $NewAcl.ShareAppRW) {
        $ExistingAcl.ShareAppRW = $NewAcl.ShareAppRW
        $ExistingAcl.ShareAppRO = $null;
    }
    elseif ($null -ne $NewAcl.ShareAppRO) {
        $ExistingAcl.ShareAppRO = $NewAcl.ShareAppRO
    }
}

Function Export-SollCsv ($Content, $FileName, $Append=$false) {
    $Content | 
        Select-Object @{expression={$_.Entity}; label="Entità" },
            @{expression={$_.BusinessOwnerName}; label="Business Owner" },
            @{expression={$_.BusinessOwnerEmail}; label="E-mail Business Owner" },
            @{expression={$_.ShareFullPath}; label="Full path" },
            @{expression={$_.ShareAccessGroup}; label="Gruppo di accesso" },
            @{expression={$_.ShareUserRW}; label="User RW" },
            @{expression={$_.ShareUserRO}; label="User RO" },
            @{expression={$_.ShareUserDeny}; label="User Deny" },
            @{expression={$_.ShareAppRW}; label="Application RW" },
            @{expression={$_.ShareAppRO}; label="Application RO" },
            @{expression={$_.ShareAppDeny}; label="Application Deny" },
            @{expression={$_.ShareAdmin}; label="Admin" } | 
        Sort-Object -Property "Entità","Gruppo di accesso" |
        Export-Csv -Delimiter $CSVSeparator -Encoding UTF8 -NoTypeInformation "$OutputFolder\$FileName" -Append:$Append;
}

Function Export-PersonalAccountCsv ($Content, $FileName) {
    $lContent = $Content | Where-Object {@('user', 'group') -contains $_.Type -and $_.IsNPA -eq $false}

    #$lContent | Export-Csv -Delimiter $CSVSeparator -Encoding UTF8 -NoTypeInformation "OUTPUT_TST\TEST$FileName";

    [System.Collections.ArrayList]$expContent = @();

    ForEach ($lContentRow in $lContent) {
        if (-not ($lContentRow.Type -eq 'group' -and $null -eq $lContentRow.ParentGroup)) {
            $expContent.Add((Get-UlmDetails -Account $lContentRow))
        }
    }
        <#Select-Object @{expression={$_.Entity}; label="Entità" },
            @{expression={$_.BusinessOwnerName}; label="Business Owner" },
            @{expression={$_.BusinessOwnerEmail}; label="E-mail Business Owner" },
            @{expression={$_.ShareFullPath}; label="Full path" },
            @{expression={$_.ShareAccessGroup}; label="Gruppo di accesso" },
            @{expression={$_.ShareUserRW}; label="User RW" },
            @{expression={$_.ShareUserRO}; label="User RO" },
            @{expression={$_.ShareUserDeny}; label="User Deny" },
            @{expression={$_.ShareAppRW}; label="Application RW" },
            @{expression={$_.ShareAppRO}; label="Application RO" },
            @{expression={$_.ShareAppDeny}; label="Application Deny" },
            @{expression={$_.ShareAdmin}; label="Admin" } -Unique | 
        Sort-Object -Property "Entità","Gruppo di accesso" | #>
    #$lContent | Select-Object AccountName, Type, IsNPA, ContainsNPA, @{expression={$_.Members.AccountName -join ","}; label="Members" } -Unique | Export-Csv -Delimiter $CSVSeparator -Encoding UTF8 -NoTypeInformation "OUTPUT_TST\$FileName";
    $expContent | Export-Csv -Delimiter $CSVSeparator -Encoding UTF8 -NoTypeInformation "$OutputFolder\$FileName";
}

Function Get-UlmDetails($Account) {
    $lAccWithoutDomain = $Account.AccountName.ToString().Substring($Account.AccountName.ToString().IndexOf("\")+1);
    $lAccount = $null

    If ($Account.Type -eq "group") {
        $lAccount = [PSCustomObject]@{
            "AccountName" = $Account.AccountName;
            "Group" = $Account.ParentGroup;
            "FirstName" = $null;
            "LastName" = $null;
            "Email" = $null;
            "Area" = $null;
            "Department" = $null;
            "JobTitle" = $null;
            "ULMModel" = $null;
        }
    }
    Else {
        $UserDetails = $ULM.Where({$_.LoginId -eq $lAccWithoutDomain -or $_.LeaseUser -eq $lAccWithoutDomain -or $_.MasterUser -eq $lAccWithoutDomain -or $_.RootUser -eq $lAccWithoutDomain -or $_.SupportUser -eq $lAccWithoutDomain })
        $lAccount = [PSCustomObject]@{
            "AccountName" = $lAccWithoutDomain;
            "Group" = $Account.ParentGroup;
            "FirstName" = $UserDetails.FirstName;
            "LastName" = $UserDetails.LastName;
            "Email" = $UserDetails.EmailAddress;
            "Area" = $UserDetails.Area;
            "Department" = $UserDetails.Department;
            "JobTitle" = $UserDetails.JobTitle;
            "ULMModel" = $UserDetails.Model;
        }
    }
    
    $lAccount;
}

# Check input file and formats
Function Test-InputParameters {
    # Check ShareBusinessOwnerFile
    If (Test-Path -Path $ShareBusinessOwnerFile) {
        $ShareBusinessOwnerFileRequiredHeaders = @(
            'Share', 'Owner'
        )

        Test-CSVHeaders -FileName $ShareBusinessOwnerFile -RequiredHeaders $ShareBusinessOwnerFileRequiredHeaders
    }
    Else {
        Write-Host "Input file ""$ShareBusinessOwnerFile""  passed as ShareBusinessOwnerFile parameter does not exist, or is not reachable." -ForegroundColor Red
        Exit
    }
    
    # Check AdminGroupsFile
    If (Test-Path -Path $AdminGroupsFile) {
        $AdminGroupsFileRequiredHeaders = @(
            'AdminGroup'
        )

        Test-CSVHeaders -FileName $AdminGroupsFile -RequiredHeaders $AdminGroupsFileRequiredHeaders
    }
    Else {
        Write-Host "Input file ""$AdminGroupsFile""  passed as AdminGroupsFile parameter does not exist, or is not reachable." -ForegroundColor Red
        Exit
    }
    
    # Check NPAFile
    If (Test-Path -Path $NPAFile) {
        $NPAFileRequiredHeaders = @(
            'UserID'
        )

        Test-CSVHeaders -FileName $NPAFile -RequiredHeaders $NPAFileRequiredHeaders
    }
    Else {
        Write-Host "Input file ""$NPAFile""  passed as NPAFile parameter does not exist, or is not reachable." -ForegroundColor Red
        Exit
    }

    # Check ULMFile
    If (-not (Test-Path -Path $ULMFile)) {
        Write-Host "Input file ""$ULMFile""  passed as ULMFile parameter does not exist, or is not reachable." -ForegroundColor Red
        Exit
    }

    # Check OutputFolder, or create it
    If (-not (Test-Path -Path $OutputFolder)) {
        New-Item -ItemType Directory -Force -Path $OutputFolder
    }

    Clear-GC
}

# Validate CSV Headers
Function Test-CSVHeaders ($FileName, $RequiredHeaders) {
    # put all the headers into a comma separated array
    $headers = (Get-Content $FileName | Select-Object -First 1).Split($CSVSeparator)
    foreach ($reqHeader in $RequiredHeaders) {
        if ($headers -notcontains $reqHeader) {
            Write-Host "$FileName failed to validate because it does not contain header  $reqHeader; please check it and try again." -ForegroundColor Red
			
            $error = $true
        }
    }

    if ($error -eq $true) {
        Exit
    }

    Clear-GC
}

Function Clear-GC {
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}

# Start script
Measure-Command {
. Main
}