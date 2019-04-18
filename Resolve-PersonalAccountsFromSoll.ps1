[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing ULM export with username and user details")]
    [String]
    $ULMFile = "ULM.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing non personal accounts")]
    [String]
    $NPAFile = "NPA.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV containing all AD groups that should be flagged as admin")]
    [String]
    $AdminGroupsFile = "admin-groups.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Input folder in which to search for _ShareFunctionalMatrix.csv files")]
    [String]
    $InputFolder = "SOLL-OUTPUT",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Folder in which to save output CSV files")]
    [String]
    $OutputFolder = $InputFolder,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Indicate the server on which to search for local (builtin) groups")]
    [String]
    $LocalGroupsResolveTo
) #Accept input parameters

$CSVSeparator = ";"

$NoResolveADGroups = @(
    "Domain Users",
    "Authenticated Users",
    "Users",
    "Everyone"
)

# Declare main variables
[System.Collections.ArrayList]$ULM = @();
[System.Collections.ArrayList]$PersonalAccounts = @();
[System.Collections.ArrayList]$SollAccounts = @();
$Searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'');

# Main function
Function Main {
    Test-InputParameters

    [System.Collections.ArrayList]$ULM = Import-Csv $ULMFile -Delimiter $CSVSeparator | Select-Object -Unique LoginId, MasterUser, RootUser, SupportUser, LastName, FirstName, Area, Department, JobTitle, EmailAddress, Model

    $NonPersonalAccounts = @()
    Import-Csv $NPAFile -Delimiter $CSVSeparator | ForEach-Object { $NonPersonalAccounts += $_.UserID }

    $AdminGroups = @()
    Import-Csv $AdminGroupsFile -Delimiter $CSVSeparator | ForEach-Object { $AdminGroups += $_.AdminGroup }

    $csvSOLLFiles = Get-ChildItem $InputFolder | Where-Object {$_.Name.EndsWith("_ShareFunctionalMatrix.csv")}

    Foreach ($csvFile in $csvSOLLFiles) {
        Write-Progress "Elaborating file $($csvFile)..." -Status "Importing rows"
        $SollAccounts = @(Import-Csv $csvFile.FullName -Delimiter $CSVSeparator | Select-Object -Unique 'Gruppo di accesso')

        ForEach ($acc in $SollAccounts) {
            Write-Progress "Elaborating file $($csvFile)..." -Status "Resolve group $($acc.'Gruppo di accesso')"
            Add-PersonalAccount -Account $acc.'Gruppo di accesso'
        }

        Write-Progress "Elaborating file $($csvFile)..." -Status "Export to file $OutputFolder\$($csvFile.Name.Replace('_ShareFunctionalMatrix.csv', '_SubjectMatrixPA.csv'))"
        Export-PersonalAccountCsv -Content $PersonalAccounts -FileName $csvFile.Name.Replace('_ShareFunctionalMatrix.csv', '_SubjectMatrixPA.csv')
        
        $PersonalAccounts.Clear()

        Clear-GC
    }
}

Function Add-PersonalAccount($Account, $ParentGroup = $null, $IsAdmin = $false) {
    If (-not $Account.ToString().StartsWith("S-1-5-21")) {        
        $PA = $PersonalAccounts.Where({ $_.AccountName -eq $Account })
        If (-not $PA) {
            $accIsAdmin = if($AdminGroups -contains $Account.ToString() -or $AdminGroups -contains $accWithoutDomain -or $IsAdmin -eq $true) { $true } else { $false };
            $accWithoutDomain = $Account.ToString().Substring($Account.ToString().IndexOf("\")+1)
            
            if ($NoResolveADGroups -contains $accWithoutDomain) {
                $TempPA = [PSCustomObject]@{
                    "AccountName" = $Account;
                    "Type" = $group;
                    "IsNPA" = $false;
                    "ContainsNPA" = $false;
                    "IsAdmin" = $accIsAdmin;
                    "ParentGroup" = $ParentGroup;
                }
            }
            else {
                # If account is local group then resolve it from server
                if($Account.ToString().StartsWith("BUILTIN\")) {
                    $lGroup = [ADSI]"WinNT://$($LocalGroupsResolveTo)/$($accWithoutDomain)"
                    $lGroupMembers = @($lGroup.PSBase.Invoke("Members"))
                    foreach ($lGroupMember in $lGroupMembers) {
                        $lGroupMemberName = $lGroupMember.GetType().InvokeMember("Name", 'GetProperty', $null, $lGroupMember, $null)
                        $lGroupMemberPath = $lGroupMember.GetType().InvokeMember("AdsPath", 'GetProperty', $null, $lGroupMember, $null)
                        
                        if ($lGroupMemberPath -like "*$($LocalGroupsResolveTo.Substring(0,$LocalGroupsResolveTo.IndexOf(".")))*") {
                            $lGroupMemberAccount = $lGroupMemberName

                        }
                        else {
                            $lGroupMemberSid = $lGroupMember.GetType().InvokeMember("objectSid", 'GetProperty', $null, $lGroupMember, $null)
                            $lGroupMemberSidAccount = New-Object Security.Principal.SecurityIdentifier ($lGroupMemberSid, 0)
                            $lGroupMemberAccount = $lGroupMemberSidAccount.Translate([Security.Principal.NTAccount]).Value
                        }
                        
                        If ($NonPersonalAccounts -contains $lGroupMemberName) {
                            $lContainsNpa = $true
                        }

                        Add-PersonalAccount -Account $lGroupMemberAccount -ParentGroup $Account.ToString() -IsAdmin $accIsAdmin | Out-Null
                    }

                    $TempPA = [PSCustomObject]@{
                        "AccountName" = $Account;
                        "Type" = "group"
                        "IsNPA" = If ( $lGroupClass -ne "Group" -and ($NonPersonalAccounts -contains $Account -or $NonPersonalAccounts -contains $accWithoutDomain) ) { $true } Else { $false };
                        "ContainsNPA" = $lContainsNpa;
                        "IsAdmin" = $accIsAdmin;
                        "ParentGroup" = $ParentGroup;
                    }
                }
                else {
                    $Searcher.filter = "(&(sAMAccountName= $($accWithoutDomain))(!userAccountControl:1.2.840.113556.1.4.803:=2))"

                    $lAccount = $Searcher.FindOne()

                    If ($null -ne $lAccount) {
                        $lType = $lAccount.Properties['objectClass']

                        $lMembers = $null
                    
                        $lContainsNpa = $false

                        If ( $lType -contains "group" ) {
                            $lType = "group"
                            $lMembers = @()

                            ForEach ($lMember in $lAccount.Properties["member"]) {
                                $lGroupMember = [ADSI]"LDAP://$lMember"
                                $lMembers += $lGroupMember
                            
                                If ($NonPersonalAccounts -contains $lGroupMember.Properties["sAMAccountName"]) {
                                    $lContainsNpa = $true
                                }

                                If (@('group', 'user') -contains $lGroupMember.SchemaClassName) {
                                    Add-PersonalAccount -Account $lGroupMember.Properties["sAMAccountName"].Value -ParentGroup $Account.ToString() -IsAdmin $accIsAdmin | Out-Null
                                }
                            }
                        }
                        else {
                            $lType = "user"
                        }
                    
                        $TempPA = [PSCustomObject]@{
                            "AccountName" = $Account;
                            "Type" = $lType;
                            "IsNPA" = If ( $lType -eq "user" -and ($NonPersonalAccounts -contains $Account -or $NonPersonalAccounts -contains $accWithoutDomain) ) { $true } Else { $false };
                            "ContainsNPA" = $lContainsNpa;
                            "IsAdmin" = $accIsAdmin;
                            "ParentGroup" = $ParentGroup;
                        }
                    }
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
                    "IsAdmin" = if($obj.IsAdmin -eq $true -or $IsAdmin -eq $true) { $true } else { $null };
                    "ParentGroup" = $ParentGroup;
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
            "IsAdmin" = $Account.IsAdmin;
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
        $UserDetails = $ULM.Where({$_.LoginId -eq $lAccWithoutDomain -or $_.MasterUser -eq $lAccWithoutDomain -or $_.RootUser -eq $lAccWithoutDomain -or $_.SupportUser -eq $lAccWithoutDomain })
        $lAccount = [PSCustomObject]@{
            "AccountName" = $lAccWithoutDomain;
            "Group" = $Account.ParentGroup;
            "IsAdmin" = $Account.IsAdmin;
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
        if ($headers -cnotcontains $reqHeader) {
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