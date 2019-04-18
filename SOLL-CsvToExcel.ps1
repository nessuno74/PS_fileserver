
#commento
#dfgvdsfvs
#ssdfsdf

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Excel file to use as template for Owner SOLL report")]
    [String]
    $excelSOLLTemplateOwner = "00-SOLL-OwnerTemplate.xlsx",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Excel file to use as template for Technical SOLL report")]
    [String]
    $excelSOLLTemplateTech = "00-SOLL-TechTemplate.xlsx",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="CSV file containing ULM export with username and user details")]
    [String]
    $ULMFile = "ULM.csv",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Input folder in which to search for _ShareFunctionalMatrix.csv and _SubjectMatrixPA.csv files")]
    [String]
    $InputFolder = "SOLL-OUTPUT",
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Folder in which to save Excel files")]
    [String]
    $OutputFolder = $InputFolder,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage="Set which type of SOLL should be generated: Owner or Technical")]
    [ValidateSet('Owner','Technical')]
    [String]
    $SollType = 'Owner'
) #Accept input parameters

$CSVSeparator = ";"
[System.Collections.ArrayList]$PAFilesAlreadyElaborated = @();

# Main function
Function Main {
    Test-InputParameters

    $csvSOLLFiles = Get-ChildItem $InputFolder | Where-Object {$_.Name.EndsWith("_ShareFunctionalMatrix.csv")}
    $csvPAFiles = Get-ChildItem $InputFolder | Where-Object {$_.Name.EndsWith("_SubjectMatrixPA.csv")}
    [System.Collections.ArrayList]$ULM = Import-Csv $ULMFile -Delimiter $CSVSeparator | Select-Object -Unique LoginId, MasterUser, RootUser, SupportUser, LastName, FirstName, Area, Department, JobTitle, EmailAddress, Model, CorporateKey

    # Convert from ShareFunctionalMatrix first
    foreach ($sollFile in $csvSOLLFiles) {
        Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))..."

        #If ($SollType -eq 'Owner') {
            $FunctionalMatrix = @(Import-Csv -Path $sollFile.FullName -Delimiter $CSVSeparator | Where-Object ({-not $_."Gruppo di accesso".StartsWith("S-1-5-21") -and @("CREATOR OWNER", "NT AUTHORITY\SYSTEM") -notcontains $_."Gruppo di accesso" -and $_."Admin" -ne "X"}))
            #$FunctionalMatrix = @(Import-Csv -Path $sollFile.FullName -Delimiter $CSVSeparator | Where-Object ({-not $_."Gruppo di accesso".StartsWith("S-1-5-21") -and @("CREATOR OWNER", "NT AUTHORITY\SYSTEM") -notcontains $_."Gruppo di accesso" -and $_."Admin" -ne "X" -or $null -ne $_."Application RW" -or $null -ne $_."Application RO" -or $null -ne $_."Application Deny"}}))
        #}
        
        $Excel = New-Object -ComObject excel.application

        $Excel.Visible = $false

        $ExcelWorkBook = $Excel.Workbooks.Open((Get-Item $excelSOLLTemplateOwner))
        
        Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))..." -Status "Compiling sheet 'Share Functional Matrix'"

        Update-ShareFunctionalMatrixWorkSheet -WorkBook $ExcelWorkBook -WorkSheetContent $FunctionalMatrix

        If ($csvPAFiles.FullName -contains $sollFile.FullName.Replace('_ShareFunctionalMatrix.csv', '_SubjectMatrixPA.csv')) {
            $PAMatrix = @(Import-Csv -Path $sollFile.FullName.Replace('_ShareFunctionalMatrix.csv', '_SubjectMatrixPA.csv') -Delimiter ";" | Where-Object {$_.IsAdmin -ne $true})
            
            Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))..." -Status "Compiling sheet 'Subject Matrix - PA'"
            Update-SubjectMatrixPAWorkSheet -WorkBook $ExcelWorkBook -WorkSheetContent $PAMatrix

            $PAFilesAlreadyElaborated.Add($sollFile.FullName.Replace('_ShareFunctionalMatrix.csv', '_SubjectMatrixPA.csv'))
            
            Remove-Variable -Name PAMatrix
        }

        Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))..." -Status "Compiling cover page"
        
        $OwnerDetails = $ULM.Where({$_.EmailAddress -eq $FunctionalMatrix[0].'E-mail Business Owner'});

        Update-CoverPageWorkSheet -WorkBook $ExcelWorkBook -OwnerDetails $OwnerDetails
        
        Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))..." -Status "Save to xlsx file"

        $ExcelWorkBook.SaveAs("$((Get-Item $OutputFolder).FullName)\$($sollFile.Name.Replace('_ShareFunctionalMatrix.csv', '_PersonalAccounts.xlsx'))");
        $ExcelWorkBook.Close()
        $Excel.Quit()
        
        Remove-Variable -Name Excel
        Remove-Variable -Name ExcelWorkBook
        Remove-Variable -Name FunctionalMatrix
        Remove-Variable -Name OwnerDetails
            
        Clear-GC
    }

    foreach ($paFile in $csvPAFiles) {
        If (-not $PAFilesAlreadyElaborated -contains $paFile.FullName) {
            Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_SubjectMatrixPA.csv', '_PersonalAccounts.xlsx'))..."

            #If ($SollType -eq 'Owner') {
                $PAMatrix = Import-Csv -Path $paFile.FullName -Delimiter $CSVSeparator
            #}
            
            $Excel = New-Object -ComObject excel.application

            $Excel.Visible = $false

            $ExcelWorkBook = $Excel.Workbooks.Open((Get-Item $excelSOLLTemplateOwner))

            Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_SubjectMatrixPA.csv', '_PersonalAccounts.xlsx'))..." -Status "Compiling sheet 'Subject Matrix - PA'"

            Update-SubjectMatrixPAWorkSheet -WorkBook $ExcelWorkBook -WorkSheetContent $PAMatrix

            $OwnerDetails = $ULM.Where({$_.EmailAddress -eq $row.'E-mail Business Owner'});

            Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_SubjectMatrixPA.csv', '_PersonalAccounts.xlsx'))..." -Status "Compiling cover page"

            Update-CoverPageWorkSheet -WorkBook $ExcelWorkBook -OwnerDetails $OwnerDetails
            
            Write-Progress "Create file $OutputFolder\$($sollFile.Name.Replace('_SubjectMatrixPA.csv', '_PersonalAccounts.xlsx'))..." -Status "Save to xlsx file"

            $ExcelWorkBook.Saveas("$((Get-Item $OutputFolder).FullName)\$($paFile.Name.Replace('_SubjectMatrixPA.csv', '_PersonalAccounts.xlsx'))");
            $ExcelWorkBook.Close()
            $Excel.Quit()
            
            Remove-Variable -Name Excel
            Remove-Variable -Name ExcelWorkBook
            Remove-Variable -Name PAMatrix
            Remove-Variable -Name OwnerDetails
            
            Clear-GC
        }
    }
}

function Update-ShareFunctionalMatrixWorkSheet {
    param (
        $WorkBook,
        $WorkSheetContent
    )

    $ExcelWorkSheet = $WorkBook.WorkSheets.item('Share Functional Matrix')
    $ExcelWorkSheet.activate()

    $lastRow = 1

    $row = $null

    Foreach ($row in $WorkSheetContent) {
        $lastRow += 1
        $ExcelWorkSheet.cells.Item($lastRow,1) = $row.'Entità'
        $ExcelWorkSheet.cells.Item($lastRow,2) = $row.'Business Owner'
        $ExcelWorkSheet.cells.Item($lastRow,3) = $row.'E-mail Business Owner'
        $ExcelWorkSheet.cells.Item($lastRow,4) = $row.'Full path'
        $ExcelWorkSheet.cells.Item($lastRow,5) = $row.'Gruppo di accesso'
        $ExcelWorkSheet.cells.Item($lastRow,6) = $row.'User RW'
        $ExcelWorkSheet.cells.Item($lastRow,7) = $row.'User RO'
        $ExcelWorkSheet.cells.Item($lastRow,8) = $row.'User Deny'
        $ExcelWorkSheet.cells.Item($lastRow,9) = $row.'Application RW'
        $ExcelWorkSheet.cells.Item($lastRow,10) = $row.'Application RO'
        $ExcelWorkSheet.cells.Item($lastRow,11) = $row.'Application Deny'
        $ExcelWorkSheet.cells.Item($lastRow,12) = $row.Admin
    }

    If ($lastRow -lt 2) {
        $lastRow = 2
    }

    $range = $ExcelWorkSheet.Range("A2",("E{0}"  -f $lastRow))
    $range.Select() | Out-Null

    $range.Font.Size = 8
    $range.Font.Name = "Arial"
    $range.Font.Bold = $true
    $range.VerticalAlignment = -4108
    $range.Interior.ColorIndex =48
    $range.Borders.Item(12).LineStyle = 1
    $range.Borders.Item(7).LineStyle = 1
    $range.Borders.Item(10).LineStyle = 1
    $range.Borders.Item(9).LineStyle = 1

    $range = $ExcelWorkSheet.Range("F2",("L{0}"  -f $lastRow))
    $range.Select() | Out-Null

    $range.Font.Size = 8
    $range.Font.Name = "Arial"
    $range.Font.Bold = $false
    $range.HorizontalAlignment = -4108
    $range.VerticalAlignment = -4108
    $range.Interior.ColorIndex = 36
    $range.Borders.Item(12).LineStyle = 1
    $range.Borders.Item(11).LineStyle = 1
    $range.Borders.Item(7).LineStyle = 1
    $range.Borders.Item(10).LineStyle = 1
    $range.Borders.Item(9).LineStyle = 1

    $ExcelWorkSheet.cells.Item(2,1).Select() | Out-Null

    Clear-GC
}

function Update-SubjectMatrixPAWorkSheet {
    param (
        $WorkBook,
        $WorkSheetContent
    )
    
    $ExcelWorkSheet = $WorkBook.WorkSheets.item('Subject Matrix - PA')
    $ExcelWorkSheet.activate()

    $lastRow = 1

    $row = $null

    Foreach ($row in $WorkSheetContent) {
        if ($SollType -eq 'Owner') {
            $lastRow += 1
            $ExcelWorkSheet.cells.Item($lastRow,1) = $row.AccountName
            $ExcelWorkSheet.cells.Item($lastRow,2) = $row.Group
            $ExcelWorkSheet.cells.Item($lastRow,3) = $row.FirstName
            $ExcelWorkSheet.cells.Item($lastRow,4) = $row.LastName
            $ExcelWorkSheet.cells.Item($lastRow,5) = if ($row.Email -eq "null") { $null } else { $row.Email }
            $ExcelWorkSheet.cells.Item($lastRow,6) = if ($row.Area -eq "null") { $null } else { $row.Area }
            $ExcelWorkSheet.cells.Item($lastRow,7) = if ($row.Department -eq "null") { $null } else { $row.Department }
            $ExcelWorkSheet.cells.Item($lastRow,8) = if ($row.JobTitle -eq "null") { $null } else { $row.JobTitle }
            $ExcelWorkSheet.cells.Item($lastRow,9) = if ($row.ULMModel -eq "null") { $null } else { $row.ULMModel }
        }
    }

    If ($lastRow -lt 2) {
        $lastRow = 2
    }

    $range = $ExcelWorkSheet.Range("A2",("I{0}"  -f $lastRow))

    $range.Font.Size = 8
    $range.Font.Name = "Arial"
    $range.Font.Bold = $false
    $range.VerticalAlignment = -4108
    $range.Interior.ColorIndex = 36
    $range.Borders.Item(12).LineStyle = 1
    $range.Borders.Item(11).LineStyle = 1
    $range.Borders.Item(7).LineStyle = 1
    $range.Borders.Item(10).LineStyle = 1
    $range.Borders.Item(9).LineStyle = 1

    $ExcelWorkSheet.cells.Item(2,1).Select() | Out-Null

    Clear-GC
}

function Update-CoverPageWorkSheet {
    param (
        $WorkBook,
        $OwnerDetails
    )
    
    $ExcelWorkSheet = $WorkBook.WorkSheets.item('Cover page')
    $ExcelWorkSheet.activate()
    $ExcelWorkSheet.cells.Item(19,3) = $OwnerDetails.FirstName + " " + $OwnerDetails.LastName
    $ExcelWorkSheet.cells.Item(20,3) = $OwnerDetails.CorporateKey
    $ExcelWorkSheet.cells.Item(21,3) = $OwnerDetails.JobTitle
    $ExcelWorkSheet.cells.Item(22,3) = $OwnerDetails.Area
    $ExcelWorkSheet.cells.Item(23,3) = $OwnerDetails.Department
    
    $CurDate = (Get-Date -UFormat "%m/%d/%Y").ToString()

    $ExcelWorkSheet.cells.Item(17,3) = $CurDate
    $ExcelWorkSheet.cells.Item(25,3) = $CurDate
    
    Remove-Variable -Name CurDate
    Clear-GC
}


# Check input file and formats
Function Test-InputParameters {
    # Check Excel template
    If (-not (Test-Path -Path $excelSOLLTemplateOwner)) {
        Write-Host "Excel template file ""$excelSOLLTemplateOwner""  passed as excelSOLLTemplateOwner parameter does not exist, or is not reachable." -ForegroundColor Red
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

function Clear-GC {
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}

# Start script
Measure-Command {
    . Main
}