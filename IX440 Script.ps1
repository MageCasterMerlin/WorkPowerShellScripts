
# Define the folder containing the four Excel workbooks
$folderPath = "$env:USERPROFILE\Downloads\IX440 Report"

# Define the input workbooks
$PX110_File = Join-Path $folderPath "PX110.csv"
$IX440_STORE_File = Join-Path $folderPath "IX440 STORE.csv"
$IX440_SPD_File = Join-Path $folderPath "IX440 SPD.csv"

# Define the output workbook
$dateString = Get-Date -Format "MM.dd.yyyy"
$outputFile = Join-Path $folderPath "IX440 $dateString.xlsx"

# Columns to delete (right to left to avoid shifting issues)
$columnsToDelete = @("Y", "X", "W", "V", "U", "T", "S", "R", "P", "O", "N", "J", "I", "E", "A")

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
        # Assign yesterday's date, or last Friday if today is Monday
    if ((Get-Date).DayOfWeek -eq 'Monday') {
        $yesterDateString = (Get-Date).AddDays(-3).Date
    }
    else {
        $yesterDateString = (Get-Date).AddDays(-1).Date
    }
    $yesterDateString = $yesterDateString.ToString("MM.dd.yyyy")

    # Create new workbook
    $newWorkbook = $Excel.Workbooks.Add()

    # ==================================
    # COPY PX110 DATA
    # ==================================

    # Open source workbook and sheet
    $sourcePx110Workbook = $Excel.Workbooks.Open($PX110_File)
    $sourceSheet = $sourcePx110Workbook.Worksheets.Item(1)

    # Get the first worksheet in the new workbook and rename
    $newSheet = $newWorkbook.Worksheets.Item(1)
    $newSheet.Name = "PX110"

    # Copy all data from source worksheet
    $sourceSheet.UsedRange.Copy()

    # Paste data and formats into new worksheet
    $newSheet.Range("A1").PasteSpecial(-4163) # xlPasteValues
    $newSheet.Range("A1").PasteSpecial(-4122) # xlPasteFormats
    
        # Move Column I to Column A
    $newSheet.Columns.Item("I").Cut()
    $newSheet.Columns.Item("A").Insert()

    # Move Column O to Column C
    $newSheet.Columns.Item("O").Cut()
    $newSheet.Columns.Item("C").Insert()

    # Move Column AE to Column D
    $newSheet.Columns.Item("AE").Cut()
    $newSheet.Columns.Item("D").Insert()

     # Determine used range of data
    $lastRow = $newSheet.UsedRange.Row + $newSheet.UsedRange.Rows.Count - 1

    # Runs Text-to-Column in Column A
    $rangeA = $newSheet.Range("A2:A$lastRow")
    $rangeA.TextToColumns(
        $rangeA,
        1,       # xlDelimited
        1,       # xlTextQualifierDoubleQuote
        $false,  # ConsecutiveDelimiter
        $true,   # Tab delimiter
        $false,  # Semicolon
        $false,  # Comma
        $false,  # Space
        $false   # Other
    )
    
    # ==================================
    # COPY IX440 STORE DATA
    # ==================================

    # Open source workbook and sheet
    $sourceIx440StoreWorkbook = $Excel.Workbooks.Open($IX440_STORE_File)
    $sourceSheet = $sourceIx440StoreWorkbook.Worksheets.Item(1)

    # Make a new worksheet in the new workbook and rename
    $newSheet = $newWorkbook.Worksheets.Add()
    $newSheet.Name = "IX440 STORE"

    # Copy all data from source worksheet
    $sourceSheet.UsedRange.Copy()

    # Paste data and formats into new worksheet
    $newSheet.Range("A1").PasteSpecial(-4163) # xlPasteValues
    $newSheet.Range("A1").PasteSpecial(-4122) # xlPasteFormats
    
    # Delete columns
    foreach ($column in $columnsToDelete) {
        $newSheet.Columns.Item($column).Delete()
    }

    # Move Column B to Column A
    $newSheet.Columns.Item("B").Cut()
    $newSheet.Columns.Item("A").Insert()

    # Determine used range of data
    $lastRow = $newSheet.UsedRange.Row + $newSheet.UsedRange.Rows.Count - 1

    # Runs Text-to-Column in Column A
    $rangeA = $newSheet.Range("A2:A$lastRow")
    $rangeA.TextToColumns(
        $rangeA,
        1,       # xlDelimited
        1,       # xlTextQualifierDoubleQuote
        $false,  # ConsecutiveDelimiter
        $true,   # Tab delimiter
        $false,  # Semicolon
        $false,  # Comma
        $false,  # Space
        $false   # Other
    )

    # Insert PX110 Comments Vlookup
    $newSheet.Range("K1").Value2 = "PX110 COMMENTS"
    $newSheet.Range("K2").Value2 = "=VLOOKUP(A2,'PX110'!A:D,4,FALSE)"

    # Select column K from K2 through last used row
    $rangeK = $newSheet.Range("K2:K$lastRow")
    $rangeK.FillDown()

    # Insert IX440 Previos Comments Vlookup
    $newSheet.Range("L1").Value2 = "IX440 PREVIOUS COMMENTS"
    $newSheet.Range("L2").Value2 = "=VLOOKUP(A2,'[IX440 $yesterDateString.xlsx]IX440 STORE'!`$A:`$L,12,FALSE)"
   
    # Select column L from L2 through last used row
    $rangeL = $newSheet.Range("L2:L$lastRow")
    $rangeL.FillDown()

    # Insert Review Date VLOOKUP
    $newSheet.Range("M1").Value2 = "REVIEW DATE"
    $newSheet.Range("M2").Value2 = "=VLOOKUP(A2,'[IX440 $yesterDateString.xlsx]IX440 STORE'!`$A:`$M,13,FALSE)"
    
    # Select column M from M2 through last used row
    $rangeM = $newSheet.Range("M2:M$lastRow")
    $rangeM.FillDown()
    
    foreach ($cell in $rangeK.Cells) {
        if ($cell.Value2 -is [double] -and $cell.Value2 -eq 0) {
            $cell.ClearContents()
        }
    }

    # ==================================
    # COPY IX440 SPD DATA
    # ==================================

    # Open source workbook and sheet
    $sourceIx440SpdWorkbook  = $Excel.Workbooks.Open($IX440_SPD_File)
    $sourceSheet = $sourceIx440SpdWorkbook.Worksheets.Item(1)

    # Make a new worksheet in the new workbook and rename
    $newSheet = $newWorkbook.Worksheets.Add()
    $newSheet.Name = "IX440 SPD"

    # Copy all data from source worksheet
    $sourceSheet.UsedRange.Copy()

    # Paste data and formats into new worksheet
    $newSheet.Range("A1").PasteSpecial(-4163) # xlPasteValues
    $newSheet.Range("A1").PasteSpecial(-4122) # xlPasteFormats

    # Delete columns
    foreach ($column in $columnsToDelete) {
        $newSheet.Columns.Item($column).Delete()
    }

    # Move Column B to Column A
    $newSheet.Columns.Item("B").Cut()
    $newSheet.Columns.Item("A").Insert()

    # Determine used range of data
    $lastRow = $newSheet.UsedRange.Row + $newSheet.UsedRange.Rows.Count - 1

    # Runs Text-to-Column in Column A
    $rangeA = $newSheet.Range("A2:A$lastRow")
    $rangeA.TextToColumns(
        $rangeA,
        1,       # xlDelimited
        1,       # xlTextQualifierDoubleQuote
        $false,  # ConsecutiveDelimiter
        $true,   # Tab delimiter
        $false,  # Semicolon
        $false,  # Comma
        $false,  # Space
        $false   # Other
    )

    # Insert PX110 Vlookup
    $newSheet.Range("K1").Value2 = "PX110 COMMENTS"
    $newSheet.Range("K2").Value2 = "=VLOOKUP(A2,'PX110'!A:D,4,FALSE)"

    # Select column K from K2 through last used row
    $rangeK = $newSheet.Range("K2:K$lastRow")
    $rangeK.FillDown()

    # Insert IX440 Previos Comments Vlookup
    $newSheet.Range("L1").Value2 = "IX440 PREVIOUS COMMENTS"
    $newSheet.Range("L2").Value2 = "=VLOOKUP(A2,'[IX440 $yesterDateString.xlsx]IX440 SPD'!`$A:`$L,12,FALSE)"

    # Select column L from L2 through last used row
    $rangeL = $newSheet.Range("L2:L$lastRow")
    $rangeL.FillDown()
    
    # Insert Review Date VLOOKUP
    $newSheet.Range("M1").Value2 = "REVIEW DATE"
    $newSheet.Range("M2").Value2 = "=VLOOKUP(A2,'[IX440 $yesterDateString.xlsx]IX440 SPD'!`$A:`$M,13,FALSE)"
    
    # Select column M from M2 through last used row
    $rangeM = $newSheet.Range("M2:M$lastRow")
    $rangeM.FillDown()
    
    foreach ($cell in $rangeK.Cells) {
        if ($cell.Value2 -is [double] -and $cell.Value2 -eq 0) {
            $cell.ClearContents()
        }
    }
    # ==================================
    # Close Out Workbook
    # ==================================
       
    # Save new workbook
    $newWorkbook.SaveAs($outputFile, 51) # 51 = Excel .xlsx format

    Write-Host "Successfully created: $outputFile"
}

catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    # Close workbooks
    if ($newWorkbook) {
        $newWorkbook.Close($true)
    }

    if ($sourcePx110Workbook) {
        $sourcePx110Workbook.Close($false)
    }
    if ($sourceIx440StoreWorkbook) {
        $sourceIx440StoreWorkbook.Close($false)
    }
    if ($sourceIx440SpdWorkbook) {
        $sourceIx440SpdWorkbook.Close($false)
    }
    # Close Excel
    $Excel.Quit()


    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
