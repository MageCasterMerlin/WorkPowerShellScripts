
# Define the folder containing the four Excel workbooks
$folderPath = "C:\Users\FXS404\Downloads\Reports"

# Define the output workbook
$dateString = Get-Date -Format "MM.dd.yyyy"
$OutputFile = Join-Path $PSScriptRoot "IX440 $dateString.xlsx"

# List of input workbooks
$workbooks = @("IX440 SPD.csv", "IX440 Storeroom.csv", "IX112.csv", "PX110.csv")

# Columns to delete (right to left to avoid shifting issues)
$columnsToDelete = @("Y", "X", "W", "V", "T", "S", "R", "P", "O", "N", "J", "I", "E", "A")
# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$outputWorkbook = $null

try {
    # Create a new workbook to hold the combined sheets
    $outputWorkbook = $excel.Workbooks.Add()

    # Remove the default worksheets, keeping one temporarily
    while ($outputWorkbook.Worksheets.Count -gt 1) {
        $outputWorkbook.Worksheets.Item(
            $outputWorkbook.Worksheets.Count
        ).Delete()
    }

    foreach ($file in $workbooks) {
        $inputPath = Join-Path $folderPath $file

        if (-not (Test-Path $inputPath)) {
            Write-Warning "File not found: $inputPath"
            continue
        }

        # Open the source workbook as read-only
        $sourceWorkbook = $excel.Workbooks.Open(
            $inputPath, 0, $true
        )

        try {
            # Copy each worksheet into the output workbook
            foreach ($sourceSheet in @($sourceWorkbook.Worksheets)) {
                $sourceSheet.Copy(
                    [System.Reflection.Missing]::Value,
                    $outputWorkbook.Worksheets.Item(
                        $outputWorkbook.Worksheets.Count
                    )
                )

                # Give the copied worksheet a unique name
                $newSheet = $outputWorkbook.Worksheets.Item(
                    $outputWorkbook.Worksheets.Count
                )

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension(
                    $file
                )
                $sheetName = $baseName
                $counter = 1

                while (@($outputWorkbook.Worksheets | Where-Object {
                    $_.Name -eq $sheetName
                }).Count -gt 0) {
                    $sheetName = "$baseName$counter"
                    $counter++
                }

                $newSheet.Name = $sheetName
            }
        }

        finally {
            $sourceWorkbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $sourceWorkbook
            ) | Out-Null
        }
    }

    # Remove the original blank worksheet if it is still present
    if ($outputWorkbook.Worksheets.Count -gt 1) {
        foreach ($sheet in @($outputWorkbook.Worksheets)) {
            if ($sheet.UsedRange.Count -eq 1 -and
                $sheet.UsedRange.Value2 -eq $null) {
                $sheet.Delete()
                break
            }
        }
    }

    # Save the combined workbook
    $outputWorkbook.SaveAs($outputFile, 51)

    Write-Host ""
    Write-Host "Successfully combined workbooks into:"
    Write-Host $outputFile

    # Delete Columns in 440 SPD
    $worksheet = $outputWorkbook.Worksheets.Item("IX440 SPD")
    
    foreach ($column in $columnsToDelete) {
        $worksheet.Columns.Item($column).Delete()
    }

    # Delete Columns in 440 Storeroom
    $worksheet = $outputWorkbook.Worksheets.Item("IX440 Storeroom")
    
    foreach ($column in $columnsToDelete) {
        $worksheet.Columns.Item($column).Delete()
    }

    Write-Host "Successfully deleted Columns"
}

catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    if ($null -ne $outputWorkbook) {
        $outputWorkbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $outputWorkbook
        ) | Out-Null
    }

    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
        $excel
    ) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
