function searchRedExcel($path) {
    if(-not (Test-path $path)){
        Write-Error "$path is Not Found"
        return $path
    }

    $sheetName = "Sheet1"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks
    $book = $workbook.Open($path, 0, $true)
    $sheets = $book.Worksheets
    for($sheetidx=1; $sheetidx -le $sheets.Count; $sheetidx++){
        $sheet = $sheets.Item($sheetidx)
        Write-Host $sheet.Name
        $cells = $sheet.Cells
        $usedRange = $sheet.UsedRange
        $rows = $usedRange.Rows

        Write-Host ("rows count:" + $rows.Count)
        for($ridx=1; $ridx -le $rows.Count; $ridx++){
            $row = $rows.Item($ridx)
            Write-Host ("cols count:" + $row.Count)
            for($cidx=1; $cidx -le $row.Count; $cidx++){
                $cell = $cells.Item(2, 2)
                $text = $cell.Text
                
                $chars = $cell.Characters(3,1)
                $font = $chars.font
                
                # Write-host $font.colorIndex + ":" + $font.color
                # write-host ("color: " + $cell.font.ColorIndex)
                # Write-Host $text
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($font)
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($chars)
                [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($cell)
            }
            [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($row)
        }
        
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($rows)
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($usedRange)
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($cells)
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($sheet)
    }
    
    [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($sheets)
    [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($book)
    [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($workbook)
} catch {
    Write-Error $_.Exception
}
try{
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
} catch {
    Write-Error $_.Exception
}
[gc]::Collect()
}

searchRedExcel (convert-path C:\Users\sumit\Documents\search_red.xlsx)