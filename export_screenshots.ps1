# Generate screenshots of the Xbox dashboard using Excel COM
$excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$dashPath = "C:\Users\conta\DIO\xbox-dashboard-excel\dashboard_xbox_vendas.xlsx"
$outDir   = "C:\Users\conta\DIO\xbox-dashboard-excel"

Write-Host "Opening Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($dashPath)
    Write-Host "Workbook opened: $($wb.Sheets.Count) sheets"

    # Export each sheet as PDF for screenshot
    foreach ($ws in $wb.Sheets) {
        $name = $ws.Name
        if ($name -eq '_data') { continue }  # skip hidden data sheet
        
        $pdfPath = "$outDir\$name.pdf"
        Write-Host "Exporting $name -> $pdfPath"
        
        # Export as PDF (0 = xlTypePDF)
        $ws.ExportAsFixedFormat(0, $pdfPath)
        
        Write-Host "  OK: $pdfPath"
    }
    
    $wb.Save()
    $wb.Close()
    Write-Host "All exports done!"
}
catch {
    Write-Host "ERROR: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
