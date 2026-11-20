Write-Host "=========================================================="
Write-Host "CONVERTING ALL PROJECT DELIVERABLES TO PDF VIA MS OFFICE"
Write-Host "=========================================================="

$docsDir = Resolve-Path "docs"
Write-Host "Target Directory: $docsDir"

# 1. Convert Word (.docx) files
$docxFiles = Get-ChildItem -Path $docsDir -Filter "*.docx"
Write-Host "Found $($docxFiles.Count) Word documents to convert."

if ($docxFiles.Count -gt 0) {
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

        foreach ($file in $docxFiles) {
            $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
            Write-Host "Converting Word: $($file.Name) -> $([System.IO.Path]::GetFileName($pdfPath))..."
            
            try {
                $doc = $word.Documents.Open($file.FullName, $false, $true)
                # 17 = wdFormatPDF
                $doc.SaveAs([ref]$pdfPath, [ref]17)
                $doc.Close($false)
                Write-Host "  [OK] Saved: $([System.IO.Path]::GetFileName($pdfPath))"
            } catch {
                Write-Host "  [ERROR] Failed to convert $($file.Name): $($_.Exception.Message)"
            }
        }
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    } catch {
        Write-Host "[ERROR] Word COM instantiation failed: $($_.Exception.Message)"
    }
}

# 2. Convert PowerPoint (.pptx) files
$pptxFiles = Get-ChildItem -Path $docsDir -Filter "*.pptx"
Write-Host "`nFound $($pptxFiles.Count) PowerPoint presentations to convert."

if ($pptxFiles.Count -gt 0) {
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        # PowerPoint doesn't support $ppt.Visible = $false directly, use WithWindow = msoFalse (0) in Open
        
        foreach ($file in $pptxFiles) {
            $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
            Write-Host "Converting PowerPoint: $($file.Name) -> $([System.IO.Path]::GetFileName($pdfPath))..."
            
            try {
                # 32 = ppSaveAsPDF, msoFalse = 0
                $presentation = $ppt.Presentations.Open($file.FullName, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
                $presentation.SaveAs($pdfPath, 32)
                $presentation.Close()
                Write-Host "  [OK] Saved: $([System.IO.Path]::GetFileName($pdfPath))"
            } catch {
                Write-Host "  [ERROR] Failed to convert $($file.Name): $($_.Exception.Message)"
            }
        }
        $ppt.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    } catch {
        Write-Host "[ERROR] PowerPoint COM instantiation failed: $($_.Exception.Message)"
    }
}

# 3. Convert Excel (.xlsx) files
$xlsxFiles = Get-ChildItem -Path $docsDir -Filter "*.xlsx"
Write-Host "`nFound $($xlsxFiles.Count) Excel workbooks to convert."

if ($xlsxFiles.Count -gt 0) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        foreach ($file in $xlsxFiles) {
            $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
            Write-Host "Converting Excel: $($file.Name) -> $([System.IO.Path]::GetFileName($pdfPath))..."
            
            try {
                $wb = $excel.Workbooks.Open($file.FullName, [System.Type]::Missing, $true)
                # 0 = xlTypePDF
                $wb.ExportAsFixedFormat(0, $pdfPath)
                $wb.Close($false)
                Write-Host "  [OK] Saved: $([System.IO.Path]::GetFileName($pdfPath))"
            } catch {
                Write-Host "  [ERROR] Failed to convert $($file.Name): $($_.Exception.Message)"
            }
        }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Host "[ERROR] Excel COM instantiation failed: $($_.Exception.Message)"
    }
}

Write-Host "`n=========================================================="
Write-Host "PDF CONVERSION COMPLETE. VERIFYING GENERATED FILES:"
Write-Host "=========================================================="
Get-ChildItem -Path $docsDir -Filter "*.pdf" | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize
