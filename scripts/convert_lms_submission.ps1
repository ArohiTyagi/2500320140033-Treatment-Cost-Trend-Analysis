
# Convert new LMS submission docs to PDF and rename all docs to submission format
# Roll No: 2500320140033 | Arohi Tyagi

param()

$ErrorActionPreference = "Continue"
$DOCS = "C:\Users\Arohi Tyagi\OneDrive\Desktop\033TrendCostAnalysis\docs"
$ROLL = "2500320140033"
$NAME = "ArohiTyagi"

Write-Host "============================================"
Write-Host "CONVERTING LMS SUBMISSION DOCS TO PDF"
Write-Host "============================================"

# --- Word COM ---
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $newDocs = @(
        "2500320140033_ArohiTyagi_Source_Code.docx",
        "2500320140033_ArohiTyagi_10_Data_Validation_Report.docx",
        "2500320140033_ArohiTyagi_GitHub_Streamlit_Details.docx"
    )

    foreach ($fname in $newDocs) {
        $docPath = Join-Path $DOCS $fname
        if (Test-Path $docPath) {
            $pdfName = [System.IO.Path]::ChangeExtension($fname, ".pdf")
            $pdfPath = Join-Path $DOCS $pdfName
            Write-Host "Converting: $fname -> $pdfName ..."
            try {
                $doc = $word.Documents.Open($docPath, $false, $true)
                $doc.ExportAsFixedFormat($pdfPath, 17) # 17 = wdExportFormatPDF
                $doc.Close($false)
                Write-Host "  [OK] $pdfName"
            } catch {
                Write-Host "  [ERROR] $_"
            }
        } else {
            Write-Host "  [SKIP] Not found: $fname"
        }
    }

    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
    Write-Host "[ERROR] Word COM: $_"
}

Write-Host ""
Write-Host "============================================"
Write-Host "RENAMING STEP DOCS TO SUBMISSION FORMAT"
Write-Host "============================================"

# Map: existing filename -> new submission filename
$renames = @{
    "01_Project_Understanding.pdf"                      = "${ROLL}_${NAME}_01_Project_Understanding.pdf"
    "02_HMW_Statements.pdf"                             = "${ROLL}_${NAME}_02_HMW_Statements.pdf"
    "03_Needs_Analysis.pdf"                             = "${ROLL}_${NAME}_03_Needs_Analysis.pdf"
    "04_Traceability_Matrix.xlsx"                       = "${ROLL}_${NAME}_04_Traceability_Matrix.xlsx"
    "05_BRD.pdf"                                        = "${ROLL}_${NAME}_05_BRD.pdf"
    "07_TRD.pdf"                                        = "${ROLL}_${NAME}_07_TRD.pdf"
    "08_Database_Design.xlsx"                           = "${ROLL}_${NAME}_08_Database_Schema.xlsx"
    "09_ER_Diagram.pdf"                                 = "${ROLL}_${NAME}_09_ER_Diagram.pdf"
    "11_User_Stories.pdf"                               = "${ROLL}_${NAME}_11_User_Stories.pdf"
    "12_UI_Framework.pdf"                               = "${ROLL}_${NAME}_12_UI_Framework.pdf"
    "13_Test_Cases.pdf"                                 = "${ROLL}_${NAME}_13_Test_Cases.pdf"
    "033_Project_Document.pdf"                          = "${ROLL}_${NAME}_Project_Report.pdf"
    "033_Presentation.pdf"                              = "${ROLL}_${NAME}_Presentation.pdf"
}

$copied = 0
foreach ($src in $renames.Keys) {
    $srcPath = Join-Path $DOCS $src
    $dstPath = Join-Path $DOCS $renames[$src]
    if (Test-Path $srcPath) {
        Copy-Item $srcPath $dstPath -Force
        Write-Host "  [OK] $src  ->  $($renames[$src])"
        $copied++
    } else {
        Write-Host "  [SKIP] Not found: $src"
    }
}

Write-Host ""
Write-Host "============================================"
Write-Host "VERIFYING ALL SUBMISSION FILES"
Write-Host "============================================"

$submissionFiles = Get-ChildItem $DOCS -Filter "${ROLL}_${NAME}_*" | Sort-Object Name
Write-Host ""
$submissionFiles | Select-Object Name, @{N='Size(KB)';E={[math]::Round($_.Length/1KB,1)}} | Format-Table -AutoSize

$total = ($submissionFiles | Measure-Object Length -Sum).Sum
Write-Host "Total files: $($submissionFiles.Count)"
Write-Host "Total size:  $([math]::Round($total/1MB,2)) MB"
Write-Host "Done."
