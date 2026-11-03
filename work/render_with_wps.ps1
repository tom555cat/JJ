param([string]$InputDoc,[string]$OutputPdf)
$ErrorActionPreference='Stop'
$examWps=$null
$examDoc=$null
try {
    $examWps=New-Object -ComObject KWPS.Application
    $examWps.Visible=$false
    $examWps.DisplayAlerts=0
    $examDoc=$examWps.Documents.Open($InputDoc,$false,$true)
    $examDoc.ExportAsFixedFormat($OutputPdf,17)
    Write-Output ('EXPORTED '+$OutputPdf)
} finally {
    if ($null -ne $examDoc) { $examDoc.Close(0); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($examDoc) | Out-Null }
    if ($null -ne $examWps) { $examWps.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($examWps) | Out-Null }
}
