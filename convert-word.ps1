param([string]$Source, [string]$Target)
$ErrorActionPreference = 'Stop'
$wordApp = $null
$document = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $wordApp.Options.UpdateLinksAtOpen = $false
    $document = $wordApp.Documents.Open($Source, $false, $true, $false)
    $document.ExportAsFixedFormat($Target, 17)
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($wordApp) { $wordApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) }
}
