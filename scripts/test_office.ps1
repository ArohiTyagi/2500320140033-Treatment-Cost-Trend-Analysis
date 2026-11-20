try {
    $word = New-Object -ComObject Word.Application
    Write-Host "Word Version:" $word.Version
    $word.Quit()
} catch {
    Write-Host "Word COM Error:" $_.Exception.Message
}

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    Write-Host "PowerPoint Version:" $ppt.Version
    $ppt.Quit()
} catch {
    Write-Host "PowerPoint COM Error:" $_.Exception.Message
}
