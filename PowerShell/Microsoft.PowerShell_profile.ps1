Set-Location $env:USERPROFILE
$Shell.WindowTitle="Bananas"
$Shell = $Host.UI.RawUI
$size = $Shell.WindowSize
$size.width=70
$size.height=25
$Shell.WindowSize = $size
$size = $Shell.BufferSize
$size.width=70
$size.height=5000
$Shell.BufferSize = $size
