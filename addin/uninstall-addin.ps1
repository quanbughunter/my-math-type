# Gỡ đăng ký MyMath add-in khỏi Word.
param([string]$Manifest = (Join-Path $PSScriptRoot 'manifest.xml'))
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (Test-Path -LiteralPath $Manifest) {
  [xml]$x = Get-Content -LiteralPath $Manifest -Encoding UTF8
  Remove-ItemProperty -Path $key -Name $x.OfficeApp.Id -ErrorAction SilentlyContinue
}
Write-Host 'Đã gỡ MyMath add-in. Mở lại Word để thấy thay đổi.' -ForegroundColor Yellow
