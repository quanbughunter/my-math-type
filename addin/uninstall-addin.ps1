# Gỡ MyMath add-in khỏi Word (xoá danh mục tin cậy, thư mục chia sẻ và bản cài kiểu Developer).
param([string]$Manifest = (Join-Path $PSScriptRoot 'manifest.xml'))
$ShareName   = 'MyMathAddin'
$CatalogGuid = '{6c1f4b2a-8d3e-4f5a-9b7c-2e1d0a3f5b84}'
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid" -Recurse -ErrorAction SilentlyContinue
if (Test-Path -LiteralPath $Manifest) {
  [xml]$x = Get-Content -LiteralPath $Manifest -Encoding UTF8
  Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -Name $x.OfficeApp.Id -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath "\\$env:COMPUTERNAME\$ShareName") {
  Write-Host 'Windows sẽ hỏi quyền quản trị để bỏ chia sẻ thư mục add-in. Bấm Yes.' -ForegroundColor Yellow
  $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes("Remove-SmbShare -Name '$ShareName' -Force"))
  try { Start-Process powershell.exe -Verb RunAs -Wait -WindowStyle Hidden -ArgumentList "-NoProfile -EncodedCommand $enc" } catch {}
}
Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'MyMath\WordAddin') -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'Đã gỡ MyMath add-in. Nếu Word vẫn còn hiện MyMath: Home → Add-ins → More Add-ins → My Add-ins → ... → Remove.' -ForegroundColor Yellow
