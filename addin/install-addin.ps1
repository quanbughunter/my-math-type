# Đăng ký MyMath add-in cho Word trên máy này (chỉ tài khoản hiện tại, không cần quyền quản trị).
# Cách dùng: bấm đúp "cai-addin-word.cmd" (cùng thư mục), hoặc:
#   powershell -ExecutionPolicy Bypass -File install-addin.ps1 [-Manifest duong\dan\manifest.xml]
param([string]$Manifest = (Join-Path $PSScriptRoot 'manifest.xml'))
$ErrorActionPreference = 'Stop'
$Manifest = (Resolve-Path -LiteralPath $Manifest).Path
[xml]$x = Get-Content -LiteralPath $Manifest -Encoding UTF8
$id = $x.OfficeApp.Id
$src = $x.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
New-ItemProperty -Path $key -Name $id -Value $Manifest -PropertyType String -Force | Out-Null
Write-Host ''
Write-Host 'Đã đăng ký MyMath add-in cho Word.' -ForegroundColor Green
Write-Host "  Manifest : $Manifest"
Write-Host "  Trang web: $src"
Write-Host ''
Write-Host 'Đóng hết cửa sổ Word rồi mở lại. Nút "Công thức" nằm ở tab Chèn (Insert) → nhóm MyMath.'
Write-Host 'Nếu chưa thấy: Chèn → Bổ trợ của tôi (My Add-ins) → Bổ trợ nhà phát triển (Developer Add-ins).'
