# Cài MyMath add-in cho Word theo cách "thư mục add-in tin cậy" (Trusted Add-in Catalog) của Office.
# Chỉ cần thêm add-in MỘT LẦN trong Word; sau đó mỗi lần mở Word, tab MyMath tự có sẵn.
# Cách dùng: bấm đúp "cai-addin-word.cmd" (cùng thư mục), hoặc:
#   powershell -ExecutionPolicy Bypass -File install-addin.ps1 [-Manifest duong\dan\manifest.xml] [-ClearCache]
param(
  [string]$Manifest = (Join-Path $PSScriptRoot 'manifest.xml'),
  [switch]$ClearCache
)
$ErrorActionPreference = 'Stop'
$ShareName   = 'MyMathAddin'
$CatalogGuid = '{6c1f4b2a-8d3e-4f5a-9b7c-2e1d0a3f5b84}'

function Invoke-Elevated([string]$Script) {
  $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
  $p = Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -PassThru -WindowStyle Hidden `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $enc"
  return $p.ExitCode
}

try {
  $Manifest = (Resolve-Path -LiteralPath $Manifest).Path
  [xml]$x = Get-Content -LiteralPath $Manifest -Encoding UTF8
  $id  = $x.OfficeApp.Id
  $ver = $x.OfficeApp.Version
  $src = $x.OfficeApp.DefaultSettings.SourceLocation.DefaultValue

  # 0) Word phải được tắt hẳn
  while (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host 'Word đang mở. Hãy lưu tài liệu, đóng HẾT cửa sổ Word rồi nhấn Enter...' -ForegroundColor Yellow
    [void](Read-Host)
  }

  # 1) Thư mục chứa manifest của MyMath
  $dir = Join-Path $env:LOCALAPPDATA 'MyMath\WordAddin'
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  Copy-Item -LiteralPath $Manifest -Destination (Join-Path $dir 'MyMath.xml') -Force

  # 2) Chia sẻ thư mục đó (Office chỉ nhận danh mục add-in dạng \\tên-máy\thư-mục)
  $unc = "\\$env:COMPUTERNAME\$ShareName"
  if (-not (Test-Path -LiteralPath $unc)) {
    Write-Host 'Windows sẽ hỏi quyền quản trị MỘT lần để chia sẻ thư mục add-in (chỉ đọc, chỉ trên máy này). Bấm Yes.' -ForegroundColor Yellow
    $user = "$env:USERDOMAIN\$env:USERNAME"
    $script = @"
`$ErrorActionPreference = 'Stop'
try {
  New-SmbShare -Name '$ShareName' -Path '$dir' -ReadAccess '$user' | Out-Null
} catch {
  `$all = (New-Object System.Security.Principal.SecurityIdentifier 'S-1-1-0').Translate([System.Security.Principal.NTAccount]).Value
  New-SmbShare -Name '$ShareName' -Path '$dir' -ReadAccess `$all | Out-Null
}
"@
    [void](Invoke-Elevated $script)
    Start-Sleep -Seconds 2
    if (-not (Test-Path -LiteralPath $unc)) { throw "Chưa tạo được thư mục chia sẻ $unc (có thể bạn đã bấm No ở hộp thoại quyền quản trị)." }
  }

  # 3) Khai báo thư mục chia sẻ là "danh mục add-in tin cậy" của Office (hiện trong menu)
  $key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"
  New-Item -Path $key -Force | Out-Null
  New-ItemProperty -Path $key -Name 'Id'    -Value $CatalogGuid -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $key -Name 'Url'   -Value $unc         -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $key -Name 'Flags' -Value 1            -PropertyType DWord  -Force | Out-Null

  # 4) Đăng ký thêm kiểu "Developer" (cách cũ, chắc chắn hiện ngay trong Home → Add-ins → Developer Add-ins)
  $dev = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
  if (-not (Test-Path $dev)) { New-Item -Path $dev -Force | Out-Null }
  New-ItemProperty -Path $dev -Name $id -Value $Manifest -PropertyType String -Force | Out-Null

  # 5) (Tuỳ chọn) xoá bộ đệm add-in của Office
  if ($ClearCache) {
    $wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
    if (Test-Path $wef) {
      Get-ChildItem -LiteralPath $wef -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
      Write-Host 'Đã xoá bộ đệm add-in của Office.'
    }
  }

  Write-Host ''
  Write-Host "Đã cài MyMath add-in $ver." -ForegroundColor Green
  Write-Host "  Danh mục add-in: $unc"
  Write-Host "  Trang web      : $src"
  Write-Host ''
  Write-Host 'Mở Word:' -ForegroundColor Cyan
  Write-Host '  - Dùng ngay: Home → Add-ins → mục Developer Add-ins → MyMath (như trước).'
  Write-Host '  - Để Word nhớ luôn (làm MỘT lần): Home → Add-ins → More Add-ins → thẻ SHARED FOLDER → MyMath → Add.'
} catch {
  Write-Host "Lỗi: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}
