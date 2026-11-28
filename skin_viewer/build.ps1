param(
    [string]$Version = "",
    [string]$ExcelSource = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ExcelSource)) {
    $ExcelSource = Join-Path (Split-Path -Parent $Root) "hero_skin_scan.xlsx"
}

$UninstProj = Join-Path $Root "Uninstaller/Uninstaller.csproj"
$ViewerProj = Join-Path $Root "SkinViewer/SkinViewer.csproj"

# 未显式传入 -Version 时，以 SkinViewer.csproj 中的 <Version> 为唯一版本来源
if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = ([xml](Get-Content -Raw -Encoding UTF8 $ViewerProj)).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($Version)) { throw "无法从 SkinViewer.csproj 读取 <Version>" }
}
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw "版本号格式应为 X.Y.Z，当前为: $Version" }
$AssetsDir  = Join-Path $Root "SkinViewer/Assets"
$DistDir    = Join-Path $Root "dist"
$TmpOut     = Join-Path $Root "_tmp_build"

Write-Host "=== [1/4] 准备构建目录 (v$Version) ==="
Remove-Item -Recurse -Force $AssetsDir, $TmpOut, $DistDir -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $AssetsDir, $TmpOut, $DistDir | Out-Null

if (-not (Test-Path $ExcelSource)) {
    throw "找不到初始 Excel 文件: $ExcelSource"
}
Copy-Item $ExcelSource (Join-Path $AssetsDir "hero_skin_scan.xlsx") -Force

Write-Host "=== [2/4] 编译 NativeAOT 卸载程序 Uninstall.exe ==="
$UninstOut = Join-Path $TmpOut "uninst"
dotnet publish $UninstProj -c Release -r win-x64 -o $UninstOut
if ($LASTEXITCODE -ne 0) { throw "Uninstaller 编译失败" }
$BuiltUninst = Join-Path $AssetsDir "Uninstall.exe"
Copy-Item (Join-Path $UninstOut "Uninstall.exe") $BuiltUninst -Force

Write-Host "=== [3/4] 编译独立单文件查看器与安装向导 ==="
$ViewerOut = Join-Path $TmpOut "viewer"
dotnet publish $ViewerProj -c Release -r win-x64 -p:Version=$Version -o $ViewerOut
if ($LASTEXITCODE -ne 0) { throw "SkinViewer 编译失败" }

$BuiltExe = Join-Path $ViewerOut "王者荣耀皮肤格子查看器.exe"
$SetupExe = Join-Path $DistDir "SkinViewer_Setup_v$Version.exe"
$UpdateZip = Join-Path $DistDir "SkinViewer_v$Version.zip"

# Setup.exe 与主程序共用同一个独立单文件包（按文件名/安装态自动进入安装向导）
Copy-Item $BuiltExe $SetupExe -Force

# 生成供已安装用户在线自更新用的压缩包（同时包含主程序与卸载器，上传至 GitHub Releases 即可被旧版查看器自动检测升级）
Compress-Archive -Path $BuiltExe, $BuiltUninst -DestinationPath $UpdateZip -CompressionLevel Optimal -Force

Write-Host "=== [4/4] 清理所有编译中间文件 (bin / obj / Assets / _tmp_build) ==="
Remove-Item -Recurse -Force `
    $AssetsDir, `
    $TmpOut, `
    (Join-Path $Root "Uninstaller/bin"), `
    (Join-Path $Root "Uninstaller/obj"), `
    (Join-Path $Root "SkinViewer/bin"), `
    (Join-Path $Root "SkinViewer/obj") `
    -ErrorAction SilentlyContinue

Write-Host "`n构建完成！产物位于: $DistDir"
Get-ChildItem $DistDir | Select-Object Name, @{Name="SizeMB";Expression={[math]::Round($_.Length / 1MB, 2)}}, LastWriteTime | Format-Table -AutoSize
