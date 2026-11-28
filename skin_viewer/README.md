# 王者荣耀皮肤格子查看器（客户端 & 图形安装器）

本目录是面向普通用户的**独立查看与同步客户端及图形安装向导**源码工程，无需安装 Python 或配置任何运行环境即可在 Windows x64 系统上开箱使用。

## 功能概览

1. **图形化安装向导（`SkinViewer_Setup_vX.Y.Z.exe`）**
   - 支持自由选择安装目录（默认优先非系统盘根目录下的 `王者荣耀皮肤格子查看器` 文件夹，避免系统盘权限限制）；
   - 一键释放查看器主程序 **`王者荣耀皮肤格子查看器.exe`**、初始数据表 **`hero_skin_scan.xlsx`** 与原生卸载程序 **`Uninstall.exe`**；
   - 可选创建桌面快捷方式、开始菜单快捷方式，并注册到 Windows「设置 → 应用 → 已安装的应用」标准卸载列表。
2. **数据同步与差异对比（`王者荣耀皮肤格子查看器.exe`）**
   - 自动通过多条国内加速 CDN / 镜像 / GitHub 官方源实时拉取仓库最新的 `hero_skin_scan.xlsx`；
   - 精确对比本地与线上表格，直观列出 `[新增皮肤]`、`[补充命名]`、`[名称修正]`、`[原画更新]`、`[英雄命名]` 明细；
   - 支持一键原子替换更新本地表格，内置 Excel / WPS 文件占用检测保护。
3. **客户端在线自更新（带完整性校验）**
   - 检查更新时同步检测仓库 **GitHub Releases** 中的 `SkinViewer_vX.Y.Z.zip` 更新包（按附件文件名识别，与 Release 标签名无关）；
   - 发现新版查看器时自动弹出升级横幅，支持一键在线下载、原地替换并自动重启；
   - 下载后会核对 GitHub 为每个附件提供的官方 SHA256 校验值，镜像线路内容不一致会被丢弃并改走官方地址；拿不到官方校验值时只从 GitHub 官方地址下载，不使用第三方镜像。
4. **干净卸载（`Uninstall.exe`）**
   - 双击安装目录内的 `Uninstall.exe` 或在系统「已安装的应用」中点击卸载，即可自动关闭运行中的查看器并移除程序文件、本地表格、快捷方式与注册表项（不会误删用户自行放入目录的其他无关文件）。

## 版本号约定

- **查看器版本号的唯一来源**是 `SkinViewer/SkinViewer.csproj` 中的 `<Version>`，`build.ps1` 默认自动读取它。
- 查看器 Release 使用 **`viewer-vX.Y.Z`** 标签；扫描工具（`hero_scan.py` 等）继续使用 `v2.x.x` 标签，两者互不影响。
- 首个查看器 `1.0.0` 随仓库 `v2.3.0` 一同发布，之后的查看器版本都单独发布。

## 从源码编译与打包

需要安装 **.NET 9 SDK**（`Uninstall.exe` 默认启用 NativeAOT 编译，需同时安装 Visual Studio C++ 生成工具）。在 PowerShell 中进入 `skin_viewer` 目录执行：

```powershell
# 按 SkinViewer.csproj 中的 <Version> 打包
powershell -ExecutionPolicy Bypass -File ./build.ps1

# 临时指定版本号打包（一般不需要，发版请直接改 csproj）
powershell -ExecutionPolicy Bypass -File ./build.ps1 -Version 1.1.0
```

编译完成后会自动清理所有 `bin` / `obj` 中间目录，并在 `dist/` 下生成两个发布包：
- **`SkinViewer_Setup_v<版本号>.exe`**：供新用户下载安装的独立安装包（纯英文文件名避免 GitHub Releases 附件名转义，安装后程序与快捷方式显示中文名 `王者荣耀皮肤格子查看器`）；
- **`SkinViewer_v<版本号>.zip`**：在线自更新包（内含主程序与 `Uninstall.exe`）。

## 发布新版查看器步骤

1. 修改 `SkinViewer/SkinViewer.csproj` 中的 `<Version>`（例如 `1.1.0`），在 `CHANGELOG.md` 记一条 `viewer-v1.1.0`，提交并推送。
2. 运行 `build.ps1` 生成 `dist/` 下的两个文件。
3. 在 GitHub 新建 Release，标签填 **`viewer-v1.1.0`**，把 `SkinViewer_Setup_v1.1.0.exe` 和 `SkinViewer_v1.1.0.zip` 都作为附件上传后发布。
4. 已安装旧版的用户打开查看器即会看到升级横幅；GitHub 会自动为附件生成 SHA256 校验值，无需手动填写。
