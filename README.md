目前是靠官网的英雄编号来确认的，皮肤格子也只能靠俩皮肤中间夹的空格来确认，目前只能扫到已有的皮肤。英雄及皮肤的名称与编号的确认都是靠手输的，目前没有找到有什么地方是可以把英雄编号、皮肤编号与他们对应的名称对应上的地方
下载下来后，python运行hero_scan扫描新的已上线的格子、英雄，再运行skin_scan对这些进行信息的填充。
目前扫出来的疑似老格子产品，四美废案、五虎上将废案（机甲版本）、狼狗（成吉思汗版）赛季皮，另外赵云在龙限和最新的西行封妖记中间还有两个空格子，云缨的破晓联动比七夕限定的格子还早

2026-09-25 起 hero_scan 会在扫描结束后自动从官网英雄列表（pvp.qq.com 的 herolist.json）填充空缺的英雄名字，已有名字一律不覆盖；官网列表里没有的特殊编号仍需手填。

## 桌面查看器客户端（免环境一键安装 & 自动同步）

如果只是想查看最新的全英雄皮肤格子 Excel 表格、不想配置 Python 或 Git 环境，可以直接使用仓库提供的 **Windows 独立桌面查看器（`王者荣耀皮肤格子查看器`）**。

### 功能简介
- **开箱即用**：独立单文件图形安装器，无需安装 Python、.NET 运行时或 Git，安装后即自带初始 `hero_skin_scan.xlsx` 表格与主程序。
- **多源实时同步与变更明细**：内置多条国内 CDN / 镜像加速源与防缓存校验，自动拉取仓库最新表格并与本地表格逐格比对，直观列出 **`[新增皮肤]`**、**`[补充命名]`**、**`[名称修正]`**、**`[原画更新]`** 等具体英雄槽位变化。
- **安全一键更新表格**：点击按钮即可原子替换更新本地 Excel 表格；若表格正被 Excel / WPS 打开会自动拦截提醒，防止损表。
- **程序在线自更新 & 干净卸载**：支持自动检测 Releases 中的新版查看器更新包并一键在线升级重启；自带原生 `Uninstall.exe` 卸载程序，支持桌面/开始菜单快捷方式及 Windows「已安装的应用」标准卸载。

### 下载与使用教程
1. **下载安装包**
   前往仓库右侧 **[Releases](https://github.com/Guichen12138/HonorOfKingsCNServer-Skin-Slot-Scanner/releases/latest)** 页面，在底部 **Assets（资产）** 中下载 **`王者荣耀皮肤格子查看器_Setup.exe`**。
   > 若首次双击弹出 Windows SmartScreen「Windows 已保护你的电脑」提示，点击 **「更多信息」→「仍要运行」** 即可。
2. **安装程序**
   打开安装向导后可自由选择安装文件夹（默认优先选择非系统盘根目录下的 `王者荣耀皮肤格子查看器` 目录，避开系统盘权限限制），按需勾选创建桌面快捷方式，点击 **「立即安装」**。
3. **日常检查与同步最新皮肤表**
   - 启动 **「王者荣耀皮肤格子查看器」**，程序会自动连接仓库检查最新数据：
     - 状态显示 **「已是最新」**：说明本地 `hero_skin_scan.xlsx` 已与仓库完全一致，直接点击 **「打开 Excel 表格」** 即可查看；
     - 状态显示 **「发现 N 处更新」**：下方日志框会完整列出本次新增了哪些英雄皮肤槽位或补充了哪些皮肤名字，确认无误后点击绿色 **「更新本地 Excel 表格」** 按钮即可同步到本地（同步前请先关闭正在打开该表格的 Excel / WPS 窗口）。
4. **升级查看器与卸载**
   - **升级查看器**：当后续版本升级查看器功能时，软件顶部会自动出现紫色升级横幅，点击 **「立即升级查看器」** 即可自动在线升级并重启（无需重新下载安装包，也不会丢失本地表格）。
   - **卸载程序**：双击安装目录内的 `Uninstall.exe`，或在系统「设置 → 应用 → 已安装的应用」中找到 **「王者荣耀皮肤格子查看器」** 点击卸载即可彻底清理（不会误删你自行放入该目录的其他文件）。

## 运行与编译（扫描脚本与源码构建）

- **Python 依赖**：`pip install -r requirements.txt`（直接运行 `hero_scan.py` / `skin_scan.py` 即可，无需额外编译）
- **可选 C# 加速引擎（`hero_scan_cs/`）**：同目录存在 `hero_scan_cs.exe` 时，`hero_scan.py` 会自动调用它扫描（全量 999 英雄约 2~3 秒）；不存在或启动失败时自动回退原有的 Python 多线程扫描。
  - 默认 NativeAOT 编译（需 .NET 9 SDK + Visual Studio C++ 生成工具）：
    ```bash
    dotnet publish hero_scan_cs/hero_scan_cs.csproj -c Release -r win-x64 -o .
    ```
  - 未安装 C++ 生成工具时，可关闭 AOT 编译为普通单文件：
    ```bash
    dotnet publish hero_scan_cs/hero_scan_cs.csproj -c Release -r win-x64 -p:PublishAot=false -p:PublishSingleFile=true -p:SelfContained=true -o .
    ```
- **桌面查看器与安装器源码构建（`skin_viewer/`）**：需安装 .NET 9 SDK 与 Visual Studio C++ 生成工具，一键编译生成安装包与在线自更新包（构建完成后自动清理 `bin` / `obj` 中间目录）：
  ```powershell
  powershell -ExecutionPolicy Bypass -File skin_viewer/build.ps1
  ```

## 更新日志

各版本详细功能演进见 [CHANGELOG.md](./CHANGELOG.md) 或仓库右侧 **Releases**。
