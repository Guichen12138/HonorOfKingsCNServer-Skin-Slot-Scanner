目前是靠官网的英雄编号来确认的，皮肤格子也只能靠俩皮肤中间夹的空格来确认，目前只能扫到已有的皮肤。英雄及皮肤的名称与编号的确认都是靠手输的，目前没有找到有什么地方是可以把英雄编号、皮肤编号与他们对应的名称对应上的地方
下载下来后，python运行hero_scan扫描新的已上线的格子、英雄，再运行skin_scan对这些进行信息的填充。
目前扫出来的疑似老格子产品，四美废案、五虎上将废案（机甲版本）、狼狗（成吉思汗版）赛季皮，另外赵云在龙限和最新的西行封妖记中间还有两个空格子，云缨的破晓联动比七夕限定的格子还早

2026-09-25 起 hero_scan 会在扫描结束后自动从官网英雄列表（pvp.qq.com 的 herolist.json）填充空缺的英雄名字，已有名字一律不覆盖；官网列表里没有的特殊编号仍需手填。

2026-10-31 起新增 C# 高并发扫描引擎（`hero_scan_cs/`）：同目录下存在编译好的 `hero_scan_cs.exe` 时，`hero_scan.py` 会自动调用它进行扫描（全量 999 英雄约 2~3 秒），不存在或启动失败时自动回退到原有的 Python 多线程扫描。
- 自行编译 C# 引擎（需 .NET 9 SDK；默认开启 NativeAOT，需安装 Visual Studio C++ 生成工具）：
  `dotnet publish hero_scan_cs/hero_scan_cs.csproj -c Release -r win-x64 -o .`
- 若未安装 C++ 生成工具，可关闭 AOT 编译为普通单文件：
  `dotnet publish hero_scan_cs/hero_scan_cs.csproj -c Release -r win-x64 -p:PublishAot=false -p:PublishSingleFile=true -p:SelfContained=true -o .`

2026-11-21 起支持皮肤槽位动态自适应续探与候选名参考：
- 当某个英雄在临近默认上限（`00~19` 槽中的 `17~19` 槽）扫出有效资源时，扫描器会自动向后顺延探测至连续 4 个空槽为止（最高到 `99` 槽），并自动在 Excel 表头补齐对应列，中间夹的空格子仍原样保留。
- `skin_scan.py` 在打开未命名皮肤原画时，会自动对比官网 `herolist.json` 与表中该英雄已填皮肤名的差集，在终端打印候选皮肤名供手填参考（不会自动覆盖或盲填表格）。
