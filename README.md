目前是靠官网的英雄编号来确认的，皮肤格子也只能靠俩皮肤中间夹的空格来确认，目前只能扫到已有的皮肤。英雄及皮肤的名称与编号的确认都是靠手输的，目前没有找到有什么地方是可以把英雄编号、皮肤编号与他们对应的名称对应上的地方
下载下来后，python运行hero_scan扫描新的已上线的格子、英雄，再运行skin_scan对这些进行信息的填充。
目前扫出来的疑似老格子产品，四美废案、五虎上将废案（机甲版本）、狼狗（成吉思汗版）赛季皮，另外赵云在龙限和最新的西行封妖记中间还有两个空格子，云缨的破晓联动比七夕限定的格子还早

2026-09-25 起 hero_scan 会在扫描结束后自动从官网英雄列表（pvp.qq.com 的 herolist.json）填充空缺的英雄名字，已有名字一律不覆盖；官网列表里没有的特殊编号仍需手填。

## 运行与编译

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

## 更新日志

各版本详细功能演进见 [CHANGELOG.md](./CHANGELOG.md) 或仓库右侧 **Releases**。
