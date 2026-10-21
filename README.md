# ChromeAutoRun

网页自动化工具 — 基于 Python + Selenium 的 FineReport 报表自动查询、导出、批量填充工具。

## 环境要求

- Python 3.8+
- Google Chrome 浏览器
- Windows 源码运行：chromedriver.exe（与 Chrome 版本匹配，放在项目根目录）
- Linux DEB：Chrome/Chromium 浏览器（ChromeDriver 随 DEB 预装）

```bash
pip install selenium openpyxl
```

## Windows 启动与打包菜单

在项目根目录双击 `启动菜单.bat`，可选择：

1. 启动开发版：运行 `main.py`，优先使用本机的 HZPBC Python 环境；若该环境不存在，则使用 PATH 中的 `python`。
2. 打包 Windows EXE：调用 `build_quick.bat`，按仓库规范使用 `%USERPROFILE%\miniconda3\envs\HZPBC\python.exe`。产物在 `dist/ChromeAutoRun.exe`，并复制到 `dist_hzpbc/ChromeAutoRun.exe` 和带时间戳的归档文件。
3. 打包 Linux DEB：通过 WSL Ubuntu 运行 `packaging/linux/build-deb.sh`，产物在 `dist/linux/`。

也可以直接运行 `build_quick.bat` 打包 Windows EXE。Windows 打包需要安装 HZPBC 环境、Selenium、openpyxl 和 PyInstaller；Linux 打包需要已安装 WSL Ubuntu，并可从 WSL 访问构建依赖的下载源。

## 麒麟 V10 / Linux DEB

仓库提供标准 `.deb` 包，内置 CPython、Selenium、openpyxl 和 ChromeDriver，不依赖目标机预装 Python，也不需要应用启动时联网下载驱动。Chrome/Chromium 仍需在目标机单独安装。构建机需要 Debian 系 Linux、`dpkg-deb`、`curl` 或 `wget`，以及网络访问 Python 运行时、PyPI 和 Chrome for Testing；应在与目标机相同的 CPU 架构上构建（`amd64` 或 `arm64`）。

```sh
sh packaging/linux/build-deb.sh
sudo apt install ./dist/linux/chrome-autorun_1.0.2_*.deb
chrome-autorun
```

如果目标机是 ARM64，构建产物文件名会带 `_arm64.deb`。安装后可从应用菜单打开，也可在终端运行 `chrome-autorun`。首次启动会把空白配置复制到 `~/.local/share/chrome-autorun/`（若设置了 `XDG_DATA_HOME` 则放在其下）；在菜单选择“11. 打开 config.xlsx”填写登录地址和账号。

DEB 默认内置 ChromeDriver `151.0.7922.108`（与当前目标麒麟机器的 Chrome 版本匹配），启动时优先使用该驱动，适合目标机器无法访问公网的场景。构建时会把驱动版本写入 `/usr/share/doc/chrome-autorun/build-info.txt`。如果目标浏览器版本有变化，可用 `CHROMEAUTORUN_CHROMEDRIVER_VERSION=浏览器对应的驱动版本 sh packaging/linux/build-deb.sh` 构建匹配版本。也可设置 `SE_CHROMEDRIVER=/驱动路径` 覆盖包内驱动。如浏览器未在标准路径发现，可设置 `CHROMEAUTORUN_CHROME_BINARY=/浏览器路径`。Chrome/Chromium 需能在当前桌面会话中启动。只有源码运行且未提供本地驱动时，才会回退到 Selenium Manager（需要联网）。

这份 DEB 包含自动化程序、Python 运行时和 Linux ChromeDriver，不包含 Chrome 浏览器，也不会打包仓库里的 Windows `chromedriver.exe`。麒麟 V10 的桌面版本、CPU 架构、浏览器构建和安全策略可能不同，需按目标机架构及浏览器版本构建对应的包。

## 快速开始

```bash
python main.py
```

首次运行会提示 `config.xlsx` 不存在，自动创建模板。填写 `login` sheet 中的 4 个值即可使用：

| key | value |
|-----|-------|
| url | 登录页地址 |
| credit_code | 统一社会信用代码 |
| username | 登录用户名 |
| password | 登录密码 |

## 菜单功能

| 菜单 | 功能 | 说明 |
|------|------|------|
| 1 | 启动浏览器 | 启动 Selenium Chrome |
| 2 | 导航并登录 | 读取 `login` 配置自动登录 |
| 3 | 地区汇总报表查询导出（county） | 读取 `county` sheet，字段填充方式由 `field.xxx.method` 控制 |
| 4 | 单家机构报表查询导出（org） | 读取 `org` sheet，字段填充方式由 `field.xxx.method` 控制 |
| 5 | 注册当地法人机构报表查询导出（county2） | 读取 `county2` sheet，字段填充方式由 `field.xxx.method` 控制 |
| 6 | 报表组批量修改 | 跳过向导，直接在当前报表组条件页填字段 |
| 7 | 提取当前页面 HTML | 含 iframe 递归，逐个显示 HTML 文件的绝对路径 |
| 8 | 提取所有标签页 HTML | 多标签页批量提取，逐个显示 HTML 文件的绝对路径 |
| 9 | 环境自检 | 检查依赖、配置完整性 |
| 10 | 重置配置模板 | 重置 `county/org/county2` 默认模板 |
| 11 | 打开 config.xlsx | 用默认表格软件打开配置 |
| 12 | 浏览器冒烟测试 | 可离线点击内置测试按钮，或只读检查百度/指定内网网页的加载状态与可交互元素数量 |

菜单 12 的网页检查只负责打开网址并读取标题、加载状态和控件数量，不会点击网站控件或提交表单。默认地址为 `https://www.baidu.com`，也可以输入内网的只读页面地址；本地按钮模式不访问网络。

## 配置说明

所有配置集中在 `config.xlsx`，主要 sheet：

| Sheet | 用途 |
|-------|------|
| `login` | 登录信息（url / credit_code / username / password） |
| `county` | 分县区查询配置（报表 title + 字段 mode/values/method；机构类型可配置 validate_options） |
| `org` | 分机构查询配置（报表 title + 字段 mode/values/method；机构名称可配置 validate_options） |
| `county2` | 注册当地法人机构查询配置（报表 title + 字段 mode/values/method；机构类型可配置 validate_options） |
| `batch_fill` | 批量填充全局开关（max_items / sync / save / generate 等） |
| `batch_fill_tables` | 批量填充任务表（task_id / report_group / report_name） |
| `batch_fill_rules` | 批量填充字段规则（condition_index / field_label / field_value） |
| `fr_widget_map` | FineReport widget 名称映射（精确定位输入框） |

详细配置写法见 `使用说明与配置说明.md`。

`county/org/county2` 会优先按 `task.report_title_keyword` 切换外层报表标签页；如果同名候选多个且未配置 `task.report_title_index`，程序会停止并打印候选，避免导错报表。

## 运行输出

页面 HTML 保存在 `output/` 下。菜单 7/8 会显示每个 HTML 文件的绝对路径和保存目录。DEB 安装版默认把 `output/` 放在 `~/.local/share/chrome-autorun/output/`（或 `$XDG_DATA_HOME/chrome-autorun/output/`）；源码和 Windows 版放在程序目录下的 `output/`。

Chrome 下载默认保存在用户目录下的 `Downloads/`。也可以在 `config.json` 中配置 `download_dir`。启动浏览器时会显示下载目录，程序导出文件后会逐个显示文件的绝对路径。

每次执行后在 `output/` 下生成：

```
output/
├── batch_fill_<时间>_<id>/
│   ├── run_log.jsonl        # 步骤级日志
│   ├── screenshots/          # 每步截图
│   └── snapshots/            # 条件快照 JSON
├── county_<时间>_<id>/      # 分县区查询结果
└── *.html                    # 页面提取结果
```

## 项目结构

```
├── main.py              # 主程序（单文件，约 2300 行）
├── config.xlsx          # 配置文件
├── config.json          # 运行时配置（由 config.xlsx 自动同步）
├── chromedriver.exe     # Chrome 驱动
├── output/              # 运行输出目录
└── README.md
```
