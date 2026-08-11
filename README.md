# 产品净值 VBA 报表

本项目用于维护 `上层产品净值数据库.xlsm` 的 VBA 模块、图表模板、报表图片资源和少量外部数据抓取脚本。

## 操作指南

### 在外网机

1. 安装 Python 依赖
2. 运行 `scripts\py\fetch_reits_total_return.py`，导出中证 REITs 全收益数据。
3. 运行 `.\sync-vba.ps1`，将 `scripts\vba` 中的模块同步导入数据库工作簿。
4. 将更新后的数据库工作簿和导出的数据文件转移到内网机。

### 在内网机

1. 下载或放置 `HS-181_多账套净值查询_yyyymmdd.xlsx` 到数据库所在目录。
2. 确认外部数据文件、模板文件和图片资源均位于下方“文件目录”说明的位置。
3. 打开 `上层产品净值数据库.xlsm`。
4. 首次升级到配置驱动版本时，运行 `Report00_MigrateConfiguration` 创建配置表和“是否导出”字段。
5. 按 `Alt + F8` 运行 `AAA启动控制台` 宏。
6. 如果流程生成“开放日待确认”，核对并确认后再次运行一键流程。

## Python 程序运行

Python 脚本位于 `scripts\py`，当前用于抓取并导出中证 REITs 全收益数据：

```powershell
python scripts\py\fetch_reits_total_return.py
```

注意：

- 该脚本需要可访问外网，因为数据来自 `akshare`。
- `--end-date` 格式必须是 `yyyymmdd`。
- 结束日期不能早于脚本内的 REITs 基准日期 `20250717`。
- 如果提示缺少依赖，运行 `pip install pandas akshare openpyxl`。

## VBA 同步

将仓库中的 VBA 源码同步导入 `上层产品净值数据库.xlsm`：

```powershell
.\sync-vba.ps1
```

VBA 源文件与运行环境要求：

- `.bas`、`.cls` 和 `.frm` 文件使用 UTF-8 编码及 CRLF 行尾。
- 同步脚本支持 Windows PowerShell 5.1 和 PowerShell 7。
- Excel VBE 所在 Windows 的 ANSI code page 必须为 CP936；不支持开启“使用 Unicode UTF-8 提供全球语言支持”的环境。
- 导入前会严格检查 UTF-8、CRLF 和 CP936 可表示性；导入后会从 VBE 读回代码并与源码比较（忽略 VBE 自动调整的标识符大小写和行尾空白），其他差异会阻止保存工作簿。
- `.bas` 文件的仓库规则记录在 `AGENTS.md`。

## 文件目录

```txt
.
├── 上层产品净值数据库.xlsm              # 主工作簿，VBA 宏运行入口
├── README.md
├── AGENTS.md
├── .gitignore
├── sync-vba.ps1                         # 将仓库源码导入主工作簿的同步脚本
├── assets
│   ├── chart_templates                  # Excel 图表模板（.crtx）
│   │   ├── 净值图表_红.crtx
│   │   ├── 净值图表_蓝.crtx
│   │   ├── 收益率图表_红.crtx
│   │   └── 收益率图表_蓝.crtx
│   ├── images                           # 报表固定图片资源（背景图/logo/标题图）
│   │   ├── background.png
│   │   ├── logo.png
│   │   ├── logo_white.png
│   │   └── title.png
│   └── 产品一页通-交鑫致远-模板.pptx    # 产品一页通 PPT 模板
└── scripts
    ├── py
    │   └── fetch_reits_total_return.py  # 中证 REITs 全收益数据抓取
    └── vba
        ├── data                         # 配置预检 → 数据导入 → 开放日确认 → 指标计算 → 展示输出
        │   ├── 00_report_configuration.bas
        │   ├── 00_open_date_confirmation.bas
        │   ├── 00_report_pipeline.bas
        │   ├── 01_auto_input.bas
        │   ├── 02_calculate_open_date.bas
        │   ├── 03_output_date.bas
        │   └── 04_output_report.bas
        ├── chart                        # 净值数据整理 → 图表生成 → 图表图片导出
        │   ├── 01_auto_input.bas
        │   ├── 02_output_data.bas
        │   ├── 03_output_chart.bas
        │   └── 04_output_image.bas
        ├── weekly_recommendation        # 周度推荐材料：依赖更新 → 材料生成
        │   ├── 01_update_weekly_recommendation_dependencies.bas
        │   └── 02_generate_weekly_recommendation.bas
        ├── product_one_page             # 产品一页通：数据 → 图表 → PPT 输出
        │   ├── 00_check_import_data.bas
        │   ├── 01_output_data.bas
        │   ├── 02_output_chart.bas
        │   └── 03_output_ppt.bas
        ├── optional_panel               # VBA 操作面板（模块/类/窗体）
        │   ├── 00_operation_panel.bas
        │   ├── 00_operation_panel_button.cls
        │   └── 00_operation_panel_form.frm
        └── tool                         # 工具宏：清理/检查/查询
            ├── t_01_clean_data.bas
            ├── t_02_del_data.bas
            ├── t_03_next_open_date_by_interval.bas
            ├── t_04_check_nav_data.bas
            ├── t_05_query_181_nav_stats.bas
            └── t_06_delete_empty_rows.bas
```

## 更新须知

### 更新 VBA 程序

1. 使用 IDE 直接编辑 `scripts\vba\*.bas`、`.cls` 或 `.frm` 文件。
2. 确认 `.bas` 文件仍为 UTF-8 + CRLF。
3. 运行 `.\sync-vba.ps1` 导入 VBA 程序到数据库。
4. 打开数据库进行测试。

### 更新数据

1. 产品增减、唯一归属、系列和基准收益率：工作表`产品分类`；使用“是否导出”停用产品，不删除历史行。
2. 分类、展示分组、输出字段方案和图表位置：工作表`报表配置`。
3. 图表历史净值：工作表`绘图净值数据`；需要绘制的产品由“报表配置”决定。
4. 推算开放日：先在工作表`开放日待确认`确认，再写入`开放日`台账。
5. 推荐材料修改或者新增：工作表`产品要素`。
6. 推荐材料流程图关系：工作表`主要底层资产`。
7. 推荐材料具体数据计算关系：工作表`底层资产对应关系`。
8. 缺失净值：VBA 工具 `Tool04_CheckNavData`。

完整配置说明见 [`docs/报表配置与操作手册.md`](docs/报表配置与操作手册.md)。

### 更新资源文件

1. 图表模板放在 `assets\chart_templates`。
2. 报表固定图片放在 `assets\images`。
3. 产品一页通 PPT 模板放在 `assets` 根目录。
