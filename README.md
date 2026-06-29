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
4. 按 `Alt + F8` 运行 `AAA启动控制台` 宏。

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

VBA 源文件要求：

- `.bas` 文件统一使用 UTF-8 编码。
- `.bas` 文件统一使用 CRLF 行尾。
- 规则记录在 `AGENTS.md`。

## 文件目录

```txt
.
|- 上层产品净值数据库.xlsm
|- README.md
|- AGENTS.md
|- .gitignore
|- sync-vba.ps1
|- assets
|  |- chart_templates
|  |  |- 净值图表_红.crtx
|  |  |- 净值图表_蓝.crtx
|  |  |- 收益率图表_红.crtx
|  |  `- 收益率图表_蓝.crtx
|  |- images
|  |  |- background.png
|  |  |- logo.png
|  |  |- logo_white.png
|  |  `- title.png
|  `- 产品一页通-交鑫致远-模板.pptx
`- scripts
   |- py
   |  `- fetch_reits_total_return.py
   `- vba
      |- chart
      |  |- 01_auto_input.bas
      |  |- 02_output_data.bas
      |  |- 03_output_chart.bas
      |  `- 04_output_image.bas
      |- data
      |  |- 01_auto_input.bas
      |  |- 02_calculate_open_date.bas
      |  |- 03_output_date.bas
      |  `- 04_output_report.bas
      |- optional_panel
      |  |- 00_operation_panel.bas
      |  |- 00_operation_panel_button.cls
      |  `- 00_operation_panel_form.frm
      |- product_one_page
      |  |- 00_check_import_data.bas
      |  |- 01_output_data.bas
      |  |- 02_output_chart.bas
      |  `- 03_output_ppt.bas
      |- tool
      |  |- t_01_clean_data.bas
      |  |- t_02_del_data.bas
      |  |- t_03_next_open_date_by_interval.bas
      |  |- t_04_check_nav_data.bas
      |  |- t_05_query_181_nav_stats.bas
      |  `- t_06_delete_empty_rows.bas
      `- weekly_recommendation
         |- 01_update_weekly_recommendation_dependencies.bas
         `- 02_generate_weekly_recommendation.bas
```

目录说明：

- `上层产品净值数据库.xlsm`：主工作簿，VBA 宏运行入口。
- `sync-vba.ps1`：将仓库源码导入主工作簿的同步脚本。
- `assets\chart_templates`：Excel 图表模板，`scripts\vba\chart\03_output_chart.bas` 会从这里读取 `.crtx`。
- `assets\images`：报表固定图片资源，包含背景图、logo 和标题图。
- `assets\产品一页通-交鑫致远-模板.pptx`：产品一页通 PPT 模板。
- `scripts\py`：Python 数据抓取脚本。
- `scripts\vba\data`：数据导入、开放日计算、展示报表输出。
- `scripts\vba\chart`：产品净值数据整理、图表生成、图表图片导出。
- `scripts\vba\weekly_recommendation`：周度推荐材料依赖更新和材料生成。
- `scripts\vba\product_one_page`：产品一页通数据、图表和 PPT 输出。
- `scripts\vba\optional_panel`：VBA 操作面板相关模块、类和窗体。
- `scripts\vba\tool`：清理、检查和查询类工具宏。

## 更新须知

### 更新 VBA 程序

1. 使用 IDE 直接编辑 `scripts\vba\*.bas`、`.cls` 或 `.frm` 文件。
2. 确认 `.bas` 文件仍为 UTF-8 + CRLF。
3. 运行 `.\sync-vba.ps1` 导入 VBA 程序到数据库。
4. 打开数据库进行测试。

### 更新数据

1. 基准收益率：工作表`产品分类`。
2. 增加图表：工作表`绘图净值数据`。
3. 推荐材料修改或者新增：工作表`产品要素`。
4. 推荐材料流程图关系：工作表`主要底层资产`。
5. 推荐材料具体数据计算关系：工作表`底层资产对应关系`。
6. 缺失开放日：VBA 工具 `Tool04_CheckNavData`。

### 更新资源文件

1. 图表模板放在 `assets\chart_templates`。
2. 报表固定图片放在 `assets\images`。
3. 产品一页通 PPT 模板放在 `assets` 根目录。
