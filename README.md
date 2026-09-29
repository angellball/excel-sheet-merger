# 合并同名 Excel 工作表

该脚本会递归扫描指定文件夹及其子文件夹内的 `.xls`、`.xlsx`、`.xlsm` 和 `.csv` 文件，把名称完全相同的工作表纵向追加到同一张表中，并输出为一个新的 `.xlsx` 文件。CSV 文件按文件名（不含扩展名）作为工作表名。

默认规则：

- 按相对路径进行自然升序处理，例如 `0-9`、`9-13`、`13-16`；
- 可以选择标题行数，默认只保留第一张同名表的 1 行标题；
- `.xlsx`/`.xlsm` 保留常见单元格格式、超链接、批注、行高、列宽和合并单元格；
- 合并时跳过空行，并自动调整 `.xlsx`/`.xlsm` 公式中的相对引用；
- 忽略 Excel 临时文件（文件名以 `~$` 开头）、本次输出文件及其合并临时文件；
- 空白工作表不写入结果；
- 始终严格按照排序结果逐一合并，可选并发预读；
- 先写临时文件，成功后再替换结果，避免失败时损坏原输出文件。

保存时会确保 ZIP 归档和临时文件句柄关闭。Windows 短暂占用文件时会重试替换；如果结果文件一直被 Excel/WPS 等程序占用，程序会提示关闭文件，并保留本次完整结果的 `.tmp` 文件及其路径。解除占用后，可将该文件改名为 `.xlsx` 使用。临时文件不会参与后续合并。

## 输入格式

| 格式 | 读取与合并规则 |
| --- | --- |
| `.xlsx` / `.xlsm` | 按原工作表名称合并，保留现有格式与公式处理。输出为 `.xlsx`，不保留 VBA 宏。 |
| `.xls` | 用 `xlrd` 读取原工作表名称、文本、数字、日期、时间、布尔值和 Excel 错误值，保留数字格式、行高、列宽和合并单元格。公式使用文件中已有的计算结果；不保留公式本身、字体、填充、边框、超链接或批注。 |
| `.csv` | 每个文件作为一张表，名称取文件名（不含扩展名）；非法表名字符替换为 `_`，名称最多 31 个字符。优先按 UTF-8（包括 BOM）读取，失败后尝试 GB18030。自动识别逗号、分号、制表符和竖线分隔符，无法识别时使用逗号。 |

CSV 的所有非空字段保持为文本，因此前导零和以 `=` 开头的内容不会被转换。CSV 与 Excel 中名称相同的表也会合并，遵循相同的标题行与空行规则。

`.xls` 中部分中文等地区内置格式没有可读取的格式代码。遇到这类格式或缺失格式定义时，保留原始数值类型，日期/时间使用 `yyyy-mm-dd hh:mm:ss`，其他值使用 `General` 兜底；日期不会变成 Excel 序列号，原有有效格式继续保留。

## 独立程序

`dist/excel_sheet_merger`（Windows 下为 `dist\excel_sheet_merger.exe`）是单文件程序，运行时不需要安装 Python、openpyxl 或 xlrd。更新源码后需要重新构建独立程序。

独立程序必须在目标操作系统上构建。程序包含 `.xls` 读取所需的 `xlrd` 依赖。Linux/macOS 构建命令：

```bash
python3 -m venv .build-venv
source .build-venv/bin/activate
python -m pip install -r requirements-build.txt
python build_standalone.py
```

Windows 构建命令：

```bash
py -m venv .build-venv
.build-venv\Scripts\python -m pip install -r requirements-build.txt
.build-venv\Scripts\python build_standalone.py
```

## 源码运行

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## 使用

直接运行，按照提示输入要扫描的文件夹路径和标题行数：

```bash
python3 merge_workbooks.py
```

程序会递归查找所输入文件夹及其所有子文件夹中的 Excel/CSV 文件。标题行数留空时默认为 `1`，输入 `0` 表示工作表没有标题行。

在终端运行时会显示当前文件、工作表和总体进度。重定向输出或使用 `--no-progress` 时不显示进度条。

合并指定文件夹：

```bash
python3 merge_workbooks.py /path/to/excel_folder
```

指定输出文件：

```bash
python3 merge_workbooks.py /path/to/excel_folder -o /path/to/result.xlsx
```

非交互运行时指定标题行数，例如标题占 2 行：

```bash
python3 merge_workbooks.py /path/to/excel_folder --header-rows 2
```

如果工作表没有标题行：

```bash
python3 merge_workbooks.py /path/to/excel_folder --header-rows 0
```

如果需要保留每个来源工作表的所有标题：

```bash
python3 merge_workbooks.py /path/to/excel_folder --keep-all-headers
```

默认并发预读数为 1。实测本地文件使用 2 个线程反而略慢并占用更多内存；网络盘或慢盘可以自行尝试更高的值：

```bash
python3 merge_workbooks.py /path/to/excel_folder --workers 2
```

## 测试

安装 `requirements.txt` 中的运行依赖后执行：

```bash
python -m unittest discover -s tests -v
```

测试包含有效的旧版 Excel BIFF 工作簿样本，不需要额外安装 `.xls` 写入库。
