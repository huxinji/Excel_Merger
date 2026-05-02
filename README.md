# ⊞ Excel Merger · Excel 合并工具

> A clean desktop app for merging Excel & CSV files — no coding required.
> 一款简洁的桌面工具，用于合并 Excel 与 CSV 文件，无需编程基础。

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat-square&logo=python)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?style=flat-square&logo=pandas)
![openpyxl](https://img.shields.io/badge/openpyxl-3.x-green?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-lightgrey?style=flat-square)

---

## Installation · 安装

**1. 确认 Python 版本 / Confirm Python version**

```bash
python --version   # requires 3.10+
```

**2. 安装依赖库 / Install dependencies**

```bash
pip install pandas openpyxl
```

| Package | Purpose |
|---|---|
| `pandas` | Data reading, merging, and transformation |
| `openpyxl` | Writing `.xlsx` output files |

> GUI is built with Python's built-in `tkinter` — no extra install needed.
> 界面基于 Python 内置的 `tkinter`，无需额外安装。

**3. 运行 / Run**

```bash
python excel_merger.py
```

---

## Merge Modes · 合并模式

### Mode A — All into One Sheet · 合并成一张表

All files and sheets are stacked into a single worksheet.  
Columns are auto-aligned; missing values are left blank.

所有文件和工作表合并为一张表。  
列自动对齐，缺失列填空。

```
File 1: Name | Score          File 2: Name | Grade
        ─────────────                 ─────────────
        Alice | 90                    Bob  | A
        Bob   | 85

→ Merged:  Name | Score | Grade
           ─────────────────────
           Alice | 90   |
           Bob   | 85   |
           Bob   |      | A
```

---

### Mode B — Separate Sheets · 分开成多个工作表

Each source file or sheet becomes an **individual sheet** in the output workbook.  
Sheet names are auto-deduplicated (e.g. `Sales_2` if `Sales` already exists).

每个源文件或工作表在输出文件中保留为**独立的工作表**。  
重名自动追加编号（如已有 `Sales` 则命名为 `Sales_2`）。

```
File 1 → Sheet: "Jan_Sales"
File 2 → Sheet: "Feb_Sales"
File 3 → Sheet: "Feb_Sales_2"   ← auto-renamed · 自动重命名
```

---

### Mode C — Group by Column · 按列值分组

Specify a column name. All rows across all files are grouped by its **unique values**, each group becoming its own sheet.

指定一个列名，所有文件中的行按该列的**唯一值**分组，每个分组生成独立的工作表。

```
Column: "Region"

Row data:                      Output sheets:
  Alice | East   →  Sheet "East":   Alice, Carol
  Bob   | West   →  Sheet "West":   Bob
  Carol | East   →  Sheet "North":  David
  David | North
```

---

## Merge Direction · 合并方向

### Vertical · 纵向合并

Rows from multiple files are **appended** one after another.  
Use when files share the same structure (e.g. monthly reports).

多个文件的行**依次追加**。  
适用于结构相同的文件（如每月报表）。

```
File A      File B      Result
──────── +  ────────  = ──────────────
A | B       A | C       A | B | C
1 | x       3 | p       1 | x |
2 | y       4 | q       2 | y |
                        3 |   | p
                        4 |   | q
```

### Horizontal · 横向合并

Files are **joined side-by-side** via a shared key column.  
Supports `inner` / `left` / `outer` join types.

通过共同的关联列将文件**横向拼接**。  
支持 `inner`（交集）/ `left`（左表全保留）/ `outer`（并集）三种连接方式。

```
File A          File B         Result (key = "ID")
──────────── +  ────────────  = ──────────────────────
ID | Name       ID | Score      ID | Name  | Score
1  | Alice      1  | 90         1  | Alice | 90
2  | Bob        3  | 78         2  | Bob   |        ← outer: kept
                                3  |       | 78     ← outer: kept
```

---

## Features · 功能概览

- ✅ Add individual files or entire folders (with optional subfolder scan)  
  支持逐个添加文件或整个文件夹（可含子文件夹）
- ✅ CSV auto-encoding detection (UTF-8, GBK, GB18030 …)  
  CSV 自动检测编码
- ✅ Output never overwrites — auto-renamed if file exists  
  输出不覆盖已有文件，自动重命名
- ✅ Background threading — UI stays responsive during processing  
  后台线程处理，界面不卡顿
- ✅ Styled output: frozen header row, auto column widths  
  输出带样式：冻结表头、列宽自适应
- ✅ Real-time log panel with colour-coded messages  
  实时日志面板，彩色状态提示

---

## Requirements · 环境要求

| | Minimum |
|---|---|
| Python | 3.10+ |
| OS | Windows / macOS / Linux |
| pandas | 1.5+ |
| openpyxl | 3.0+ |

---

## License · 许可

MIT License — free to use, modify, and distribute.  
MIT 协议 — 可自由使用、修改与分发。
