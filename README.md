# 分摊工具

一个本地 Windows 桌面工具，用于把 Excel/WPS 表格中的费用列按指定基数列占比自动分摊。

## 主要功能

- 选择 `.xlsx` / `.xlsm` 文件。
- 选择工作表。
- 指定第几行是表头。
- 多选“参与占比计算”的列，例如：完工入库材料成本、本期人工费。
- 多选“需要分配”的费用列，例如：本期共耗料、水电费、维修费。
- 指定一个过滤列，并选择该列中不参与计算的值，例如：生产车间 = 销售配货部。
- 输出新 Excel 文件，不修改原文件。
- 自动新增“分摊明细”工作表，方便核对参与行、排除原因、占比和分摊结果。

## 分摊规则

对每个需要分配的费用列：

1. 先汇总该列原始金额，作为待分摊总额。
2. 每行分摊基数 = 所选参与占比列的数值合计。
3. 过滤命中的行不参与分摊。
4. 分摊基数小于等于 0 的行不参与分摊。
5. 参与行金额 = 待分摊总额 × 本行分摊基数 / 参与行分摊基数合计。
6. 金额按 2 位小数四舍五入，尾差自动放到分摊基数最大的参与行。

## 本地开发运行

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements-dev.txt
python main.py
```

## 运行测试

```powershell
python -m pytest
```

## GitHub 自动打包

把整个项目上传到 GitHub 后，GitHub Actions 会在每次 push 或手动运行时自动生成 Windows exe。

普通 push 或手动运行时，GitHub Actions 的 `Artifacts` 仍会以 zip 形式保存构建产物，这是 GitHub Actions 的默认机制。

如果希望用户直接下载 `.exe`，请打一个版本标签发布 Release：

```powershell
git tag v0.1.0
git push origin v0.1.0
```

发布后下载位置：

1. 打开 GitHub 仓库。
2. 进入 `Releases`。
3. 打开对应版本。
4. 在 `Assets` 里直接下载：
   - `FanchanTool-win-x86.exe`：32 位版本，适合 Win7 32 位，也可在 Win7 64 位运行。
   - `FanchanTool-win-x64.exe`：64 位版本，适合数据量较大且系统是 Win7 64 位的电脑。

## Win7 兼容说明

项目使用 Python 3.8 和 Tkinter。Python 3.8 是最后一代适合 Win7 的官方 Python 主版本；GitHub Actions 使用 `windows-2022` 作为构建环境，并固定 Python 3.8 和 PyInstaller 4.10，尽量避免高版本运行时破坏 Win7 兼容性。

Win7 机器如果缺少系统运行库，可能需要安装 Microsoft Visual C++ 2015-2019 Redistributable 或系统补丁。通常 32 位 exe 兼容面更广，64 位 exe 处理大文件更稳。

注意：GitHub 已在 2025-06-30 退役 `windows-2019` 托管环境，所以自动构建不能继续使用 `windows-2019`。PyInstaller 官方文档对 Win7 的表述是“应该可用，但不正式支持”。所以项目同时构建 x86 和 x64 两个版本，建议先在目标 Win7 电脑上试 x86 版本；如果要求 100% 按 Win7 环境构建，需要改用一台 Win7/Win10 老系统电脑作为 self-hosted runner。
