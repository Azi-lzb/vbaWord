# vbaWord - 办公自动化：Word 问卷批量汇总工具

`vbaWord` 是一个基于 Excel/Word VBA 开发的办公自动化套件，专门用于处理受保护（限制编辑：填写窗体）的 Word 问卷文档。它可以将分散的问卷数据一键提取并汇总到 Excel 表格或生成一份精简的 Word 报告。

## 🌟 核心功能

### 1. 汇总至 Excel (`extractWord.bas`)
- **定向输出**：自动将结果写入当前工作簿的 `output` 工作表，不覆盖现有数据。
- **智能表头**：自动提取第一个文档的书签名称作为表头，并应用加粗和背景色。
- **文件名映射**：支持通过 `mapping` 表将原始文件名替换为易读的名称（如：`Sample_Data_1` -> `北京分公司`）。

### 2. 汇总至 Word (`extractToWord.bas`)
- **动态题目识别**：自动提取填空位（窗体域）前方的文本作为题目内容。
- **精简输出格式**：
  ```text
  题目名称
  【显示名称1】: 回答内容; 【显示名称2】: 回答内容; 
  ```
- **按题分组**：所有文件的同一题目答案会横向排列，方便快速对比。

### 3. 测试数据生成器 (`GenerateTestData.bas`)
- 在 Word 中运行此宏，可一键生成 3 份模拟已填写的问卷，方便测试汇总逻辑。

## 🛠️ 如何使用

### 1. 准备工作
- 在 Excel 中按 `Alt + F11` 打开 VBA 编辑器。
- 点击 **工具 (Tools)** -> **引用 (References)**，务必勾选：
  - `Microsoft Word 16.0 Object Library` (或您当前最高的版本)
  - `Microsoft Scripting Runtime`

### 2. 配置映射 (可选)
- 在 Excel 中新建一个名为 `mapping` 的工作表。
- **A 列**：填入文件名中的关键字（如 `Sample_Data_1`）。
- **B 列**：填入您想显示的名称（如 `财务部`）。
- 汇总时，程序会自动匹配并替换名称。

### 3. 运行宏
- 导入对应的 `.bas` 文件。
- 运行 `BatchSummarizeWordForms` (Excel 版) 或 `SummarizeToNewWordDoc` (Word 版)。
- 在弹出的对话框中批量选择 Word 文件即可。

## 📂 项目结构
- `VBA_Export/Modules/`: 存放所有核心 VBA 模块。
- `SampleForTest/`: 存放用于测试的 Word 样本文件。
- `VBA_Export/backup/`: 历史版本备份。

## ✒️ 开发者
- **Azi-lzb**
