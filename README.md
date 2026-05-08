# YiChaFen-Fucker

跨平台易查分批量查询工具，用于把易查分类查询页面里的人工查询流程自动化：解析查询链接、匹配本地数据列、处理验证码策略、逐条查询、实时写出结果。

## 亮点

- 跨平台 GUI：基于 PySide6，Windows、macOS、Linux 均可运行。
- 自动解析查询项：支持查询主页和单条查询链接，解析后自动加载查询条件。
- 多格式数据导入：支持 `txt`、`csv`、`xls`、`xlsx`。
- 数据预览：文本/CSV 默认按英文逗号分隔，也可自定义分隔符；Excel 支持选择工作表。
- 条件映射：将网页查询条件和本地数据列手动映射，避免列名不一致导致误查。
- 验证码策略：支持 OCR、刷新请求、混合模式。
- 实时保存：每成功解析一条结果就立即写入文件，降低中途异常导致的数据损失。
- 失败日志：运行结束后在 `logs/` 中单独生成失败 CSV，记录失败时间、查询数据和失败原因。
- 暂停/继续：长任务可暂停后恢复，停止按钮用于尽快终止当前任务。
- 双格式导出：支持 `csv` 和 `xlsx`；`xlsx` 会自动调整格式。
- 运行缓存可清理：一键清理 cookies、验证码图片和临时缓存。

## 界面预览

<img width="1120" height="820" alt="main-empty" src="https://github.com/user-attachments/assets/682fadf4-9185-44af-bcc0-d365c6f84c6d" />


数据文件加载后可以预览解析结果，并把查询条件映射到对应列。

<img width="1120" height="820" alt="data-preview" src="https://github.com/user-attachments/assets/b175e935-49b0-4d07-a830-52a1916a2ce6" />

导出可选择 CSV 或 XLSX。选择 XLSX 时会提示速度开销。

<img width="1120" height="820" alt="xlsx-export" src="https://github.com/user-attachments/assets/9646a13c-6348-4055-a5fc-26a71a81b283" />

## 解决的问题

易查分页面通常需要逐条输入姓名、班级等条件，再进入结果页查看竖排成绩。数据量稍大时，手工操作容易出错，也很难在中途失败时保留已查询结果。

本工具把这套流程拆成可控步骤：

1. 解析查询主页或单条查询链接。
2. 自动识别网页查询条件。
3. 导入本地学生数据并预览。
4. 手动映射网页条件和本地列头。
5. 按策略处理验证码。
6. 查询成功一条就立即写入结果文件。

## 安装

建议使用虚拟环境：

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 运行

```bash
python main.py
```

## 使用方法

1. 在“链接解析”里输入查询主页链接或单条查询链接，点击“解析链接”。
2. 在“查询项”里选择要查询的项目，程序会自动刷新查询条件。
3. 点击“上传数据文件”，选择 `txt`、`csv`、`xls` 或 `xlsx`。
4. 如果是文本/CSV，可修改分隔符后点击“重新解析并预览”；如果是 Excel，选择工作表。
5. 在“查询条件与数据列匹配”中，把每个查询条件映射到本地列头。
6. 选择验证码策略，并按需要填写刷新请求条件。
7. 选择导出格式和保存目录，文件名会在开始查询时按“时间戳 + 查询项名称”自动生成。
8. 点击“开始查询”。

## 导出说明

- CSV：速度快，适合大量数据。
- XLSX：每条结果写入后都会保存文件，并为已写入区域设置水平/垂直居中和黑色边框；数据量大时会明显慢于 CSV。

## 目录说明

```text
output/          查询结果输出
logs/            程序日志
logs/failures_*  每次运行生成的失败明细
cache/           运行缓存
cache/cookies/   请求 Cookie
cache/captcha/   OCR 验证码图片
cache/temp/      临时文件
ycf_app/         主程序模块
```

“清理缓存”按钮只清理 `cache/cookies/`、`cache/captcha/`、`cache/temp/`，不会删除 `output/` 里的查询结果。
