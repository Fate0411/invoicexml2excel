# invoicexml2excel
- JLU发票出入库，使用python将xml发票转换为excel文件，进而拷贝到给出的申购物资上传模板，快速进行出入库.
- 利用GPT生成并进行简单修改
- 只需要更改xml文件和输出文件夹即可批量转换，也就是将整个文件夹下的xml发票全部转换
- 可以对嘉立创生成的xml发票进行自动合并折扣
- 原始内容在sheet2
- 经过测试仍需要手动复制粘贴到对应的“申购物资上传模板.xls”
- 出入库系统里的总价和所属分类仍需手动填写

# XML to Excel Invoice Converter

## 项目简介

这个项目提供了一个 Python 脚本，用于将存储在 XML 文件中的发票信息转换成结构化的 Excel 文件。该脚本解析 XML 文件中的 `IssuItemInformation` 节点，提取相关数据，并将其分布到两个不同的 Excel 工作表中。此外，它还包括合并特定发票项的高级功能。

## 功能特点

- **数据解析**：从 XML 文件提取发票数据。
- **Excel 输出**：生成包含两个工作表的 Excel 文件。
- **数据合并**：自动合并符合特定条件的连续发票项。
- **格式化输出**：将总价字段输出为文本格式，避免在 Excel 中的自动格式化。

## 安装指南

1. 确保你的机器已安装 Python 3。
2. 安装必需的 Python 库：

   ```bash
   pip install pandas openpyxl
   ```
克隆仓库或下载脚本到本地目录。
使用方法
将你的 XML 文件放在指定的输入文件夹中。

修改脚本中的 xml_folder 和 output_folder 变量，以指向相应的输入和输出文件夹。

运行脚本：

  ```bash
  python convert_invoices.py
  ```
检查输出文件夹中的 Excel 文件。

文件结构
- convert_invoices.py: 主脚本，包含 XML 解析和 Excel 输出的全部逻辑。
- README.md: 项目的使用说明文档。

