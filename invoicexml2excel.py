import os
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import load_workbook

def parse_invoice(xml_path):
    """
    解析单个 XML 发票文件，提取所有 IssuItemInformation 数据。

    :param xml_path: XML 文件的路径
    :return: 包含 Sheet1 和 Sheet2 数据的元组
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # 提取所有 IssuItemInformation 节点
    items = root.findall('.//IssuItemInformation')

    invoice_items_sheet1 = []
    invoice_items_sheet2 = []

    prev_item_sheet2 = None  # 用于记录 Sheet2 中的前一个项目

    for idx, item in enumerate(items):
        # 提取所需字段
        item_name = item.findtext('ItemName', default='')  # 项目名称
        spec_mod = item.findtext('SpecMod', default='')    # 规格型号
        units = item.findtext('MeaUnits', default='')     # 单位
        quantity = item.findtext('Quantity', default='0') # 数量
        amount = item.findtext('Amount', default='0')     # 金额
        tax_rate = item.findtext('TaxRate', default='0')  # 税率
        tax_amount = item.findtext('ComTaxAm', default='0') # 税额

        # 计算总价（金额 + 税额），不使用 Excel 公式
        try:
            amount_num = float(amount)
            tax_num = float(tax_amount)
            total_price = round(amount_num + tax_num, 2)
        except ValueError:
            total_price = 0.0

        # 添加到 Sheet1
        invoice_items_sheet1.append([
            item_name, spec_mod, units, quantity, amount, tax_rate, tax_amount, total_price
        ])

        # 构建当前项目的数据字典
        current_item_sheet2 = {
            '物资名称': item_name,
            '规格型号': spec_mod,
            '包装': '',
            '单位': units,
            '数量': quantity,
            '单价': '',
            '总价': total_price,
            '备注': ''  # 初始化备注为空
        }

        # 检查是否需要合并
        if (prev_item_sheet2 and
            item_name == prev_item_sheet2['物资名称'] and
            spec_mod == '' and
            total_price < 0):

            # 合并当前项到前一项
            prev_item_sheet2['总价'] = round(prev_item_sheet2['总价'] + total_price, 2)
            prev_item_sheet2['备注'] = '已合并'

            # 规格型号使用前一项的规格型号
            current_item_sheet2['规格型号'] = prev_item_sheet2['规格型号']

            # 不将当前项添加到列表中（已合并）
            continue
        else:
            # 将当前项添加到 Sheet2 列表中
            invoice_items_sheet2.append(current_item_sheet2)
            # 更新前一个项目
            prev_item_sheet2 = current_item_sheet2

    return invoice_items_sheet1, invoice_items_sheet2
    
def convert_xml_to_excel(xml_folder, output_folder):
    """
    将指定文件夹下的所有 XML 发票文件转换为 Excel 文件。

    :param xml_folder: 包含 XML 文件的文件夹路径
    :param output_folder: 输出 Excel 文件的文件夹路径
    """
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)

    # 遍历文件夹下的所有 XML 文件
    for filename in os.listdir(xml_folder):
        if filename.lower().endswith('.xml'):
            xml_path = os.path.join(xml_folder, filename)
            print(f'正在处理文件: {xml_path}')

            # 解析 XML 文件，提取项目数据
            invoice_items_sheet1, invoice_items_sheet2 = parse_invoice(xml_path)

            if not invoice_items_sheet1:
                print(f'文件 {filename} 中没有找到 IssuItemInformation 节点。')
                continue

            # 创建 DataFrame，指定列标题
            columns_sheet1 = ['项目名称', '规格型号', '单位', '数量', '金额', '税率', '税额', '总价']
            df_sheet1 = pd.DataFrame(invoice_items_sheet1, columns=columns_sheet1)

            columns_sheet2 = ['物资名称', '规格型号', '包装', '单位', '数量', '单价', '总价', '备注']
            df_sheet2 = pd.DataFrame(invoice_items_sheet2, columns=columns_sheet2)

            # 生成输出 Excel 文件名（与 XML 文件同名，但扩展名为 .xlsx）
            excel_filename = os.path.splitext(filename)[0] + '.xlsx'
            excel_path = os.path.join(output_folder, excel_filename)

            # 使用 ExcelWriter 同时写入多个工作表，注意调换顺序使得 Sheet2 在 Sheet1 前面写入
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df_sheet2.to_excel(writer, index=False, sheet_name='Sheet1')  # 注意这里已经更改为 Sheet1
                df_sheet1.to_excel(writer, index=False, sheet_name='Sheet2')  # 注意这里已经更改为 Sheet2

            print(f'已生成 Excel 文件: {excel_path}')

    print('所有文件处理完成。')

if __name__ == '__main__':
    # 定义输入和输出文件夹路径
    xml_folder = '20241121_元器件等-未报'      # 请替换为你的 XML 文件夹路径
    output_folder = '20241121_元器件等-未报' # 请替换为你希望存放 Excel 文件的路径

    convert_xml_to_excel(xml_folder, output_folder)
