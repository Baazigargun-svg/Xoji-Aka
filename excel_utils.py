from openpyxl import Workbook
from openpyxl.styles import Font

def create_excel_invoice(order_id, shop_name, agent_name, date_str, price_type, cart_items, comment=''):
    wb = Workbook()
    ws = wb.active
    ws.title = f'Nakladnoy_{order_id}'
    
    headers = ['№', 'Mahsulot nomi', "Miqdori", "Narxi", 'Jami summa']
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num, value=header).font = Font(bold=True)

    row_num = 2
    for idx, item in enumerate(cart_items, 1):
        ws.cell(row=row_num, column=1, value=idx)
        ws.cell(row=row_num, column=2, value=item['name'])
        ws.cell(row=row_num, column=3, value=item['qty'])
        ws.cell(row=row_num, column=4, value=item['price'])
        ws.cell(row=row_num, column=5, value=item['qty'] * item['price'])
        row_num += 1

    file_name = f'Nakladnoy_{order_id}.xlsx'
    wb.save(file_name)
    return file_name

def create_svodka_excel(aggregated_data, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Svodka"
    
    ws.cell(row=1, column=1, value="Mahsulot nomi").font = Font(bold=True)
    ws.cell(row=1, column=2, value="Jami miqdori").font = Font(bold=True)
    
    row = 2
    for p_name, qty in aggregated_data.items():
        ws.cell(row=row, column=1, value=p_name)
        ws.cell(row=row, column=2, value=qty)
        row += 1
        
    wb.save(file_path)
    return file_path