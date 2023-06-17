import openpyxl as xl
import os

os.system('cls')


def save_excel_file() -> None:
    try:
        wb.save('links.xlsx')
    except Exception as e:
        print(f'Cannot save the excel file: {str(e)}')


pdfs_folder_path = os.getcwd() + '\\_get_pdfs'

wb = xl.load_workbook('links.xlsx')
sheet = wb['links']


list_of_file_names = [f  for f in os.listdir(pdfs_folder_path)]
list_of_file_names.sort(key=lambda xa: os.path.getctime(os.path.join(pdfs_folder_path, xa)))

for row, fn in zip(range(2, sheet.max_row + 1), list_of_file_names):
    if sheet.cell(row, 1).value is None:
        break        
    doc_number = sheet.cell(row, 1).value
    if sheet.cell(row, 2).value == "downloaded":
        new_name = os.path.join(pdfs_folder_path, doc_number + '.pdf')
        try:
            os.rename(os.path.join(pdfs_folder_path, fn), os.path.join(pdfs_folder_path, new_name))
            sheet.cell(row, 2).value = "downloaded and renamed"
        except Exception as e:
            sheet.cell(row, 2).value = f"{str(e)}"
            continue
        finally:
            save_excel_file()

wb.close()