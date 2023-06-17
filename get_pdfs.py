from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions

import time
import openpyxl as xl
import os

def save_excel_file() -> None:
    try:
        wb.save('links.xlsx')
    except Exception as e:
        print(f'Cannot save the excel file: {str(e)}')


pdfs_folder_path = os.getcwd() + '\\_get_pdfs'
options = EdgeOptions()
options.add_experimental_option('prefs', {
                                            "download.prompt_for_download": False,
                                            "plugins.always_open_pdf_externally": True,
                                            "download.default_directory": pdfs_folder_path, 
                                            "download.directory_upgrade": True
                                         })

edgeBrowser = webdriver.Edge(r"msedgedriver.exe", options=options)
edgeBrowser.maximize_window()

wb = xl.load_workbook('links.xlsx')
sheet = wb['links']

# for each row
for count, row in enumerate(range(2, sheet.max_row + 1)):
    # if an empty row - stop
    if sheet.cell(row, 1).value is None:
        break        

    if sheet.cell(row, 2).value == "downloaded":
        continue
    
    if count == 55555:
        edgeBrowser.quit()
        save_excel_file()
        wb.close()
        quit()

    try:
        llink = sheet.cell(row, 1).hyperlink.target
    except Exception as e:
        sheet.cell(row, 2).value = f"{str(e)}" 
        save_excel_file()
        continue
    
    doc_number = sheet.cell(row, 1).value
    try:
        edgeBrowser.get(llink)
    except Exception as e:
        sheet.cell(row, 2).value = f"{str(e)}"        
        save_excel_file()
        continue

    sheet.cell(row, 2).value = f'downloaded'
    save_excel_file()

edgeBrowser.quit()
wb.close()
