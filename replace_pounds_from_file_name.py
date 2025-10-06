import os
from my_utils import message_box 

os.system('cls')

pdfs_folder_path = os.getcwd() + '\\_get_pdfs'

list_of_file_names = [_ for _ in os.listdir(pdfs_folder_path)]

for fn in list_of_file_names:
    new_name = fn.replace('#', '_')
    try:
        os.rename(os.path.join(pdfs_folder_path, fn), os.path.join(pdfs_folder_path, new_name))
    except Exception as e:
        print(f"{str(e)}")
        continue

message_box('Ok?', 'Ok!', 0)