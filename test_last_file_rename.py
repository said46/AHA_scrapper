import os
import time


def last_file_rename(new_name, download_folder, time_to_wait=60):
    time_counter = 0
    fn = max([f for f in os.listdir(download_folder)], key=lambda xa :   os.path.getctime(os.path.join(download_folder, xa)))
    while '.crdownload' in fn:
        time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:
            raise Exception('Waited too long for file to download')
    fn = max([f for f in os.listdir(download_folder)], key=lambda xa :   os.path.getctime(os.path.join(download_folder, xa)))
    os.rename(os.path.join(download_folder, fn), os.path.join(download_folder, new_name))


os.system("cls")
last_file_rename('2.txt', "C:\\Users\\alexander.sverchkov\\Downloads", 5)