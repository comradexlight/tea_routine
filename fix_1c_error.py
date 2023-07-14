import shutil
import os
from zipfile import ZipFile


def fix_1c_error(path: str) -> None:
    '''this function fixes the error witn SharedStrings.xml file name  after creating the original xlsx file by 1C'''

    tmp_folder = '/tmp/convert_wrong_excel/'
    os.makedirs(tmp_folder, exist_ok=True)

    with ZipFile(path) as excel_container:
        excel_container.extractall(tmp_folder)

    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path) 

    shutil.make_archive(path, 'zip', tmp_folder)
    os.replace(f'{path}.zip', f'{path}')
