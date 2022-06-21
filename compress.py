import zipfile
import os
import shutil

def compress(path, file_list, zip_filename, compression=False):
    """
    assumption : The script must be in the same directory as the files to zip
    """

    if compression is False:
        compression = zipfile.ZIP_STORED
    elif compression is True:
        compression = zipfile.ZIP_DEFLATED 

    with zipfile.ZipFile(zip_filename, 'w', compression=compression) as my_zip:
        for file in file_list:
            my_zip.write(f'{file}')



if __name__ == '__main__':

    path = 'C:\\Users\\giovanni.osorio\\Desktop\\python_proyects\\split_set_password\\xl_files_password-20220613\\'
    files = os.listdir(path)
    print(files)

    shutil.make_archive('SGRE - 202200613', 'zip', path)