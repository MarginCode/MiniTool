import os
import shutil

path=os.getcwd()
for root, dirs, files in os.walk(path):
    for file in files:
        if file == "将文件夹内所有文件向上提升一级.py":
            break
        old_file_path = os.path.join(root, file)
        new_root = root.split("\\")[:-1]
        new_root = '\\'.join(new_root)
        new_file_path = os.path.join(new_root,file)
        shutil.move(old_file_path,new_file_path)
