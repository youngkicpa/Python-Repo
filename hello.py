import win32com.client as win32
import shutil
import os

gen_py_path = os.path.join(win32.__gen_path__, "gen_py")
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)