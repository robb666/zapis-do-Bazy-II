import shutil
import os
import win32com

cache_path = win32com.__gen_path__

if os.path.exists(cache_path):
    shutil.rmtree(cache_path)
    print(f"Successfully deleted cache at: {cache_path}")
else:
    print(f"No cache found at {cache_path}.")