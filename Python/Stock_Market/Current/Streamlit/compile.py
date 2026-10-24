# compile.py
import os
import shutil
import subprocess
import streamlit
from PyInstaller.utils.hooks import copy_metadata

# 1. Locate the exact directory of the Streamlit installation on your computer
st_path = os.path.dirname(streamlit.__file__)

# 2. Define all your local module files that must be packed inside the core directory
local_files = [
    "app.py", "kitedata.py", "kitemain.py", 
    "supportresistance.py", "trialcheck.py", 
    "basedb.py", "baselogin.py", "nseindex.py",
    "kitelogin.py", "users.py", "baselogger.py"
]

# 3. Automatically construct the data mapping arguments for PyInstaller
add_data_args = []

# Collect local modules
for file in local_files:
    if os.path.exists(file):
        add_data_args.extend(["--add-data", f"{file};."])

# Collect heavy Streamlit static components and structural metadata config folders
add_data_args.extend(["--add-data", f"{os.path.join(st_path, 'static')};streamlit/static"])
add_data_args.extend(["--add-data", f"{os.path.join(st_path, 'runtime')};streamlit/runtime"])

# 4. Extract and attach Streamlit's core distribution metadata info
streamlit_metadata = copy_metadata("streamlit")
for data_tuple in streamlit_metadata:
    # Safely convert tuples to PyInstaller's command line format string
    add_data_args.extend(["--add-data", f"{data_tuple[0]};{data_tuple[1]}"])

hidden_imports = [
    "--hidden-import=kiteconnect",
    "--hidden-import=pandas",
    "--hiddenimport=numpy"
]

# 5. Formulate the comprehensive institutional packaging command string
command = [
    "pyinstaller",
    "--name=VoluMatrix_Terminal",
    "--onedir",
    "--clean",
    "--noconfirm",
    *hidden_imports,
    *add_data_args,
    "run_app.py"
]

print("🚀 Executing automated assembly pipeline with metadata hooks...")
subprocess.run(command)
print("🎯 Compilation completed successfully!")
