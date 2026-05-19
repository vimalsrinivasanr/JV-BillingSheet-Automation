import PyInstaller.__main__
import customtkinter
import os
import sys

# Get the path to the customtkinter library
ctk_path = os.path.dirname(customtkinter.__file__)

# Define our arguments for PyInstaller
# --onefile: Bundles everything into a single .exe
# --noconsole: Prevents a command prompt from opening when the app starts
# --add-data: Includes necessary theme files for CustomTkinter
# --name: The final name of our application
# --hidden-import: Ensure all dependencies are detected

print("Starting Build Process for GCC SAP JV Automation Hub...")
print(f"Detected CustomTkinter at: {ctk_path}")

# Ensure old Windows build artifacts are removed before PyInstaller attempts to overwrite them.
dist_exe = os.path.join('dist', 'GCC_JV_Automation_Hub.exe')
backup_exe = os.path.join('dist', 'GCC_JV_Automation_Hub_old.exe')
if os.path.exists(dist_exe):
    try:
        os.remove(dist_exe)
        print(f"Removed stale executable: {dist_exe}")
    except PermissionError:
        try:
            os.replace(dist_exe, backup_exe)
            print(f"Could not delete existing executable, renamed it to: {backup_exe}")
        except PermissionError:
            print(f"ERROR: Could not remove or rename existing executable '{dist_exe}'.")
            print("Make sure the program is not currently running and try again.")
            sys.exit(1)

if sys.platform != 'win32':
    print("\n" + "!"*60)
    print("WARNING: YOU ARE NOT ON WINDOWS!")
    print("This script will create a MAC executable, NOT a Windows .exe.")
    print("To get a Windows file, please run this on a Windows computer.")
    print("!"*60 + "\n")

args = [
    os.path.join('scripts', 'main_gui.py'),
    '--onefile',
    '--noconsole',
    '--name=GCC_JV_Automation_Hub',
    f'--add-data={ctk_path}{os.pathsep}customtkinter',
    '--clean',
    # Exclude heavy libraries that might be in the environment but aren't used
    '--exclude-module=torch',
    '--exclude-module=torchvision',
    '--exclude-module=tensorflow',
    '--exclude-module=matplotlib',
    '--exclude-module=scipy',
    '--exclude-module=notebook',
    '--exclude-module=ipython'
]

# Note: On Windows, use ; as separator for --add-data. 
# os.pathsep handles this automatically.

PyInstaller.__main__.run(args)

print("\nBUILD COMPLETE!")
print(f"Find your .exe in the 'dist' folder.")
