@echo off
echo Building PowerPoint Remote WebApp...
pyinstaller --noconfirm --onefile --windowed --icon=icon.ico --add-data "static;static" --add-data "icon.ico;." --name "PowerPointRemote" main.py
echo Build complete! You can find the executable in the "dist" folder.
pause
