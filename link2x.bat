rmdir /s /q build
rmdir /s /q dist
del LinkPilot_v1.spec

pyinstaller --onefile --windowed --icon=LinkPilot.ico LinkPilot_v1.py
