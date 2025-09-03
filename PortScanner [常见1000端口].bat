@echo off
chcp 65001 > nul
echo.
python3 PortScanner.py -ip-list ip.txt -p-list port-1000.txt -threads 66
echo.
pause