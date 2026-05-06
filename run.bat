@echo off
title PDF Smeta to Excel Converter
echo Starting conversion...
echo --------------------------------------------------
.\venv\Scripts\python.exe main.py --input input --output output
echo --------------------------------------------------
echo Process finished! Results are in the output folder.
pause
