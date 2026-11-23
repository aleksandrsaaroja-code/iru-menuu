@echo off
rem Uuendab kokkade äpi arvud toidulehelt (X:\Statistika\Toidulehed_2026 UUS.xlsx) ja saadab GitHubi.
rem Esimesel korral: py -m pip install openpyxl
cd /d "%~dp0\.."
py tools\uuenda_arvud.py "X:\Statistika\Toidulehed_2026 UUS.xlsx" --push
