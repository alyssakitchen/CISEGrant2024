# Install all necessary packages required to run 2-text-to-excel.py
import subprocess

subprocess.run(["pip", "install", "pandas"])
subprocess.run(["pip", "install", "openpyxl"])
subprocess.run(["pip", "install", "pymannkendall"])
subprocess.run(["pip", "install", "statistics"])
subprocess.run((["pip", "install", "xlwt"]))
subprocess.run((["pip", "install", "xlsxwriter"]))
