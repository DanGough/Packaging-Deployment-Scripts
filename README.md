# VBScript-Remediation
This repository is for various scripts and tools to help deal with the forthcoming deprecation of VBScript.

## Get-MSIVBScriptReport.ps1

This script accepts 2 parameters:

- Path: The path to recursively scan (defaults to current working directory)
- OutputFile: The file to save the report to, in .csv or .xlsx format (a file browser will pop up if no valid path supplied)

The script will find every .MSI file and test it alone, along with every .MST file found in the same folder, and produce a report on:

- Manufacturer, ProductName, ProductVersion
- File table entries ending in .vbs
- VBScript CustomAction table entries

This script requires 2 external dependencies obtained from the PowerShell Gallery: MSI and ImportExcel.
