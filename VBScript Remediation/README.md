# VBScript Remediation
This repository is for various scripts and tools to help deal with the forthcoming deprecation of VBScript.

## Get-MSIVBScriptReport.ps1

This script accepts 2 parameters:

- Path: The path to recursively scan (defaults to current working directory)
- OutputFile: The file to save the report to, in .csv or .xlsx format (a file browser will pop up if no valid path supplied)

The script will find every VBScript custom action and file within each MSI and MST file and produce a report on:

- Manufacturer, ProductName, ProductVersion
- File table entries ending in .vbs
- VBScript CustomAction table entries

This script requires 2 external dependencies obtained from the PowerShell Gallery: MSI and ImportExcel.

To do:

- Allow more flexibility for $Path to contain a single file, array of files, or array of folders rather than just a single folder, and to accept files from the pipeline.
- Scan PSADT scripts for use of certain functions that invoke VBScript.