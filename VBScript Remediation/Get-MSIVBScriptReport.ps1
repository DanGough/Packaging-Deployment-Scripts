# -------------------------
# Get-MSIVBScriptReport.ps1
# -------------------------
# Written by Dan Gough
# -------------------------
# v1.0 - Initial release
# v1.1 - Improved MST handling logic and changed output fields

param (
    [string]$Path = $PWD,
    [string]$OutputFile
)

if (-not (Test-Path -Path $Path -PathType Container)) {
    Write-Error 'Invalid path supplied.'
    exit
}

if (-not (Get-InstalledModule -Name 'MSI' -ErrorAction SilentlyContinue)) {
    Install-Module -Name 'MSI' -Repository PSGallery -Scope CurrentUser -ErrorAction Stop
}

# Get all MSI and MST files
$Files = Get-ChildItem -Path $Path -Filter '*.ms*' -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq '.msi' -or $_.Extension -eq '.mst' }

$Counter = 0

$Report = foreach ($File in $Files) {

    $Counter++
    $PercentComplete = [System.Math]::Floor($Counter / $Files.Count * 100)
    Write-Progress -Activity "Processing $Counter of $($Files.Count)" -Status $File.Name -PercentComplete $PercentComplete

    if ($File.Extension -eq '.msi') {
        try {
            Write-Verbose "Opening $($File.FullName)..."
            $Properties = $CustomActions = $FileNames = $null # Reset to null each time in case assignment via Get-MSITable fails
            $Properties = Get-MSITable -Path $File.FullName -Table 'Property' -ErrorAction Stop # ErrorAction Stop here to trigger the catch block if MSI is locked or invalid
            $CustomActions = @(Get-MSITable -Path $File.FullName -Table 'CustomAction' -ErrorAction SilentlyContinue).Where({ ($_.Type -band 0x6) -eq 0x6 }).Action # ErrorAction SilentlyContinue as some packages have no CustomAction table
            $FileNames = @(Get-MSITable -Path $File.FullName -Table 'File' -ErrorAction SilentlyContinue).Where({ $_.FileName -like '*.vbs' }).FileName.ForEach({ $_.Split('|')[-1] }) # ErrorAction SilentlyContinue as some packages have no File table
    
            if ($CustomActions -or $FileNames) {
                [PSCustomObject]@{
                    File           = $File.FullName
                    Manufacturer   = $Properties.Where({ $_.Property -eq 'Manufacturer' }).Value
                    ProductName    = $Properties.Where({ $_.Property -eq 'ProductName' }).Value
                    ProductVersion = $Properties.Where({ $_.Property -eq 'ProductVersion' }).Value
                    CustomActions  = $CustomActions -join '; '
                    VbsFiles       = $FileNames -join '; '
                }
            }
        }
        catch {
            Write-Error $_
            continue
        }
    }
    elseif ($File.Extension -eq '.mst') {

        $MSICandidates = Get-ChildItem -Path $File.PSParentPath -Filter '*.msi' -ErrorAction SilentlyContinue # Get all MSI packages from the same folder

        # Try MST with every MSI found in the same folder, collecting CustomActions and FileNames that do not belong in the base MSI, summing together as risk of some MSIs not having both CustomAction & File tables
        $MSTCustomActions = @()
        $MSTFileNames = @()

        foreach ($MSICandidate in $MSICandidates) {

            Write-Verbose "Trying $($File.FullName) with $($MSICandidate.FullName)..."

            try {
                $TransformedProperties = $TransformedCustomActions = $TransformedFileNames = $BaseCustomActions = $BaseFileNames = $null # Reset to null each time in case assignment via Get-MSITable fails
                $TransformedProperties = Get-MSITable -Path $MSICandidate.FullName -Transform $File.FullName -Table 'Property' -ErrorAction Stop # ErrorAction Stop here to trigger the catch block if applying MST fails
                $BaseCustomActions = @(Get-MSITable -Path $MSICandidate.FullName -Table 'CustomAction' -ErrorAction SilentlyContinue).Where({ ($_.Type -band 0x6) -eq 0x6 }).Action
                $BaseFileNames = @(Get-MSITable -Path $MSICandidate.FullName -Table 'File' -ErrorAction SilentlyContinue).Where({ $_.FileName -like '*.vbs' }).FileName.ForEach({ $_.Split('|')[-1] })
                $TransformedCustomActions = @(Get-MSITable -Path $MSICandidate.FullName -Transform $File.FullName -Table 'CustomAction' -ErrorAction SilentlyContinue).Where({ ($_.Type -band 0x6) -eq 0x6 }).Action
                $TransformedFileNames = @(Get-MSITable -Path $MSICandidate.FullName -Transform $File.FullName -Table 'File' -ErrorAction SilentlyContinue).Where({ $_.FileName -like '*.vbs' }).FileName.ForEach({ $_.Split('|')[-1] })
                $MSTCustomActions += ($TransformedCustomActions | Where-Object { $_ -notin $BaseCustomActions }) 
                $MSTFileNames += ($TransformedFileNames | Where-Object { $_ -notin $BaseFileNames })
            }
            catch {
                Write-Verbose "Error applying transform: $_"
                continue
            }

        }

        # Remove duplicates
        $MSTCustomActions = $MSTCustomActions | Select-Object -Unique
        $MSTFileNames = $MSTFileNames | Select-Object -Unique

        if ($MSTCustomActions -or $MSTFileNames) {
            [PSCustomObject]@{
                File           = $File.FullName
                Manufacturer   = $null
                ProductName    = $null
                ProductVersion = $null
                CustomActions  = $MSTCustomActions -join '; '
                VbsFiles       = $MSTFileNames -join '; '
            }
        }
    }

}
Write-Progress -Completed -Activity 'Clear Progress Bar'

if ($Report.Count) {

    $Report | Out-GridView -Title "Results from $Path"

    if ($OutputFile) {
        $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
        $OutputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        $OutputType = [System.IO.Path]::GetExtension($OutputFile)

        if ($OutputType -notmatch '^\.csv|\.xlsx$') {
            Write-Warning -Message 'Unsupported file format specified'
            $OutputFile = $OutputDir = $OutputType = $null
        }

        if (-not (Test-Path -Path $OutputDir -PathType Container)) {
            try {
                New-Item -Path $OutputDir -ItemType Directory -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Warning -Message 'Invalid path supplied'
                $OutputFile = $OutputDir = $OutputType = $null  
            }

        }
    }

    if ([string]::IsNullOrEmpty($OutputFile)) {
        Add-Type -AssemblyName System.Windows.Forms
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save report?"
        $SaveFileDialog.FileName = "Get-MSIVBScriptReport.xlsx"
        $SaveFileDialog.InitialDirectory = $PWD
        $Result = $SaveFileDialog.ShowDialog()
            
        if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
            $OutputFile = $SaveFileDialog.FileName
            $OutputType = [System.IO.Path]::GetExtension($OutputFile)
        }
    }

    switch ($OutputType) {
        '.xlsx' { 
            if (-not (Get-InstalledModule -Name 'ImportExcel' -ErrorAction SilentlyContinue)) {
                Install-Module -Name 'ImportExcel' -Repository PSGallery -Scope CurrentUser -ErrorAction Stop
            }
            $Report | Export-Excel -Path $OutputFile -ClearSheet -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow
        }
        '.csv' {
            $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Force
        }
    }

}
else {
    Write-Host "No affected files found in $Path."
}