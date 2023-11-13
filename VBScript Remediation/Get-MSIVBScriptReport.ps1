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

# Where-Object seems redundant here but is necessary to stop returning msix packages
$MSIs = Get-ChildItem -Path $Path -Filter '*.msi' -Recurse -ErrorAction SilentlyContinue | Where-Object Extension -eq '.msi'

$Counter = 0

$Report = foreach ($MSI in $MSIs) {

    $Counter++
    $PercentComplete = [System.Math]::Floor($Counter / $MSIs.Count * 100)
    Write-Progress -Activity "Processing $Counter of $($MSIs.Count)" -Status $MSI.Name -PercentComplete $PercentComplete

    # Try the MSI by itself (null MST) and also with every MST found in the same folder
    $MSTs = @($null) + (Get-ChildItem -Path $MSI.PSParentPath -Filter '*.mst' -ErrorAction SilentlyContinue)

    foreach ($MST in $MSTs) {

        $CustomActions = $Files = $null # Reset as if Get-MSITable fails it does not return $null
        
        $CustomActions = @(Get-MSITable -Path $MSI.FullName -Transform $MST.FullName -Table 'CustomAction' -ErrorAction SilentlyContinue).Where({ ($_.Type -band 0x6) -eq 0x6 })
        $Files = @(Get-MSITable -Path $MSI.FullName -Transform $MST.FullName -Table 'File' -ErrorAction SilentlyContinue).Where({ $_.FileName -like '*.vbs' })
        $Properties = Get-MSITable -Path $MSI.FullName -Transform $MST.FullName -Table 'Property'
        
        if ($CustomActions -or $Files) {
            [PSCustomObject]@{
                Path           = Split-Path -Path $MSI.FullName -Parent
                MSI            = $MSI.Name
                MST            = $MST.Name
                Manufacturer   = $Properties.Where({$_.Property -eq 'Manufacturer'}).Value
                ProductName    = $Properties.Where({$_.Property -eq 'ProductName'}).Value
                ProductVersion = $Properties.Where({$_.Property -eq 'ProductVersion'}).Value
                CustomActions  = $CustomActions.Action -join "; "
                Files          = $Files.FileName.ForEach({ $_.Split('|')[-1] }) -join "; "
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
    Write-Host "No affected MSI packages found in $Path."
}