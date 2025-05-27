param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputPath = ""
)

function Convert-EmilyFile {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    Write-Host "Converting '$InputFile'..."
    
    try {
        $lines = Get-Content $InputFile -Encoding UTF8
        $recordsStart = -1
        $recordsEnd = -1
        
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -match '^\[RECORDS\]') {
                $recordsStart = $i + 1
            }
            elseif ($lines[$i] -match '^\[END\]') {
                $recordsEnd = $i
                break
            }
        }
        
        if ($recordsStart -eq -1) {
            Write-Warning "Could not find [RECORDS] section - skipping"
            return $false
        }
        
        if ($recordsEnd -eq -1) {
            $recordsEnd = $lines.Length
        }
        
        $outputLines = @()
        
        for ($i = $recordsStart; $i -lt $recordsEnd; $i++) {
            $line = $lines[$i].Trim()
            
            if ($line -eq "") {
                continue
            }
            
            $columns = $line -split '\s+' | Where-Object { $_ -ne "" }
            
            if ($columns.Length -lt 7) {
                continue
            }
            
            $dataColumns = $columns[1..6]
            
            $formattedColumns = @()
            for ($j = 0; $j -lt $dataColumns.Length; $j++) {
                if ($j -eq 0) {
                    $formattedColumns += "  " + $dataColumns[$j] + " ,"
                }
                elseif ($j -eq $dataColumns.Length - 1) {
                    $formattedColumns += " " + $dataColumns[$j] + "  "
                }
                else {
                    $formattedColumns += " " + $dataColumns[$j] + " ,"
                }
            }
            $outputLine = "MOVEJ/" + ($formattedColumns -join "")
            
            $outputLines += $outputLine
        }
        
        $outputDir = Split-Path $OutputFile -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        $outputLines | Out-File -FilePath $OutputFile -Encoding UTF8
        
        Write-Host "  Success: $($outputLines.Length) lines written"
        return $true
        
    }
    catch {
        Write-Error "Error: $($_.Exception.Message)"
        return $false
    }
}

if (-not (Test-Path $InputPath)) {
    Write-Error "Input path not found!"
    exit 1
}

$inputItem = Get-Item $InputPath

if ($inputItem.PSIsContainer) {
    Write-Host "Processing directory: $InputPath"
    Write-Host "=================================================="
    
    $emilyFiles = Get-ChildItem -Path $InputPath -Filter "*.emily" -File
    
    if ($emilyFiles.Count -eq 0) {
        Write-Warning "No .emily files found in directory"
        exit 1
    }
    
    Write-Host "Found $($emilyFiles.Count) .emily files to process"
    Write-Host ""
    
    if ($OutputPath -eq "") {
        $OutputPath = $InputPath
    }
    
    $successCount = 0
    $failCount = 0
    
    foreach ($file in $emilyFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $outputFileName = "${baseName}_MOVEJ.txt"
        $outputFilePath = Join-Path $OutputPath $outputFileName
        
        if (Convert-EmilyFile -InputFile $file.FullName -OutputFile $outputFilePath) {
            $successCount++
        }
        else {
            $failCount++
        }
    }
    
    Write-Host ""
    Write-Host "=================================================="
    Write-Host "Batch processing completed!"
    Write-Host "Successfully converted: $successCount files"
    if ($failCount -gt 0) {
        Write-Host "Failed to convert: $failCount files"
    }
    Write-Host "Output directory: $OutputPath"
    
}
else {
    if ($OutputPath -eq "") {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name)
        $outputDir = Split-Path $inputItem.FullName -Parent
        $OutputPath = Join-Path $outputDir "${baseName}_MOVEJ.txt"
    }
    
    if (Convert-EmilyFile -InputFile $inputItem.FullName -OutputFile $OutputPath) {
        Write-Host ""
        Write-Host "Conversion completed successfully!"
        Write-Host "Output file: $OutputPath"
    }
    else {
        Write-Error "Conversion failed!"
        exit 1
    }
} 