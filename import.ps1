$wbPath = "Z:\github.com\yasuhiro112358\WBS\template.xlsm"
$srcPath = "Z:\github.com\yasuhiro112358\WBS\utf8"

function Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [INFO] $message"
}

function LogError {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [ERROR] $message" -ForegroundColor Red
}

try {
    Log "Creating Excel COM object..."
    $excel = New-Object -ComObject Excel.Application
    # $excel.Visible = $true 
    $excel.Visible = $false 
    Log "Created Excel COM object."
} catch {
    LogError "Failed to create Excel COM object: $_"
    exit
}

try {
    Log "Opening Excel file...: $wbPath"
    $wb = $excel.Workbooks.Open($wbPath)
    $vbProj = $wb.VBProject
    Log "Opened Excel file."
} catch {
    LogError "Failed to open Excel file: $_"
    $excel.Quit()
    exit
}

try {
    Log "Getting the list of modules...: $srcPath"
    $files = Get-ChildItem $srcPath -Filter *.bas
    $files += Get-ChildItem $srcPath -Filter *.cls
    $files += Get-ChildItem $srcPath -Filter *.frm
    Log "Got the list of modules: $($files.Count) files found."
} catch {
    LogError "Failed to get modules: $_"
    $wb.Close($false)
    $excel.Quit()
    exit
}

foreach ($file in $files) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $tempPath = "$env:TEMP\$($file.Name)"
    # [Fix] Only for .frm files
    $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx") 
    $tempFrxPath = "$env:TEMP\$($moduleName).frx"

    try {
        Log "Handling module: $moduleName"

        if ($moduleName -like "sht*") {
            Log "Skipping sheet module: $moduleName"
            continue
        }

        try {
            $comp = $vbProj.VBComponents.Item($moduleName)
            if ($comp.Type -in 1, 2, 3) { # vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                Log "Deleting existing module...: $moduleName"
                $vbProj.VBComponents.Remove($comp)
            }
        } catch {
            Log "Skipped since module does not exist: $moduleName"
        }

        Log "Creating temporary file...: $tempPath"
        Get-Content $file.FullName -Encoding UTF8 | Out-File $tempPath -Encoding Default

        if (Test-Path $frxPath) {
            Log "Copying associated .frx file...: $frxPath"
            Copy-Item $frxPath -Destination $tempFrxPath
        }

        Log "Importing module...: $moduleName"
        $vbProj.VBComponents.Import($tempPath)

        Log "Deleting temporary file...: $tempPath"
        Remove-Item $tempPath

        if (Test-Path $tempFrxPath) {
            Log "Deleting temporary .frx file...: $tempFrxPath"
            Remove-Item $tempFrxPath
        }
    } catch {
        LogError "Error on handling module: $moduleName - $_"
    }
}

try {
    Log "Saving Excel file..."
    $wb.Save()
    $wb.Close($false)
    $excel.Quit()
    Log "Saving Excel file, Quitted Excel app."
} catch {
    LogError "Error on saving Excel file or quitting Excel app.: $_"
}

try {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Log "Released Excel COM object."
} catch {
    LogError "Error on Releasing Excel COM object.: $_"
}