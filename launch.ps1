# launch.ps1 — Silent launcher for Konwerter Szkolen
# Opens "baza danych.xlsx" (Excel auto-refreshes), starts hidden HTTP server, opens browser.

$distPath = Join-Path $PSScriptRoot "dist"
$indexPath = Join-Path $distPath "index.html"
$xlsxPath = Join-Path $PSScriptRoot "baza danych.xlsx"
$portCandidates = @(3000, 3001, 3002, 3003, 3004)

$mimeTypes = @{
    ".html"  = "text/html; charset=utf-8"
    ".htm"   = "text/html; charset=utf-8"
    ".js"    = "application/javascript; charset=utf-8"
    ".css"   = "text/css; charset=utf-8"
    ".json"  = "application/json"
    ".svg"   = "image/svg+xml"
    ".png"   = "image/png"
    ".jpg"   = "image/jpeg"
    ".jpeg"  = "image/jpeg"
    ".gif"   = "image/gif"
    ".ico"   = "image/x-icon"
    ".woff"  = "font/woff"
    ".woff2" = "font/woff2"
    ".ttf"   = "font/ttf"
    ".docx"  = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ".xlsx"  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

function Get-MimeType($extension) {
    if ($mimeTypes.ContainsKey($extension)) { return $mimeTypes[$extension] }
    return "application/octet-stream"
}

# --- Validate dist/ ---
if (-not (Test-Path $indexPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Folder dist/ nie istnieje.`nSkontaktuj sie z administratorem.",
        "Konwerter Szkolen - Blad",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# --- Resolve Excel file path ---
if (-not (Test-Path $xlsxPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Wybierz plik bazy danych Excel"
    $dialog.Filter = "Pliki Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $xlsxPath = $dialog.FileName
    } else {
        $xlsxPath = $null
    }
}

# --- Open Excel, refresh CRM data, wait, save, then copy to dist/ ---
if ($xlsxPath -and (Test-Path $xlsxPath)) {
    $excelOk = $false
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = $true
        $excel.DisplayAlerts = $false

        try {
            $workbook = $excel.Workbooks.Open($xlsxPath)
            $workbook.RefreshAll()

            # Wait for all async queries (CRM sync + possible auth popup)
            $waited = 0
            while ($waited -lt 120) {
                try {
                    $excel.CalculateUntilAsyncQueriesDone()
                    break
                } catch {
                    Start-Sleep -Seconds 2
                    $waited += 2
                }
            }

            # Save so the refreshed data is on disk
            $workbook.Save()
            $excelOk = $true
        } catch {
            # File might already be open — try to copy as-is
        }

        # Release COM but keep Excel open for the user
        if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $workbook = $null
        $excel = $null
    } catch {
        # Excel not installed — copy file as-is
    }

    # Copy (refreshed) xlsx to dist/ for the web app
    try {
        Copy-Item $xlsxPath (Join-Path $distPath "baza danych.xlsx") -Force
    } catch {
        # File locked — try shadow copy via byte read
        try {
            $bytes = [System.IO.File]::ReadAllBytes($xlsxPath)
            [System.IO.File]::WriteAllBytes((Join-Path $distPath "baza danych.xlsx"), $bytes)
        } catch {}
    }
}

# --- Start HTTP server (try ports until one works) ---
$listener = $null
$port = $null
foreach ($candidate in $portCandidates) {
    try {
        $tryListener = New-Object System.Net.HttpListener
        $tryListener.Prefixes.Add("http://localhost:$candidate/")
        $tryListener.Start()
        $listener = $tryListener
        $port = $candidate
        break
    } catch {
        if ($tryListener) { $tryListener.Close() }
        continue
    }
}

if ($null -eq $listener) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Porty 3000-3004 sa zajete.`nZamknij inne programy i sprobuj ponownie.",
        "Konwerter Szkolen - Blad",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# --- Open browser ---
Start-Process "http://localhost:$port"

# --- Serve requests until process is killed ---
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $response = $context.Response

        $urlPath = $context.Request.Url.LocalPath
        if ($urlPath -eq "/") { $urlPath = "/index.html" }

        $filePath = Join-Path $distPath ($urlPath.TrimStart("/").Replace("/", "\"))
        $filePath = [System.IO.Path]::GetFullPath($filePath)

        if (-not $filePath.StartsWith($distPath)) { $filePath = $indexPath }
        if (-not (Test-Path $filePath -PathType Leaf)) { $filePath = $indexPath }

        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        $response.ContentType = Get-MimeType $extension

        try {
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentLength64 = $bytes.Length
            $response.StatusCode = 200
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } catch {
            $response.StatusCode = 500
        } finally {
            $response.Close()
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
