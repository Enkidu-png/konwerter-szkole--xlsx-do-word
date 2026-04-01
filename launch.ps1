# launch.ps1 — Launcher for Konwerter Szkolen
# Opens "baza danych.xlsx" from Desktop, refreshes CRM data, starts local HTTP server, opens browser.

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- Configuration ---
$distPath = Join-Path $PSScriptRoot "dist"
$xlsxFileName = "baza danych.xlsx"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$xlsxPath = Join-Path $desktopPath $xlsxFileName
$portCandidates = @(3000, 3001, 3002, 3003, 3004)

# --- MIME Types ---
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
    if ($mimeTypes.ContainsKey($extension)) {
        return $mimeTypes[$extension]
    }
    return "application/octet-stream"
}

# --- Banner ---
Write-Host ""
Write-Host "  ======================================" -ForegroundColor Cyan
Write-Host "    Konwerter Szkolen - Uruchamianie" -ForegroundColor Cyan
Write-Host "  ======================================" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Validate dist/ ---
$indexPath = Join-Path $distPath "index.html"
if (-not (Test-Path $indexPath)) {
    Write-Host "  BLAD: Folder dist/ nie istnieje lub jest pusty." -ForegroundColor Red
    Write-Host "  Skontaktuj sie z administratorem." -ForegroundColor Red
    Write-Host ""
    Read-Host "  Nacisnij Enter aby zamknac"
    exit 1
}

# --- Step 2: Excel — open and refresh ---
Write-Host "  [1/3] Otwieram baze danych w Excel..." -ForegroundColor Yellow

if (-not (Test-Path $xlsxPath)) {
    Write-Host "  Uwaga: Nie znaleziono pliku '$xlsxFileName' na Pulpicie." -ForegroundColor DarkYellow
    Write-Host "         Oczekiwana sciezka: $xlsxPath" -ForegroundColor DarkYellow
    Write-Host "         Pomijam odswiezanie bazy." -ForegroundColor DarkYellow
    Write-Host ""
} else {
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = $true
        $excel.DisplayAlerts = $false

        try {
            $workbook = $excel.Workbooks.Open($xlsxPath)
            Write-Host "  Plik otwarty. Odswiezam dane z CRM..." -ForegroundColor Green

            $workbook.RefreshAll()

            # Wait for async queries to finish
            try {
                $excel.CalculateUntilAsyncQueriesDone()
            } catch {
                Start-Sleep -Seconds 3
            }

            Write-Host "  Odswiezanie zakonczone." -ForegroundColor Green
            Write-Host "  Jesli pojawilo sie okno logowania, wypelnij dane w Excel." -ForegroundColor Yellow
            Write-Host ""
        } catch {
            Write-Host "  Uwaga: Plik moze byc juz otwarty w Excel." -ForegroundColor DarkYellow
            Write-Host "         Sprawdz czy Excel jest aktywny." -ForegroundColor DarkYellow
            Write-Host ""
        }

        # Release COM references but do NOT quit Excel
        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $workbook = $null
        $excel = $null
    } catch {
        Write-Host "  Uwaga: Excel nie jest zainstalowany lub nie mozna go uruchomic." -ForegroundColor DarkYellow
        Write-Host "         Pomijam odswiezanie bazy." -ForegroundColor DarkYellow
        Write-Host ""
    }
}

# --- Step 3: Find available port ---
Write-Host "  [2/3] Uruchamiam serwer lokalny..." -ForegroundColor Yellow

$port = $null
foreach ($candidate in $portCandidates) {
    try {
        $testListener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $candidate)
        $testListener.Start()
        $testListener.Stop()
        $port = $candidate
        break
    } catch {
        continue
    }
}

if ($null -eq $port) {
    Write-Host "  BLAD: Porty 3000-3004 sa zajete." -ForegroundColor Red
    Write-Host "  Zamknij inne programy i sprobuj ponownie." -ForegroundColor Red
    Write-Host ""
    Read-Host "  Nacisnij Enter aby zamknac"
    exit 1
}

# --- Step 4: Start HTTP server ---
$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "  BLAD: Nie mozna uruchomic serwera na porcie $port." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "  Nacisnij Enter aby zamknac"
    exit 1
}

# --- Step 5: Open browser ---
Write-Host "  [3/3] Otwieram przegladarke..." -ForegroundColor Yellow
Start-Process "http://localhost:$port"

Write-Host ""
Write-Host "  ======================================" -ForegroundColor Green
Write-Host "    Aplikacja dziala!" -ForegroundColor Green
Write-Host "    http://localhost:$port" -ForegroundColor Green
Write-Host "  ======================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Zamknij to okno aby zakonczyc aplikacje." -ForegroundColor Gray
Write-Host ""

# --- Request handling loop ---
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $urlPath = $request.Url.LocalPath
        if ($urlPath -eq "/") {
            $urlPath = "/index.html"
        }

        # Map URL to file path
        $filePath = Join-Path $distPath ($urlPath.TrimStart("/").Replace("/", "\"))
        $filePath = [System.IO.Path]::GetFullPath($filePath)

        # Security: ensure path is within dist/
        if (-not $filePath.StartsWith($distPath)) {
            $filePath = $indexPath
        }

        # SPA fallback: if file not found, serve index.html
        if (-not (Test-Path $filePath -PathType Leaf)) {
            $filePath = $indexPath
        }

        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        $contentType = Get-MimeType $extension

        try {
            $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = $contentType
            $response.ContentLength64 = $fileBytes.Length
            $response.StatusCode = 200
            $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
        } catch {
            $response.StatusCode = 500
        } finally {
            $response.Close()
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
    Write-Host ""
    Write-Host "  Serwer zatrzymany." -ForegroundColor Gray
}
