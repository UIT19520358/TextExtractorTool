# ============================================================
#  rename-images.ps1
#  Rename all images in a folder to sequential numbers (1, 2, 3 ...)
#  sorted by file name (natural / numeric sort).
#
#  Usage (run from anywhere):
#    .\rename-images.ps1
#    .\rename-images.ps1 -FolderPath "data\27-02-2026\images"
#    .\rename-images.ps1 -FolderPath "data\27-02-2026\images" -DryRun
# ============================================================

param(
    [string] $FolderPath = "",
    [switch] $DryRun,
    [switch] $AutoConfirm   # Skip the y/N prompt and apply immediately
)

# ── Resolve folder ────────────────────────────────────────────────────────────
if (-not $FolderPath) {
    $FolderPath = Read-Host "Enter folder path (absolute or relative to this script)"
}

# Resolve relative paths from the script's own directory
if (-not [System.IO.Path]::IsPathRooted($FolderPath)) {
    $FolderPath = Join-Path $PSScriptRoot $FolderPath
}

if (-not (Test-Path $FolderPath -PathType Container)) {
    Write-Host "ERROR: Folder not found: $FolderPath" -ForegroundColor Red
    exit 1
}

# ── Collect image files ───────────────────────────────────────────────────────
$extensions = @(".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp")

$files = Get-ChildItem -Path $FolderPath -File |
    Where-Object { $extensions -contains $_.Extension.ToLower() }

if ($files.Count -eq 0) {
    Write-Host "No image files found in: $FolderPath" -ForegroundColor Yellow
    exit 0
}

# ── Natural / numeric sort ────────────────────────────────────────────────────
# Files already named "1.jpg", "2.jpg" ... sort numerically first.
# Non-numeric names (e.g. "z75692...jpg") sort after all numeric ones,
# then alphabetically among themselves.
$sorted = $files | Sort-Object {
    $num = 0
    if ([int]::TryParse($_.BaseName, [ref]$num)) { $num }
    else { [int]::MaxValue }
}, Name

# ── Build rename plan ─────────────────────────────────────────────────────────
$tempPrefix = "__tmp_rename_"
$counter    = 1
$plan       = @()

foreach ($file in $sorted) {
    $ext  = $file.Extension   # e.g. ".jpg"
    $plan += [PSCustomObject]@{
        File     = $file
        TempName = "$tempPrefix$counter$ext"
        NewName  = "$counter$ext"
    }
    $counter++
}

# ── Show plan ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Folder  : $FolderPath"
Write-Host "Images  : $($plan.Count)"
if ($DryRun) {
    Write-Host "Mode    : DRY-RUN  (no files will be changed)" -ForegroundColor Cyan
} else {
    Write-Host "Mode    : LIVE rename" -ForegroundColor Yellow
}
Write-Host ""

$maxOld = ($plan | ForEach-Object { $_.File.Name.Length } | Measure-Object -Maximum).Maximum
$maxOld = [Math]::Max($maxOld, "Original Name".Length)

Write-Host ("{0,-$maxOld}  →  New Name" -f "Original Name")
Write-Host ("{0,-$maxOld}     --------" -f ("-" * $maxOld))

foreach ($entry in $plan) {
    $color = if ($entry.File.Name -ne $entry.NewName) { "White" } else { "DarkGray" }
    Write-Host ("{0,-$maxOld}  →  {1}" -f $entry.File.Name, $entry.NewName) -ForegroundColor $color
}

Write-Host ""

if ($DryRun) {
    Write-Host "Dry-run complete. Re-run without -DryRun to apply." -ForegroundColor Cyan
    exit 0
}

# ── Confirm ───────────────────────────────────────────────────────────────────
if (-not $AutoConfirm) {
    $confirm = Read-Host "Apply renames? (y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-Host "Auto-confirmed." -ForegroundColor Cyan
}

# ── Pass 1: rename to temp names (avoids e.g. "2.jpg" colliding with existing "2.jpg") ──
$errors = 0
foreach ($entry in $plan) {
    try {
        Rename-Item -Path $entry.File.FullName -NewName $entry.TempName -ErrorAction Stop
    } catch {
        Write-Host "  [PASS 1 ERROR] $($entry.File.Name) → $($entry.TempName): $_" -ForegroundColor Red
        $errors++
    }
}

# ── Pass 2: rename temp names to final numbers ────────────────────────────────
foreach ($entry in $plan) {
    $tempPath = Join-Path $FolderPath $entry.TempName
    if (-not (Test-Path $tempPath)) { continue }  # skip if pass 1 failed
    try {
        Rename-Item -Path $tempPath -NewName $entry.NewName -ErrorAction Stop
        Write-Host ("  OK  {0,-$maxOld} → {1}" -f $entry.File.Name, $entry.NewName) -ForegroundColor Green
    } catch {
        Write-Host "  [PASS 2 ERROR] $($entry.TempName) → $($entry.NewName): $_" -ForegroundColor Red
        $errors++
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
if ($errors -eq 0) {
    Write-Host "Done! $($plan.Count) file(s) renamed successfully." -ForegroundColor Green
} else {
    Write-Host "Finished with $errors error(s). Check output above." -ForegroundColor Yellow
}
