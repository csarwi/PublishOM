#Requires -Version 5.1
<#
Purpose
-------
- Scan \\creativ.local\data_creativ$\DATA\DefinitiveSoftwareLibrary for "OM <version>" folders.
- Only consider versions with major >= 15. Accept 3–4 numeric parts (e.g., 15.4.31 or 15.4.1.23).
- Zip EVERY such version into \\creativ.local\data_creativ$\DATA\_om_latest as OM15.4.31.zip (no spaces/underscores).
- Include all content EXCEPT under om-apps\omofficeaddin; from that subtree include ONLY the _universal subfolder.
- Treat *.9999 as non-stable for "latest" (but still package them normally).
- Create OM_latest.zip as a copy of the latest stable zip.
- Skip rebuilding a versioned zip if nothing material changed (based on included files: path+size+mtime).
- Write sidecars: OM15.4.31.manifest.json and OM15.4.31.sha256.
- Cleanup: remove zips/sidecars in _om_latest for versions that no longer exist.

Hard requirement:
- 7-Zip must be installed (7z.exe). If not found -> stop with an error. No fallback.
#>

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# --- Configuration ---
$SourceRoot = '\\creativ.local\data_creativ$\DATA\DefinitiveSoftwareLibrary'
$PublishRoot = '\\creativ.local\data_creativ$\DATA\_om_latest'
$CompressionLevel = 5 # 7z -mx value: 0..9

# --- Helpers ---

function Get-SevenZipPath {
    $cmd = Get-Command '7z.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $candidates = @(
        'C:\Program Files\7-Zip\7z.exe',
        'C:\Program Files (x86)\7-Zip\7z.exe'
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) { return $p }
    }

    throw "7-Zip (7z.exe) not found. Please install 7-Zip or ensure 7z.exe is on PATH."
}

function Get-OmVersionInfo {
    param(
        [Parameter(Mandatory)]
        [System.IO.DirectoryInfo] $Dir
    )
    # Match exact folder names like "OM 15.4.31" or "OM 15.4.1.23"
    $rx = '^OM\s+(\d+(?:\.\d+){2,3})$'
    if ($Dir.Name -match $rx) {
        $verText = $Matches[1]
        try {
            $verObj = [version]$verText
        } catch {
            return $null
        }
        # Filter: major >= 15 only
        if ($verObj.Major -lt 15) { return $null }

        # Identify "unstable" versions (last segment = 9999)
        $parts = $verText -split '\.'
        $isUnstable = ($parts[-1] -eq '9999')

        return [pscustomobject]@{
            Name       = $Dir.Name
            VersionStr = $verText
            VersionObj = $verObj
            FullPath   = $Dir.FullName
            IsUnstable = $isUnstable
        }
    }
    return $null
}

function Get-IncludedFiles {
    param(
        [Parameter(Mandatory)] [string] $VersionFolder
    )
    # Returns a list of FileInfo objects to include,
    # enforcing the omofficeaddin/_universal-only rule.
    $allFiles = Get-ChildItem -Path $VersionFolder -Recurse -File

    $include = foreach ($f in $allFiles) {
        $full = $f.FullName
        # normalize slashes to backslashes, case-insensitive checks
        $relToRoot = $full.Substring($VersionFolder.Length).TrimStart('\','/')
        $norm = $relToRoot -replace '/', '\'
        $lower = $norm.ToLowerInvariant()

        # If path passes through om-apps\omofficeaddin\..., only include when it's under _universal
        if ($lower -match '(^|\\)om-apps\\omofficeaddin(\\|$)') {
            if ($lower -match '(^|\\)om-apps\\omofficeaddin\\_universal(\\|$)') {
                $f
            } else {
                # skip any other subfolders under omofficeaddin
                continue
            }
        } else {
            # outside omofficeaddin: include normally
            $f
        }
    }

    # Sort deterministically by relative path (for manifest/hash stability)
    $include | Sort-Object FullName
}

function New-InventoryHash {
    param(
        [Parameter(Mandatory)] [System.IO.FileInfo[]] $Files,
        [Parameter(Mandatory)] [string] $VersionFolder,
        [Parameter(Mandatory)] [string] $FolderName # e.g., "OM 15.4.31"
    )
    # Build lines of "FolderName\relativePath | length | lastWriteTicksUtc"
    $lines = foreach ($f in $Files) {
        $rel = $f.FullName.Substring($VersionFolder.Length).TrimStart('\','/')
        $rel = $rel -replace '/', '\'
        $relWithTop = Join-Path $FolderName $rel
        $len = $f.Length
        $ticks = ($f.LastWriteTimeUtc).Ticks
        "$relWithTop|$len|$ticks"
    }

    $sb = New-Object System.Text.StringBuilder
    foreach ($line in ($lines | Sort-Object)) {
        [void]$sb.AppendLine($line)
    }

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $sha = New-Object System.Security.Cryptography.SHA256Managed
    try {
        $hashBytes = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    ($hashBytes | ForEach-Object { $_.ToString('x2') }) -join ''
}

function Write-ManifestFiles {
    param(
        [Parameter(Mandatory)] [System.IO.FileInfo[]] $Files,
        [Parameter(Mandatory)] [string] $VersionFolder,
        [Parameter(Mandatory)] [string] $FolderName,  # e.g., "OM 15.4.31"
        [Parameter(Mandatory)] [string] $BaseOutPath  # without extension, e.g., ...\OM15.4.31
    )

    $manifestPathTmp = "$BaseOutPath.manifest.json.tmp"
    $hashPathTmp     = "$BaseOutPath.sha256.tmp"
    $manifestPath    = "$BaseOutPath.manifest.json"
    $hashPath        = "$BaseOutPath.sha256"

    $fileEntries = foreach ($f in $Files) {
        $rel = $f.FullName.Substring($VersionFolder.Length).TrimStart('\','/')
        $rel = $rel -replace '/', '\'
        [pscustomobject]@{
            Path               = (Join-Path $FolderName $rel)
            Length             = $f.Length
            LastWriteTimeUtc   = $f.LastWriteTimeUtc.ToString('o')
        }
    }

    $manifest = [pscustomobject]@{
        SourceFolder  = $VersionFolder
        TopFolderName = $FolderName
        FileCount     = $fileEntries.Count
        TotalBytes    = ($Files | Measure-Object Length -Sum).Sum
        Files         = $fileEntries
        GeneratedUtc  = (Get-Date).ToUniversalTime().ToString('o')
    }

    $json = $manifest | ConvertTo-Json -Depth 6
    $json | Out-File -FilePath $manifestPathTmp -Encoding UTF8 -Force

    $hash = New-InventoryHash -Files $Files -VersionFolder $VersionFolder -FolderName $FolderName
    $hash | Out-File -FilePath $hashPathTmp -Encoding ASCII -Force

    Move-Item -Path $manifestPathTmp -Destination $manifestPath -Force
    Move-Item -Path $hashPathTmp     -Destination $hashPath     -Force

    return $hash
}

function Read-ExistingHash {
    param(
        [Parameter(Mandatory)] [string] $BaseOutPath
    )
    $hashPath = "$BaseOutPath.sha256"
    if (Test-Path $hashPath) {
        (Get-Content -LiteralPath $hashPath -Raw).Trim()
    } else {
        $null
    }
}

function Invoke-7ZipArchive {
    param(
        [Parameter(Mandatory)] [string] $SevenZipPath,
        [Parameter(Mandatory)] [string] $WorkingDir,    # parent of the OM folder
        [Parameter(Mandatory)] [string] $ListFilePath,  # absolute path; entries are RELATIVE to $WorkingDir
        [Parameter(Mandatory)] [string] $ZipOutPathTmp  # final UNC temp: \\...\OMxx.zip.tmp
    )

    if (-not (Test-Path -LiteralPath $ListFilePath)) {
        throw "Listfile not found: $ListFilePath"
    }
    $quotedListArg = '@"' + $ListFilePath + '"'

    # Build to a LOCAL temp first to avoid UNC weirdness, then move
    $localTmpDir  = Join-Path $env:TEMP 'OMZipLocalTmp'
    if (-not (Test-Path $localTmpDir)) { New-Item -ItemType Directory -Path $localTmpDir | Out-Null }
    $localTmpPath = Join-Path $localTmpDir ([System.Guid]::NewGuid().ToString() + '.zip')

    # Compose a single argument string (7z is picky about quoting)
    $argLine = @(
        'a',
        '-tzip',
        ('-mx=' + [string]$CompressionLevel),
        '-y',
        '-bd',                  # no progress bar -> less stderr noise
        ('"' + $localTmpPath + '"'),
        $quotedListArg
    ) -join ' '

    $so = Join-Path $localTmpDir ([System.Guid]::NewGuid().ToString() + '.out.txt')
    $se = Join-Path $localTmpDir ([System.Guid]::NewGuid().ToString() + '.err.txt')

    $p = Start-Process -FilePath $SevenZipPath `
                       -ArgumentList $argLine `
                       -WorkingDirectory $WorkingDir `
                       -NoNewWindow -Wait -PassThru `
                       -RedirectStandardOutput $so `
                       -RedirectStandardError  $se

    $exit = $p.ExitCode
    $stdout = ''
    if (Test-Path -LiteralPath $so) { $stdout = Get-Content -LiteralPath $so -Raw }
    $stderr = ''
    if (Test-Path -LiteralPath $se) { $stderr = Get-Content -LiteralPath $se -Raw }

    # Clean up temp logs
    if (Test-Path -LiteralPath $so) { Remove-Item -LiteralPath $so -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $se) { Remove-Item -LiteralPath $se -Force -ErrorAction SilentlyContinue }

    if ($exit -ne 0) {
        throw ("7-Zip failed (exit {0}). STDERR:`n{1}`nSTDOUT:`n{2}" -f $exit, $stderr.Trim(), $stdout.Trim())
    }
    if (-not (Test-Path -LiteralPath $localTmpPath)) {
        throw ("7-Zip reported success but archive not found: {0}" -f $localTmpPath)
    }

    if (Test-Path -LiteralPath $ZipOutPathTmp) { Remove-Item -LiteralPath $ZipOutPathTmp -Force }
    Move-Item -LiteralPath $localTmpPath -Destination $ZipOutPathTmp -Force
}



function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

# --- Main ---

$sevenZip = Get-SevenZipPath
Write-Host "Using 7-Zip at: $sevenZip"

Ensure-Folder -Path $PublishRoot

# Collect source versions (major >= 15; 3–4 parts)
$versions =
    Get-ChildItem -Path $SourceRoot -Directory |
    ForEach-Object { Get-OmVersionInfo -Dir $_ } |
    Where-Object { $_ -ne $null } |
    Sort-Object -Property @{Expression = 'VersionObj'; Descending = $true}

if (-not $versions) {
    Write-Warning "No matching OM versions (>= 15) found under $SourceRoot."
    return
}

# Build a map of expected output base names (OM15.4.31)
$expectedNames = @{}
foreach ($v in $versions) {
    $zipBaseName = 'OM' + $v.VersionStr  # no spaces/underscores
    $expectedNames[$zipBaseName] = $true
}

# Process each version: build zip if changed
foreach ($v in $versions) {
    $folderName = $v.Name                  # "OM 15.4.31"
    $versionText = $v.VersionStr           # "15.4.31"
    $zipBaseName = 'OM' + $versionText     # "OM15.4.31"
    $zipPath     = Join-Path $PublishRoot ($zipBaseName + '.zip')
    $zipTmp      = $zipPath + '.tmp'
    $baseOut     = Join-Path $PublishRoot $zipBaseName

    Write-Host "-----"
    Write-Host "Preparing $zipBaseName from '$folderName' ..."

    $files = Get-IncludedFiles -VersionFolder $v.FullPath

    if (-not $files -or $files.Count -eq 0) {
        Write-Warning "No files to include for $folderName (after filters). Skipping."
        # still write empty manifest & hash to reflect that nothing included
        Write-ManifestFiles -Files @() -VersionFolder $v.FullPath -FolderName $folderName -BaseOutPath $baseOut | Out-Null
        continue
    }

    $newHash = New-InventoryHash -Files $files -VersionFolder $v.FullPath -FolderName $folderName
    $oldHash = Read-ExistingHash -BaseOutPath $baseOut

    if ($oldHash -and ($oldHash -eq $newHash) -and (Test-Path $zipPath)) {
        Write-Host "No changes detected for $zipBaseName. Skipping rebuild."
        continue
    }

    # Build @listfile with relative paths prefixed by the top-level folder name
    $parentDir = Split-Path -Path $v.FullPath -Parent
    $listFile = [System.IO.Path]::GetTempFileName()

    try {
        # Use the system ANSI codepage (preserves umlauts without BOM)
        $enc = [System.Text.Encoding]::Default
        $sw = New-Object System.IO.StreamWriter($listFile, $false, $enc)
        try {
            foreach ($f in $files) {
                $rel = $f.FullName.Substring($v.FullPath.Length).TrimStart('\', '/')
                $rel = $rel -replace '/', '\'
                $entry = Join-Path $folderName $rel   # includes the space in "OM 15.x.x" — OK
                $sw.Write($entry)
                $sw.Write("`r`n")                     # CRLF
            }
        }
        finally {
            $sw.Dispose()
        }

        if (Test-Path $zipTmp) { Remove-Item -LiteralPath $zipTmp -Force }

        Write-Host "Creating zip: $zipPath"
        Invoke-7ZipArchive -SevenZipPath $sevenZip -WorkingDir $parentDir -ListFilePath $listFile -ZipOutPathTmp $zipTmp

        # Wait briefly for the UNC tmp to appear, then rename
        for ($i = 0; $i -lt 20 -and -not (Test-Path -LiteralPath $zipTmp); $i++) { Start-Sleep -Milliseconds 500 }
        if (-not (Test-Path -LiteralPath $zipTmp)) { throw "Expected temp archive not found after 7-Zip: $zipTmp" }

        if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
        Move-Item -LiteralPath $zipTmp -Destination $zipPath -Force

        $null = Write-ManifestFiles -Files $files -VersionFolder $v.FullPath -FolderName $folderName -BaseOutPath $baseOut

    }
    finally {
        if (Test-Path $listFile) { Remove-Item -LiteralPath $listFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $zipTmp) { Remove-Item -LiteralPath $zipTmp   -Force -ErrorAction SilentlyContinue }
    }

}

# Determine latest stable (exclude *.9999)
$latestStable = $versions | Where-Object { -not $_.IsUnstable } | Select-Object -First 1
if ($latestStable) {
    $latestBase = 'OM' + $latestStable.VersionStr
    $latestZip  = Join-Path $PublishRoot ($latestBase + '.zip')
    $aliasZip   = Join-Path $PublishRoot 'OM_latest.zip'

    if (Test-Path $latestZip) {
        Write-Host "Updating OM_latest.zip -> $latestBase.zip"
        # Copy as alias (faster than re-building)
        if (Test-Path $aliasZip) { Remove-Item -LiteralPath $aliasZip -Force }
        Copy-Item -LiteralPath $latestZip -Destination $aliasZip -Force
    } else {
        Write-Warning "Expected latest zip not found: $latestZip"
    }
} else {
    Write-Warning "No stable (non-*.9999) versions found for OM_latest.zip."
}

# Cleanup: remove orphaned zips/sidecars in $PublishRoot (except OM_latest.zip)
Write-Host "Running cleanup in $PublishRoot ..."
$zipRegex = '^OM(\d+(?:\.\d+){2,3})\.zip$'
$publishZips = Get-ChildItem -Path $PublishRoot -File -Filter 'OM*.zip' | Where-Object { $_.Name -ne 'OM_latest.zip' }

foreach ($z in $publishZips) {
    if ($z.Name -match $zipRegex) {
        $basename = [System.IO.Path]::GetFileNameWithoutExtension($z.Name)
        if (-not $expectedNames.ContainsKey($basename)) {
            Write-Host "Removing orphaned: $($z.FullName)"
            $base = Join-Path $PublishRoot $basename
            Remove-Item -LiteralPath $z.FullName -Force -ErrorAction SilentlyContinue
            foreach ($ext in @('.manifest.json', '.sha256', '.tmp')) {
                $p = $base + $ext
                if (Test-Path $p) {
                    Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

Write-Host "All done."
