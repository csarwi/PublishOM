function Get-NewestRelease(){

    $path = '\\creativ.local\data_creativ$\DATA\DefinitiveSoftwareLibrary'

    # Capture versions right after "OM " with 3–4 numeric parts (e.g., 15.4.18 or 12.11.2.38)
    $pattern = '^OM\s+(\d+(?:\.\d+){2,3})\b'

    $latest = Get-ChildItem -Path $path -Directory |
        ForEach-Object {
            if ($_.Name -match $pattern) {
                $ver = $Matches[1]
                $lastPart = $ver.Split('.')[-1]
                if ($lastPart -ne '9999') {
                    [pscustomobject]@{
                        Name    = $_.Name
                        Version = [version]$ver   # Proper numeric sort
                    }
                }
            }
        } |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($latest) {
        $fullpath = Join-Path $path $latest.Name
        return $fullpath
    } else {
        Write-Error "No stable OM release found."
        return $null
    }
}

$latestReleasePath = Get-NewestRelease
Write-Host "Newest OM folder: $latestReleasePath"

if ($latestReleasePath) {

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $destRoot = '\\creativ.local\data_creativ$\DATA\_om_latest'
    $zipPath  = Join-Path $destRoot 'OM_latest.zip'

    # Ensure destination exists
    if (-not (Test-Path $destRoot)) {
        New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
    }

    # Remove existing zip if present
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }

    # Create the archive with the folder itself as the top-level entry
    # includeBaseDirectory:$true -> zip contains "OM 15.x.x\" at the root
    Write-Host "zipping file ..."
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $latestReleasePath,
        $zipPath,
        [System.IO.Compression.CompressionLevel]::Fastest,
        $true
    )

    Write-Host "Created archive: $zipPath"

}