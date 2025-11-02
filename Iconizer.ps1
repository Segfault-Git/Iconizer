function Timer {
    param(
        [switch]$start,
        [switch]$end
    )
    if ($start){
        $global:timer = [Diagnostics.Stopwatch]::StartNew()
    }
    if ($end){
        $global:timer.Stop()
        $timeRound = [Math]::Round(($global:timer.Elapsed.TotalSeconds), 2)
        $global:timer.Reset()
        Write-Host "----------`nTask completed in $timeRound`s" -ForegroundColor Cyan
    }
}

function SelectPath {
    param(
        [switch]$files
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $Topmost = New-Object System.Windows.Forms.Form
    $Topmost.TopMost = $True
    $Topmost.MinimizeBox = $True
    
    if ($files){
        $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.RestoreDirectory = $True
        $OpenFileDialog.Title = 'Select an EXE File'
        $OpenFileDialog.Filter = 'Executable files (*.exe)|*.exe'
        if (($OpenFileDialog.ShowDialog($Topmost) -eq 'OK')) {
            $file = $OpenFileDialog.FileName
        } else {
            $file = $null
        }
    } else {
        $OpenFolderDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
        $OpenFolderDialog.Description = 'Select a folder'
        $OpenFolderDialog.Rootfolder = 'MyComputer'
        $OpenFolderDialog.ShowNewFolderButton = $false
        if ($OpenFolderDialog.ShowDialog($Topmost) -eq 'OK') {
            $directory = $OpenFolderDialog.SelectedPath
        } else {
            $directory = $null
        }
    }
    
    $Topmost.Close()
    $Topmost.Dispose()
    
    if ($file){
        return $file
    } elseif ($directory){
        return $directory
    } else {
        return $null
    }
}

function Import-Type-Pull {
    try {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class IconExtractor
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, uint dwFlags);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool FreeLibrary(IntPtr hModule);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr FindResourceW(IntPtr hModule, string lpName, IntPtr lpType);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr LockResource(IntPtr hResData);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumResourceNames(IntPtr hModule, IntPtr lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);
    
    public delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);
    
    public const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
    public const int RT_GROUP_ICON = 14;
    public const int RT_ICON = 3;
    
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, uint nIconIndex);
    
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
    
    // Helper method to check if a pointer represents an integer resource
    public static bool IS_INTRESOURCE(IntPtr ptr)
    {
        return ((ulong)ptr) >> 16 == 0;
    }
    
    // Method to get the main icon index (simplified approach)
    public static int GetMainIconIndex(string filePath)
    {
        try
        {
            // Get total icon count
            uint iconCount = ExtractIconEx(filePath, -1, null, null, 0);
            return iconCount > 0 ? 0 : -1; // Return 0 for first icon, -1 if no icons
        }
        catch
        {
            return -1;
        }
    }
}
"@

    # Add .NET assemblies for image processing
    Add-Type -AssemblyName System.Drawing

    } catch [System.Exception] {
        Write-Host "An unexpected error occurred in Import-Type-Pull: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

function Get-IconsByGroup-Pull {
    [CmdletBinding()]
    param(
        [string]$FilePath,
        [int]$index = 1,
        [string]$OutputDir = ".",
        [switch]$all,
        [switch]$info,
        [switch]$png
    )
    
    if (-not (Test-Path $FilePath)) {
        Write-Host "File not found: $FilePath" -ForegroundColor Red
        return
    }
    
    if (-not $info -and -not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    
    if ($info) {
        $all = $true
    }
    
    $ICO_name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    
    $hModule = [IconExtractor]::LoadLibraryEx($FilePath, [IntPtr]::Zero, [IconExtractor]::LOAD_LIBRARY_AS_DATAFILE)
    
    if ($hModule -eq [IntPtr]::Zero) {
        Write-Host "Failed to load IconExtractor!" -ForegroundColor Red
        return
    }
    
    try {
        $script:currentGroup = 0
        $script:targetGroup = $index
        $script:extractAll = $all
        $script:totalExtracted = 0
        $script:processedGroups = @()
        $script:resourcesNames = @()
        
        if ($script:extractAll) {
            Write-Host 'Analyzing all icon groups' -ForegroundColor Yellow
        } else {
            Write-Host "Analyzing group " -NoNewline -ForegroundColor DarkGray
            Write-Host "#$index" -ForegroundColor Yellow
        }
        
        $callback = {
            param($hMod, $lpType, $lpName, $lParam)
            
            $script:currentGroup++
            
            # Variables to track the best icon from current group
            $currentGroupLargestIcon = $null
            $currentGroupLargestSize = 0
            
            if (-not $script:extractAll -and $script:currentGroup -ne $script:targetGroup) {
                return $true
            }
            
            # Determine resource name/ID
            if ([IconExtractor]::IS_INTRESOURCE($lpName)) {
                $resourceId = [int]$lpName
                $resourceName = "ID_$resourceId"
            } else {
                try {
                    $stringName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($lpName)
                    $resourceName = if ([string]::IsNullOrWhiteSpace($stringName)) { "UNNAMED" } else { $stringName }
                } catch {
                    $resourceName = "ERROR_READING_NAME"
                }
            }
            
            Write-Host "Extracting group " -ForegroundColor DarkGray -NoNewline
            Write-Host "#$script:currentGroup ($resourceName)" -ForegroundColor Green
            # Load and analyze icon group resource
            $hResInfo = [IntPtr]::Zero
            
            if ([IconExtractor]::IS_INTRESOURCE($lpName)) {
                $hResInfo = [IconExtractor]::FindResource($hMod, $lpName, [IntPtr][IconExtractor]::RT_GROUP_ICON)
            } else {
                try {
                    $stringName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($lpName)
                    if (-not [string]::IsNullOrEmpty($stringName)) {
                        $hResInfo = [IconExtractor]::FindResourceW($hMod, $stringName, [IntPtr][IconExtractor]::RT_GROUP_ICON)
                    }
                } catch {
                    Write-Host "Error processing string name" -ForegroundColor Red
                }
            }
            
            $groupExtracted = $false
            
            if ($hResInfo -ne [IntPtr]::Zero) {
                $hResData = [IconExtractor]::LoadResource($hMod, $hResInfo)
                if ($hResData -ne [IntPtr]::Zero) {
                    $pData = [IconExtractor]::LockResource($hResData)
                    $size = [IconExtractor]::SizeofResource($hMod, $hResInfo)
                    
                    if ($pData -ne [IntPtr]::Zero -and $size -gt 6) {
                        # Read group resource data
                        $iconDir = New-Object byte[] $size
                        [System.Runtime.InteropServices.Marshal]::Copy($pData, $iconDir, 0, $size)
                        
                        # Parse group icon header
                        $iconCount = [BitConverter]::ToUInt16($iconDir, 4)
                        
                        Write-Host "Icons found in group: " -ForegroundColor DarkGray -NoNewline
                        Write-Host "$iconCount" -ForegroundColor Cyan
                        
                        # Create ICO file
                        if ($all){
                            $icoPath = Join-Path $OutputDir "${ICO_name}_Group_${script:currentGroup}_${resourceName}.ico"
                        } else {
                            $icoPath = Join-Path $OutputDir "$ICO_name.ico"
                        }
                        
                        $icoData = @()
                        
                        # ICO file header
                        $icoData += @(0, 0)  # Reserved
                        $icoData += @(1, 0)  # Type (1 = ICO)
                        $icoData += [BitConverter]::GetBytes([uint16]$iconCount)  # Count
                        
                        $iconDataArray = @()
                        $currentOffset = 6 + ($iconCount * 16)  # Header + directory
                        
                        # Process each icon in group
                        for ($i = 0; $i -lt $iconCount; $i++) {
                            $offset = 6 + ($i * 14)
                            if ($offset + 13 -lt $iconDir.Length) {
                                $width = $iconDir[$offset]
                                $height = $iconDir[$offset + 1]
                                $colorCount = $iconDir[$offset + 2]
                                $reserved2 = $iconDir[$offset + 3]
                                $planes = [BitConverter]::ToUInt16($iconDir, $offset + 4)
                                $bitCount = [BitConverter]::ToUInt16($iconDir, $offset + 6)
                                $iconId = [BitConverter]::ToUInt16($iconDir, $offset + 12)
                                
                                # Load individual icon data
                                $hIconRes = [IconExtractor]::FindResource($hMod, [IntPtr]$iconId, [IntPtr][IconExtractor]::RT_ICON)
                                if ($hIconRes -ne [IntPtr]::Zero) {
                                    $hIconData = [IconExtractor]::LoadResource($hMod, $hIconRes)
                                    if ($hIconData -ne [IntPtr]::Zero) {
                                        $pIconData = [IconExtractor]::LockResource($hIconData)
                                        $iconSize = [IconExtractor]::SizeofResource($hMod, $hIconRes)
                                        
                                        if ($pIconData -ne [IntPtr]::Zero -and $iconSize -gt 0) {
                                            $iconBytes = New-Object byte[] $iconSize
                                            [System.Runtime.InteropServices.Marshal]::Copy($pIconData, $iconBytes, 0, $iconSize)
                                            
                                            # Check if this is the largest icon in current group for PNG extraction
                                            if ($png) {
                                                $actualWidth = if ($width -eq 0) { 256 } else { $width }
                                                $actualHeight = if ($height -eq 0) { 256 } else { $height }
                                                $iconPixelSize = $actualWidth * $actualHeight
                                                
                                                if ($iconPixelSize -gt $currentGroupLargestSize) {
                                                    $currentGroupLargestSize = $iconPixelSize
                                                    $currentGroupLargestIcon = @{
                                                        Width = $actualWidth
                                                        Height = $actualHeight
                                                        Data = $iconBytes
                                                        Group = $script:currentGroup
                                                        ResourceName = $resourceName
                                                    }
                                                }
                                            }
                                            
                                            # Add icon directory to ICO file
                                            $icoData += @($width, $height, $colorCount, $reserved2)
                                            $icoData += [BitConverter]::GetBytes($planes)
                                            $icoData += [BitConverter]::GetBytes($bitCount)
                                            $icoData += [BitConverter]::GetBytes([uint32]$iconSize)
                                            $icoData += [BitConverter]::GetBytes([uint32]$currentOffset)
                                            
                                            $iconDataArray += ,$iconBytes
                                            $currentOffset += $iconSize
                                            if ($bitCount -eq 0) { $bitCount = 32 }
                                            if ($width -eq 0) { $width = 256 }
                                            if ($height -eq 0) { $height = 256 }
                                            if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "  Icon $($i+1): ${width}x${height}, $bitCount bit, $iconSize bytes" -ForegroundColor Gray }
                                        }
                                    }
                                }
                            }
                        }
                        
                        # Write ICO file
                        if ($iconDataArray.Count -gt 0) {
                            if (!($info) -and !($png)){
                                $allData = @()
                                $allData += $icoData
                                foreach ($iconBytes in $iconDataArray) {
                                    $allData += $iconBytes
                                }
                                [System.IO.File]::WriteAllBytes($icoPath, $allData)
                                Write-Host "Saved: " -NoNewline -ForegroundColor DarkGray
                                Write-Host "$icoPath" -ForegroundColor Green
                            }
                            $script:totalExtracted++
                            $script:processedGroups += $script:currentGroup
                            $script:resourcesNames += $resourceName
                            $groupExtracted = $true
                        }
                    }
                }
            }
            
            if (-not $groupExtracted) {
                Write-Host "Failed to extract group #$script:currentGroup" -ForegroundColor Red
            }
            
            # Extract largest icon from current group as PNG if requested
            if ($png -and $currentGroupLargestIcon) {
                $pngFileName = if ($script:extractAll) {
                    "${ICO_name}_Group_${script:currentGroup}_${resourceName}.png"
                } else {
                    "${ICO_name}.png"
                }
                $pngPath = Join-Path $OutputDir $pngFileName
                
                Convert-IconToPNG -IconData $currentGroupLargestIcon -PngPath $pngPath -ICO_name $ICO_name
            }
            
            return $script:extractAll
        }
        
        $callbackDelegate = [IconExtractor+EnumResNameProc]$callback
        [IconExtractor]::EnumResourceNames($hModule, [IntPtr][IconExtractor]::RT_GROUP_ICON, $callbackDelegate, [IntPtr]::Zero) | Out-Null
        
        if ($script:totalExtracted -eq 0) {
            if ($script:extractAll) {
                Write-Host "No icon groups found or failed to extract any groups" -ForegroundColor Red
            } else {
                Write-Host "Group #$index not found or failed to extract" -ForegroundColor Red
            }
        } else {
            Write-Host "Total groups extracted: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$script:totalExtracted" -ForegroundColor Cyan
            Write-Host "Processed groups: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$($script:resourcesNames -join ', ')" -ForegroundColor Cyan
        }
    }
    finally {
        [IconExtractor]::FreeLibrary($hModule) | Out-Null
    }
}

function Convert-IconToPNG {
    [CmdletBinding()]
    param(
        [hashtable]$IconData,
        [string]$PngPath,
        [string]$ICO_name
    )
    
    try {
        Write-Host "Extracting icon as PNG: $($IconData.Width)x$($IconData.Height) from Group $($IconData.Group)" -ForegroundColor Cyan
        
        $conversionSuccess = $false
        
        # Method 1: Try direct PNG extraction if the icon data is already PNG
        if ($IconData.Data.Length -gt 8) {
            $pngSignature = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
            $isPng = $true
            for ($i = 0; $i -lt 8; $i++) {
                if ($IconData.Data[$i] -ne $pngSignature[$i]) {
                    $isPng = $false
                    break
                }
            }
            
            if ($isPng) {
                Write-Host "Icon data is already PNG format, saving directly..." -ForegroundColor Cyan
                [System.IO.File]::WriteAllBytes($PngPath, $IconData.Data)
                $conversionSuccess = $true
            }
        }
        
        # Method 2: Try .NET conversion if not PNG
        if (-not $conversionSuccess) {
            try {
                # Create properly formatted ICO file
                $tempIcoData = @()
                $tempIcoData += @(0, 0, 1, 0, 1, 0)  # ICO header for single icon
                
                # Icon directory entry (16 bytes)
                $width = if ($IconData.Width -eq 256) { 0 } else { $IconData.Width }
                $height = if ($IconData.Height -eq 256) { 0 } else { $IconData.Height }
                $tempIcoData += @($width, $height, 0, 0)  # width, height, colors, reserved
                $tempIcoData += @(1, 0, 32, 0)  # planes, bitcount
                $tempIcoData += [BitConverter]::GetBytes([uint32]$IconData.Data.Length)  # size
                $tempIcoData += [BitConverter]::GetBytes([uint32]22)  # offset
                
                # Add icon data
                $tempIcoData += $IconData.Data
                
                # Save temporary ICO file
                $tempIcoPath = Join-Path $env:TEMP "temp_largest_icon.ico"
                [System.IO.File]::WriteAllBytes($tempIcoPath, $tempIcoData)
                
                Write-Host "Saving as png..." -ForegroundColor Cyan
                
                # Try multiple .NET approaches
                try {
                    # Approach 1: Direct Icon loading
                    Write-Host "Using direct Icon loading..." -ForegroundColor Cyan
                    $icon = [System.Drawing.Icon]::new($tempIcoPath)
                    $bitmap = $icon.ToBitmap()
                    $bitmap.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $bitmap.Dispose()
                    $icon.Dispose()
                    $conversionSuccess = $true
                } catch {
                    Write-Host "Direct Icon loading failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    
                    # Approach 2: Try extracting from file stream
                    try {
                        Write-Host "Using FileStream approach..." -ForegroundColor Cyan
                        $fileStream = [System.IO.FileStream]::new($tempIcoPath, [System.IO.FileMode]::Open)
                        $icon = [System.Drawing.Icon]::new($fileStream)
                        $bitmap = $icon.ToBitmap()
                        $bitmap.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                        $bitmap.Dispose()
                        $icon.Dispose()
                        $fileStream.Close()
                        $conversionSuccess = $true
                    } catch {
                        Write-Host "FileStream approach failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                
                # Cleanup temp file
                Remove-Item $tempIcoPath -Force -ErrorAction SilentlyContinue
                
            } catch {
                Write-Host ".NET conversion failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # Method 3: Save raw icon data as fallback
        if (-not $conversionSuccess) {
            Write-Host "Saving raw icon data as fallback..." -ForegroundColor Yellow
            $rawPath = $PngPath -replace '\.png$', '_raw.bin'
            [System.IO.File]::WriteAllBytes($rawPath, $IconData.Data)
            Write-Host "Raw icon data saved: $rawPath" -ForegroundColor Yellow
            Write-Host "You can try converting this file manually with image editing software" -ForegroundColor Yellow
        }
        
        if ($conversionSuccess) {
            Write-Host "Icon saved as PNG: $PngPath" -ForegroundColor Green
            return $true
        } else {
            Write-Host "PNG conversion failed, but raw data was saved for manual conversion" -ForegroundColor Yellow
            return $false
        }
        
    } catch {
        Write-Host "Error during PNG extraction: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Find-Candidates {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$path,
        [parameter(Mandatory = $true)]
        [int]$search_depth,
        [ValidateSet('ico', 'exe')]
        $priority
    )
    
    $folder = Get-Item -LiteralPath $path
    
    [string]$full_path_folder = $folder.FullName
    
    $iconsFilesList = Get-ChildItem -LiteralPath "$full_path_folder" -Recurse -Filter "*.ico" -Depth $search_depth -File
    $exesFilesList = Get-ChildItem -LiteralPath "$full_path_folder" -Recurse -Filter "*.exe" -Depth $search_depth -File
    $allFiles = @($exesFilesList) + @($iconsFilesList)
    
    [string]$name_folder = ($folder.Name).ToLower()
    $name_folder = $name_folder -replace $pattern_regex, '' -replace "$pattern_regex_digits", '' -replace "$pattern_regex_symbols", ''
    
    if ($allFiles) {
        if ($VerbosePreference -ne 'SilentlyContinue') {
            Write-Host "Found: " -ForegroundColor DarkGray
            Write-Host "$($allFiles -join "`n")" -ForegroundColor DarkGray
        }
        
        $candidates = @()
        
        $folderWords = $name_folder.Trim().Split(' ')
        $fullFolderName = $name_folder -replace '\s+', ''
        
        foreach ($file in $allFiles) {
            $FileName = ($file.BaseName).ToLower() -replace $pattern_regex_symbols, '' -replace $pattern_regex, '' -replace $pattern_regex_digits, ''
            $score = 0
            
            if ($FileName -eq $fullFolderName) {
                $score = 1000
            } elseif ($FileName.Contains($fullFolderName) -or $fullFolderName.Contains($FileName)) {
                $score = 500 + ($FileName.Length - [Math]::Abs($FileName.Length - $fullFolderName.Length))
            } else {
                foreach ($word in $folderWords) {
                    if ($word.Length -gt 2) {
                        if ($FileName -eq $word) {
                            $score += 200
                        } elseif ($FileName.Contains($word)) {
                            $score += 100
                        } elseif ($word.Contains($FileName) -and $FileName.Length -gt 3) {
                            $score += 50
                        }
                    }
                }
                
                if ($score -gt 0) {
                    $lengthDiff = [Math]::Abs($FileName.Length - $fullFolderName.Length)
                    $score += [Math]::Max(0, 20 - $lengthDiff)
                }
            }
            
            if ($priority -and $file.Extension.ToLower() -eq ".$priority") {
                $score += 400
                if ($VerbosePreference -ne 'SilentlyContinue'){
                    Write-Host "Priority bonus applied to: $($file.Name)"
                }
            }
            
            if ($score -gt 0) {
                # Bonus for short names
                $score += [Math]::Max(0, 50 - $FileName.Length)
                $candidates += [PSCustomObject]@{
                    File  = $file
                    Score = $score
                    Name  = $FileName
                }
                if ($VerbosePreference -ne 'SilentlyContinue'){
                    Write-Host "Candidate: $($file.Name) | Score: $score" -ForegroundColor DarkGray
                }
            }
        }
        
        if ($candidates.Count -gt 0) {
            $bestCandidate = $candidates | Sort-Object Score -Descending | Select-Object -First 1
            $Files = $bestCandidate.File

            if ($VerbosePreference -ne 'SilentlyContinue') {
                Write-Host "Best candidate: " -NoNewline -ForegroundColor DarkGray
                Write-Host "$($Files.Name) " -NoNewline -ForegroundColor Cyan
                Write-Host "with score " -NoNewline -ForegroundColor DarkGray
                Write-Host "$($bestCandidate.Score)" -ForegroundColor Cyan
            }
        } else {
            $Files = $allFiles | Select-Object -First 1
            Write-Host "No matches found, using first exe: `'$($Files.Name)`'" -ForegroundColor Yellow
        }
    } else {
        if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host ".$priority candidates not found" }
    }
    return $Files
}

function Test-ForbiddenFolder {
    [CmdletBinding()]
    param (
        [string]$Path,
        [string[]]$ForbiddenFolders
    )

    if (-not $ForbiddenFolders) { return $false }

    $normalizedPath = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue)
    if (-not $normalizedPath) { return $false }
    $normalizedPath = $normalizedPath.Path

    foreach ($forbidden in $ForbiddenFolders) {
        if (-not $forbidden) { continue }

        # Check for full path (starts with drive or UNC)
        if ($forbidden -match '^[a-zA-Z]:\\' -or $forbidden.StartsWith('\\')) {
            if ($normalizedPath -like "$forbidden*") {
                if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Path '$normalizedPath' blocked by full path filter '$forbidden'" }
                return $true
            }
        }
        # Checking patterns with stars
        elseif ($forbidden.Contains('*')) {
            # *pattern* - в любом месте пути
            if ($forbidden.StartsWith('*') -and $forbidden.EndsWith('*')) {
                $pattern = $forbidden.Trim('*')
                if ($normalizedPath -like "*$pattern*") {
                    if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Path '$normalizedPath' blocked by wildcard filter '$forbidden' (anywhere in path)" }
                    return $true
                }
            }
            # *pattern - the last folder in the path must end with pattern
            elseif ($forbidden.StartsWith('*')) {
                $pattern = $forbidden.TrimStart('*')
                $lastFolder = Split-Path -Leaf $normalizedPath
                if ($lastFolder -like "*$pattern") {
                    if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Path '$normalizedPath' blocked by end-pattern filter '$forbidden' (last folder: '$lastFolder')" }
                    return $true
                }
            }
            # pattern* - regular wildcard
            else {
                if ($normalizedPath -like "*$forbidden*") {
                    if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Path '$normalizedPath' blocked by wildcard filter '$forbidden'" }
                    return $true
                }
            }
        }
        # Exact match only for the last folder in the path
        else {
            $lastFolder = Split-Path -Leaf $normalizedPath
            if ($lastFolder -eq $forbidden) {
                if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Path '$normalizedPath' blocked by exact match filter '$forbidden' (last folder: '$lastFolder')" }
                return $true
            }
        }
    }

    return $false
}

function Restart-ExplorerAsUser {
    $explorerProcesses = Get-Process -Name explorer -ErrorAction SilentlyContinue
    
    if ($explorerProcesses) {
        foreach ($process in $explorerProcesses) {
            try {
                $process | Stop-Process -Force
                Wait-Process -Id $process.Id -Timeout 5 -ErrorAction SilentlyContinue
            } catch {
                Write-Verbose "Process $($process.Id) already terminated or timeout reached"
            }
        }
    }
    
    cmd /c "start /b explorer.exe"
}

function Logging {
    [CmdletBinding()]
    param (
        [string]$log_path,
        [switch]$start,
        [switch]$stop
    )
        if ($stop){
            try {
                Stop-Transcript
            }
            catch {
                return
            }
        }
        
        if(!($log_path)){
            $path = Join-Path -Path $env:LOCALAPPDATA -ChildPath "iconizer.log"
        } else {
            if (Test-Path -Path $log_path -PathType Container) {
                $path = Join-Path -Path $log_path -ChildPath "iconizer.log"
            } else {
                Write-Host "Wrong log path. Specify a folder, not a file!" -ForegroundColor Red
                return
            }
        }
        
        if ($start){
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force
            }
            Start-Transcript -Path $path
        }
}

function pull {
    [CmdletBinding()]
    param (
        # regular dir path
        [Alias('d')]
        [string[]]$directory,
        # index inside exe
        [Alias('i')]
        [int]$index = 1,
        # depth of search
        [Alias('sd')]
        [int]$search_depth = 0,
        # log path
        [Alias('l')]
        [string]$log,
        # convert to png image format
        [switch]$png,
        # show only info without ext
        [switch]$info,
        # ext all posible icons in exe (main by default)
        [Alias('a')]
        [switch]$all,
        # open GUI for single file (folder by default)
        [Alias('f')]
        [switch]$file_sw,
        # pause after
        [Alias('p')]
        [switch]$pause
    )
    
    Import-Type-Pull
    
    if (!($directory)) {
        if ($file_sw){
            $directory = SelectPath -files
            $file_from_GUI = $true
        } else {
            $directory = SelectPath
        }
        
        if ($directory){
            $from_GUI = $true
        }
    }
    
    if (!($directory)) {
        Write-Host "`nNo path was selected by the user. Select the path in the GUI or specify it like -d 'full_path_to_files'`n" -ForegroundColor Red
        return
    }
    
    if ($log){
        Logging -log_path $log -start
    }
    
    $ErrorActionPreference = 'Stop'
    
    Timer -start

    Write-Host "`nProcessing:" -ForegroundColor DarkGray
    $directory | ForEach-Object { Write-Host " $($_)" -ForegroundColor DarkBlue }
    
    try {
        foreach ($i in $directory) {
            if (Test-Path -Path $i){
                $item = Get-Item -Path $i
                
                if ($item -is [System.IO.FileInfo]) {
                    if ($item.Extension -eq '.exe') {
                        $resolved_path = @($item)
                    } else {
                        Write-Host "File is not an executable:`n $($item.FullName)" -ForegroundColor Red
                        continue
                    }
                } else {
                    if (($file_from_GUI) -or ($search_depth -eq 0)) {
                        $resolved_path = Get-ChildItem -Path $i -Filter '*.exe'
                    } else {
                        $resolved_path = Get-ChildItem -Path $i -Filter '*.exe' -Recurse -Depth $search_depth
                    }
                }
                
                foreach ($_path in $resolved_path){
                    if ($_path) {
                        Write-Host "--------------`nExtracting icons from:"
                        Write-Host " $($_path.FullName)" -ForegroundColor Green
                        $params = @{
                            FilePath  = $_path.FullName
                            OutputDir = $_path.DirectoryName
                            index     = $index
                        }
                        
                        if ($all) { $params.all   = $true }
                        if ($info) { $params.info  = $true }
                        if ($png) { $params.png   = $true }
                        Get-IconsByGroup-Pull @params
                    } else {
                        Write-Host "No exe files in path:`n $($_path.FullName)" -ForegroundColor Red
                    }
                } #foreach
            } else {
                Write-Host "Path is not exist:`n $($_path.FullName)" -ForegroundColor Red
            }
        } #foreach
    } catch {
        Write-Host "`nError:$_" -ForegroundColor Red
        Write-Host "`n$($_.ScriptStackTrace)`n" -ForegroundColor Red
    }
    
    Timer -end
    
    if ($log){
        Logging -stop
    }
    
    if ($pause){
        pause
    }
}

function apply {
    [CmdletBinding()]
    param (
        # regular dir path
        [Alias('d')]
        [string[]]$directory,
        # priority ico or exe
        [Alias('p')]
        [ValidateSet('ico', 'exe')]
        $priority,
        # filter dirs
        [Alias('f')]
        [string[]]$filter,
        # log path
        [Alias('l')]
        [string]$log,
        # rules for specific folders
        [Alias('r')]
        $rules,
        # search depth for icons
        [Alias('sd')]
        [int]$search_depth = 0,
        # depth of application
        [Alias('ad')]
        [int]$apply_depth = 0,
        # remove desktop.ini
        [Alias('rm')]
        [switch]$remove,
        # DO NOT overwrite existing desktop.ini
        [Alias('nf')]
        [switch]$noforce,
        # pause after
        [switch]$pause
    )
    
    if (!($directory)) { $directory = SelectPath }
    
    if (!($directory)) {
        Write-Host "`nNo path was selected by the user. Select the path in the GUI or specify it like -d 'full_path_to_dir'`n" -ForegroundColor Red
        return
    }
    
    if ($log){
        Logging -log_path $log -start
    }
    
    $ErrorActionPreference = 'Stop'
    
    try {
        $foldersError = @()
        $folders = @()
        
        Timer -start
        
        [string[]]$Filter_main += $filter
        [string[]]$Filter_main += '*DeliveryOptimization*', "$env:Programfiles", "${env:ProgramFiles(x86)}", "$env:windir", '*OneCommander*', '*$RECYCLE.BIN*', '*System Volume Information*'
        
        foreach ($i in $directory) {
            if (Test-Path -Path "$i") {
                $fullPath = (Resolve-Path -LiteralPath "$i").Path
                if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "Adding: $fullPath" }
                if ($apply_depth -eq 0) {
                    $folders += Get-Item -LiteralPath "$i" -ErrorAction SilentlyContinue
                } else {
                    $folders += Get-ChildItem -LiteralPath "$i" -Directory -Depth $($apply_depth - 1) -ErrorAction SilentlyContinue
                }
            } else {
                Write-Host "Folder `'$i`' does not exist" -ForegroundColor Red
                continue
            }
        }
        
        $folders = $folders | Sort-Object FullName -Unique
        
        if ($folders.Count -eq 0) {
            Write-Host "Folders not found or inaccessible" -ForegroundColor Red
            return
        }
        
        if ($filter) {
            Write-Host "`nYour filter list:" -ForegroundColor DarkGray
            foreach ($i in $filter){ Write-Host " $i" -ForegroundColor Yellow }
        }
        
        Write-Host "`nProcessing folders:" -ForegroundColor DarkGray
        $folders | ForEach-Object { Write-Host " $($_.FullName)" -ForegroundColor DarkBlue }
        
        $primaryType = if ($priority -eq 'ico') { 'ico' } else { 'exe' }
        $secondaryType = if ($priority -eq 'ico') { 'exe' } else { 'ico' }
        if (!($remove)){ Write-Host "`nPriority to $primaryType" }
        foreach ($folder in $folders) {
            if ($remove) {
                try {
                    $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse:$($apply_depth -gt 0) -Depth $apply_depth -ErrorAction SilentlyContinue
                    $desktopINI | Remove-Item -Force
                } catch {
                    Write-Host 'Access to the path is denied. Cant proseed with desktop.ini file. Skiping...' -ForegroundColor Red
                    Write-Host "$($folder.FullName)"
                    continue
                }
                continue
            }
            
            $shouldSkip = Test-ForbiddenFolder -Path $folder.FullName -ForbiddenFolders $Filter_main
            
            if (-not $shouldSkip) {
                Write-Host "`nProcessing folder: " -NoNewline -ForegroundColor DarkGray
                Write-Host "$($folder.FullName)" -ForegroundColor Cyan
                $Files = ''
                $folder.Attributes = 'Directory', 'ReadOnly'
                [string]$full_path_folder = $folder.FullName
                $LastDirName = Split-Path -Path "$full_path_folder" -Leaf
                if (($rules) -and ($rules.ContainsKey($LastDirName))) {
                    $value = $rules[$LastDirName]
                    if (Test-Path "$full_path_folder\$value") {
                        $Files = Get-ChildItem -LiteralPath "$full_path_folder" -Filter $value
                    }
                }
                
                if (-not $Files) {
                    $Files = Find-Candidates -path $full_path_folder -search_depth $search_depth -priority $primaryType
                }
                
                if (-not $Files) {
                    if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host "$primaryType files not found, switching to $secondaryType search" }
                    $Files = Find-Candidates -path $full_path_folder -search_depth $search_depth -priority $secondaryType
                }
                
                if ($Files) {
                    #Testing path
                    try {
                        $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse:$($apply_depth -gt 0) -Depth $apply_depth -ErrorAction SilentlyContinue
                    } catch {
                        Write-Host 'Access to the path is denied. Cant proseed with desktop.ini file. Skiping...' -ForegroundColor Red
                        Write-Host "$($folder.FullName)"
                        continue
                    }
                    
                    if (!($noforce)){
                        #Forcing desktop.ini deletion
                        $desktopINI | Remove-Item -Force
                    } else {
                        if (!($desktopINI)) {
                            Write-Host "desktop.ini not found. Proceeding with creation" -ForegroundColor Green
                        } else {
                            $found = $false
                            $content = Get-Content -LiteralPath "$($desktopINI.FullName)" -ErrorAction Stop
                            foreach ($line in $content) {
                                if ($line -match '^IconResource=') {
                                    $found = $true
                                    break
                                }
                            }
                            if ($found) {
                                Write-Host "desktop.ini already exist. Skipping due to -noforce flag" -ForegroundColor Yellow
                                continue
                            }
                        }
                    }
                    
                    #### Creating desktop.ini file starts
                    $first_part = ''
                    
                    if (($Files.DirectoryName -ne $full_path_folder)) {
                        $exe_array = ($Files.DirectoryName).Split('\')
                        $folder_array = ($full_path_folder).Split('\')
                        $diff = (Compare-Object -ReferenceObject $exe_array -DifferenceObject $folder_array).InputObject
                        foreach ($k in $diff) {
                            $first_part = $first_part + '\' + $k
                        }
                    }
                    
                    $tmpDir = (Join-Path -Path "$env:TEMP" -ChildPath ([IO.Path]::GetRandomFileName()))
                    $null = mkdir -Path $tmpDir -Force
                    $tmp = "$tmpDir\desktop.ini"
                    
                    if ($first_part) {
                        $value = '.' + "$first_part\$Files" + ',0'
                    } else {
                        $value = '.\' + $Files + ',0'
                    }
                    
                    $ini = @(
                        '[.ShellClassInfo]'
                        "IconResource=$value"
                        #"InfoTip=$exeFiles"
                        '[ViewState]'
                        'Mode='
                        'Vid='
                        'FolderType=Generic') -join "`n"
                    
                    $null = New-Item -Path "$tmp" -Value $ini
                    
                    (Get-Item -LiteralPath $tmp).Attributes = 'Archive, System, Hidden'
                    
                    $shell = New-Object -ComObject Shell.Application
                    $shell.NameSpace($full_path_folder).MoveHere($tmp, 0x0004 + 0x0010 + 0x0400)
                    #### Creating desktop.ini file ends
                    
                    Remove-Item -Path "$tmpDir" -Force
                    
                    Write-Host "$($Files.Name) " -NoNewline -ForegroundColor DarkGray
                    Write-Host "--> " -NoNewline
                    Write-Host "$($folder.Name)" -ForegroundColor Green
                } else {
                    Write-Host "Proper file not found" -ForegroundColor Red
                    $foldersError += $full_path_folder
                }
            } else {
                Write-Host "`nSkipping filtered folder: " -NoNewline -ForegroundColor DarkGray
                Write-Host "$($folder.FullName)" -ForegroundColor Cyan
            }
        }
        
        if ($remove) {
            Write-Host "`nIcons have been removed from folders! In most cases, Explorer must be restarted!" -ForegroundColor Yellow
            Write-Host "Press " -NoNewline; Write-Host "R " -NoNewline -ForegroundColor Magenta; Write-Host "to restart"
            Write-Host "Press any other key to cancel"
            $key = [System.Console]::ReadKey($true)
            if ($key.Key -eq 'R') {
                Write-Host "Restarting Explorer..." -ForegroundColor Yellow
                Restart-ExplorerAsUser
                Write-Host "Explorer restarted." -ForegroundColor Green
            }
        }
        
        if ($foldersError) {
            Write-Host "`nFolders with errors:" -ForegroundColor Red
            Write-Host "$($foldersError -join "`n")" -ForegroundColor Red
            Write-Host "`nProper files not found. Try to increase search depth with -sd (-search_depth). Current value: " -NoNewline -ForegroundColor Red; Write-Host "$search_depth" -ForegroundColor Cyan
            Write-Host "You can use the -nf flag if you're happy with the results from the previous run - it won't overwrite the existing folder view"
            $foldersError = @()
        }
    } catch {
        Write-Host "`n$_" -ForegroundColor Red
        Write-Host "`n$($_.ScriptStackTrace)`n" -ForegroundColor Red
    }
    
    Timer -end
    
    if ($log){
        Logging -stop
    }
    
    if ($pause){
        pause
    }
}