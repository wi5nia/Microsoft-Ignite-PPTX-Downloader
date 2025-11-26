#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Download Microsoft Ignite 2025 slide decks.

.DESCRIPTION
    This script fetches the 2025 session catalogue and downloads any session
    that provides a PowerPoint deck. Files are saved into an "ignite_2025_slides"
    directory. Uses PowerShell Core's built-in web capabilities.

.PARAMETER MaxConcurrency
    Number of concurrent downloads (default: 10)

.PARAMETER DestinationPath
    Directory to save slides (default: "ignite_2025_slides")

.EXAMPLE
    ./Download-Ignite2025Slides.ps1
    
.EXAMPLE
    ./Download-Ignite2025Slides.ps1 -MaxConcurrency 5 -DestinationPath "slides"
#>

[CmdletBinding()]
param(
    [int]$MaxConcurrency = 10,
    [string]$DestinationPath = ".\ignite_2025_slides"
)

# Constants
$API_URL = "https://api-v2.ignite.microsoft.com/api/session/all/en-US"
$CHUNK_SIZE = 1MB

# Initialize counters (thread-safe using synchronized hashtable)
$script:Counters = [System.Collections.Hashtable]::Synchronized(@{
    downloaded = 0
    no_deck = 0
    existing = 0
    failed = 0
})

function Invoke-SanitizeFilename {
    <#
    .SYNOPSIS
        Sanitizes a filename by replacing spaces with underscores and removing unsafe characters.
    #>
    param([string]$Value)
    
    $sanitized = $Value.Trim() -replace '\s+', '_'
    $sanitized = $sanitized -replace '[^\w\._-]', ''
    return $sanitized
}

function Invoke-FetchSessions {
    <#
    .SYNOPSIS
        Retrieves all session data from the Ignite v2 API.
    #>
    try {
        Write-Host "Fetching session catalogue..."
        $response = Invoke-RestMethod -Uri $API_URL -Method Get -TimeoutSec 30
        return $response
    }
    catch {
        Write-Error "Failed to fetch sessions: $($_.Exception.Message)"
        throw
    }
}

function Invoke-DownloadFile {
    <#
    .SYNOPSIS
        Downloads a file from the specified URL to the destination path.
    #>
    param(
        [string]$Url,
        [string]$DestPath,
        [string]$SessionCode
    )
    
    # Headers to mimic a browser; otherwise Medius may return HTTP 403
    $headers = @{
        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        'Referer' = 'https://ignite.microsoft.com/'
    }
    
    try {
        # Create parent directory if it doesn't exist
        $parentDir = Split-Path -Parent $DestPath
        if (-not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }
        
        # Download file
        Invoke-WebRequest -Uri $Url -OutFile $DestPath -Headers $headers -TimeoutSec 30
        return @{ Success = $true; Result = $SessionCode }
    }
    catch {
        return @{ Success = $false; Result = $_.Exception.Message }
    }
}

function Invoke-DownloadSession {
    <#
    .SYNOPSIS
        Downloads a single session's slides.
    #>
    param([PSCustomObject]$Session)
    
    $slideUrl = $Session.slideDeck
    if (-not $slideUrl -or $slideUrl -eq "") {
        $script:Counters.no_deck++
        return @{ Status = "no_deck"; Code = 0 }
    }
    
    $sessionCode = Invoke-SanitizeFilename -Value ($Session.sessionCode ?? "")
    $title = Invoke-SanitizeFilename -Value ($Session.title ?? "untitled")
    $destPath = Join-Path $DestinationPath "$sessionCode`_$title.pptx"
    
    # Skip if file already exists and has content
    if (Test-Path $destPath) {
        $fileInfo = Get-Item $destPath
        if ($fileInfo.Length -gt 0) {
            $script:Counters.existing++
            return @{ Status = "existing"; Code = 0 }
        }
    }
    
    Write-Host "Downloading $sessionCode – $($Session.title)"
    $downloadResult = Invoke-DownloadFile -Url $slideUrl -DestPath $destPath -SessionCode $sessionCode
    
    if ($downloadResult.Success) {
        $script:Counters.downloaded++
        return @{ Status = "downloaded"; Code = 1 }
    }
    else {
        Write-Host "Failed to download $sessionCode`: $($downloadResult.Result)" -ForegroundColor Red
        # Clean up failed download
        if (Test-Path $destPath) {
            Remove-Item $destPath -Force
        }
        $script:Counters.failed++
        return @{ Status = "failed"; Code = 0 }
    }
}

function Start-Main {
    <#
    .SYNOPSIS
        Main function that orchestrates the download process.
    #>
    try {
        # Fetch sessions
        $sessions = Invoke-FetchSessions
        Write-Host "Found $($sessions.Count) sessions total`n"
        
        # Create destination directory
        if (-not (Test-Path $DestinationPath)) {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        }
        
        # Process sessions with job-based parallel execution
        $jobs = @()
        $processed = 0
        
        foreach ($session in $sessions) {
            # Wait if we've reached max concurrency
            while ($jobs.Count -ge $MaxConcurrency) {
                $completedJob = $jobs | Where-Object { $_.State -eq 'Completed' } | Select-Object -First 1
                if ($completedJob) {
                    $result = Receive-Job $completedJob
                    Remove-Job $completedJob
                    $jobs = $jobs | Where-Object { $_.Id -ne $completedJob.Id }
                    
                    if ($result) {
                        $script:Counters[$result.Status]++
                        if ($result.Status -eq "downloaded") {
                            $processed++
                            Write-Host "Progress: $processed/$($sessions.Count) sessions processed"
                        }
                    }
                }
                else {
                    Start-Sleep -Milliseconds 100
                }
            }
            
            # Start new job
            $job = Start-Job -ScriptBlock {
                param($sessionData, $destPath, $apiUrl)
                
                function Sanitize-Filename {
                    param([string]$Value)
                    $sanitized = $Value.Trim() -replace '\s+', '_'
                    return $sanitized -replace '[^\w\._-]', ''
                }
                
                function Download-File {
                    param($Url, $DestPath, $SessionCode)
                    
                    $headers = @{
                        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                        'Referer' = 'https://ignite.microsoft.com/'
                    }
                    
                    try {
                        $parentDir = Split-Path -Parent $DestPath
                        if (-not (Test-Path $parentDir)) {
                            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                        }
                        
                        Invoke-WebRequest -Uri $Url -OutFile $DestPath -Headers $headers -TimeoutSec 30
                        return @{ Success = $true; Result = $SessionCode }
                    }
                    catch {
                        return @{ Success = $false; Result = $_.Exception.Message }
                    }
                }
                
                # Process the session
                $slideUrl = $sessionData.slideDeck
                if (-not $slideUrl -or $slideUrl -eq "") {
                    return @{ Status = "no_deck"; Code = 0 }
                }
                
                $sessionCode = Sanitize-Filename -Value ($sessionData.sessionCode ?? "")
                $title = Sanitize-Filename -Value ($sessionData.title ?? "untitled")
                $filePath = Join-Path $destPath "$sessionCode`_$title.pptx"
                
                if (Test-Path $filePath) {
                    $fileInfo = Get-Item $filePath
                    if ($fileInfo.Length -gt 0) {
                        return @{ Status = "existing"; Code = 0 }
                    }
                }
                
                Write-Host "Downloading $sessionCode – $($sessionData.title)"
                $downloadResult = Download-File -Url $slideUrl -DestPath $filePath -SessionCode $sessionCode
                
                if ($downloadResult.Success) {
                    return @{ Status = "downloaded"; Code = 1 }
                }
                else {
                    Write-Host "Failed to download $sessionCode`: $($downloadResult.Result)" -ForegroundColor Red
                    if (Test-Path $filePath) {
                        Remove-Item $filePath -Force
                    }
                    return @{ Status = "failed"; Code = 0 }
                }
            } -ArgumentList $session, $DestinationPath, $API_URL
            
            $jobs += $job
        }
        
        # Wait for remaining jobs to complete
        while ($jobs.Count -gt 0) {
            $completedJob = $jobs | Where-Object { $_.State -eq 'Completed' } | Select-Object -First 1
            if ($completedJob) {
                $result = Receive-Job $completedJob
                Remove-Job $completedJob
                $jobs = $jobs | Where-Object { $_.Id -ne $completedJob.Id }
                
                if ($result) {
                    $script:Counters[$result.Status]++
                    if ($result.Status -eq "downloaded") {
                        $processed++
                        Write-Host "Progress: $processed/$($sessions.Count) sessions processed"
                    }
                }
            }
            else {
                Start-Sleep -Milliseconds 100
            }
        }
        
        # Print summary
        Write-Host "`n$('=' * 50)"
        Write-Host "Summary:"
        Write-Host "Downloaded: $($script:Counters.downloaded)"
        Write-Host "Skipped (no deck): $($script:Counters.no_deck)"
        Write-Host "Skipped (already existed): $($script:Counters.existing)"
        Write-Host "Failed: $($script:Counters.failed)"
        Write-Host "$('=' * 50)"
    }
    catch {
        Write-Error "Script execution failed: $($_.Exception.Message)"
        exit 1
    }
}

# Main execution
if ($MyInvocation.InvocationName -ne '.') {
    Start-Main
}