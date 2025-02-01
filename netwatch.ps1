# Initialize variables
$isPaused = $false
$inspectMode = $false
$focusMode = $false
$inspectProcess = $null
$cacheFile = Join-Path $PWD "ip_cache.json"
$focusConnections = @{}
$inactiveFocusConnections = @{}
$focusStartTime = $null
$script:ipCache = @{}

# Function to display command menu
function Show-CommandMenu {
    $menuItems = @(
        "Press 'P' - Pause/Unpause",
        "Press 'E' - Export to CSV",
        "Press 'I' - Inspect Mode",
        "Press 'F' - Focus Mode",
        "Press 'C' - Export Cache"
    )
    # First row
    Write-Host "$($menuItems[0]) || $($menuItems[1]) || $($menuItems[2])"
    # Second row
    Write-Host "$($menuItems[3]) || $($menuItems[4])"
}

# Function to convert JSON to HashTable
function ConvertFrom-JsonToHashtable {
    param([string]$json)
    try {
        $obj = ConvertFrom-Json $json
        $hash = @{}
        foreach ($property in $obj.PSObject.Properties) {
            $hash[$property.Name] = $property.Value
        }
        return $hash
    }
    catch {
        Write-Host "Error converting JSON to hashtable: $_" -ForegroundColor Red
        return @{}
    }
}

# Load cache from file if it exists
if (Test-Path $cacheFile) {
    try {
        $script:ipCache = ConvertFrom-JsonToHashtable (Get-Content $cacheFile -Raw)
        Write-Host "Cache loaded successfully!" -ForegroundColor Green
        Write-Host "Loaded $($script:ipCache.Count) cached IP entries" -ForegroundColor DarkGray
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "Error loading cache file. Starting with empty cache." -ForegroundColor Red
        Start-Sleep -Seconds 2
        $script:ipCache = @{}
    }
}

# Function to update focus connections
function Update-FocusConnections {
    param (
        [Parameter(Mandatory=$true)]
        [array]$currentConnections
    )
    
    if ($script:focusMode) {
        $currentKeys = @()
        
        # Process current connections
        foreach ($conn in $currentConnections) {
            $connectionKey = "$($conn.ProcessName)_$($conn.RemoteAddress)"
            $currentKeys += $connectionKey
            
            # Add new connection or update existing one
            if (-not $script:focusConnections.ContainsKey($connectionKey)) {
                $script:focusConnections[$connectionKey] = @{
                    ProcessName = $conn.ProcessName
                    RemoteAddress = $conn.RemoteAddress
                    ISP = $conn.ISP
                    FirstSeen = Get-Date
                    LastSeen = Get-Date
                    Status = "Active"
                }
            } else {
                $script:focusConnections[$connectionKey].LastSeen = Get-Date
                $script:focusConnections[$connectionKey].Status = "Active"
            }
        }
        
        # Mark inactive connections
        $script:focusConnections.Keys | Where-Object { $_ -notin $currentKeys } | ForEach-Object {
            $script:focusConnections[$_].Status = "Inactive"
        }
    }
}

# Function to display focus connections
function Display-FocusConnections {
    if ($script:focusMode) {
        # Display active connections
        $activeConns = $script:focusConnections.Values | 
            Where-Object { $_.Status -eq "Active" } |
            Select-Object @{N='ProcessName';E={$_.ProcessName}},
                        @{N='RemoteAddress';E={$_.RemoteAddress}},
                        @{N='ISP';E={$_.ISP}},
                        @{N='FirstSeen';E={$_.FirstSeen}},
                        @{N='LastSeen';E={$_.LastSeen}}
        
        if ($activeConns) {
            Write-Host "`nActive Connections:" -ForegroundColor Green
            $activeConns | Format-Table -AutoSize
        }
        
        # Display inactive connections
        $inactiveConns = $script:focusConnections.Values | 
            Where-Object { $_.Status -eq "Inactive" } |
            Select-Object @{N='ProcessName';E={$_.ProcessName}},
                        @{N='RemoteAddress';E={$_.RemoteAddress}},
                        @{N='ISP';E={$_.ISP}},
                        @{N='FirstSeen';E={$_.FirstSeen}},
                        @{N='LastSeen';E={$_.LastSeen}}
        
        if ($inactiveConns) {
            Write-Host "`nInactive Connections:" -ForegroundColor Red
            $inactiveConns | Format-Table -AutoSize
        }
    }
}

# Function to save focus mode report
function Save-FocusReport {
    if ($script:focusMode) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportPath = Join-Path $PWD "focus_report_$timestamp.txt"
        
        $report = "Focus Mode Report - $(Get-Date)`n"
        $report += "Session Started: $($script:focusStartTime)`n"
        $report += "Duration: $([math]::Round(((Get-Date) - $script:focusStartTime).TotalMinutes, 2)) minutes`n"
        
        $activeConns = $script:focusConnections.Values | Where-Object { $_.Status -eq "Active" }
        $inactiveConns = $script:focusConnections.Values | Where-Object { $_.Status -eq "Inactive" }
        
        $report += "Total Tracked Connections: $($script:focusConnections.Count)`n"
        $report += "Active Connections: $($activeConns.Count)`n"
        $report += "Inactive Connections: $($inactiveConns.Count)`n`n"
        
        # Active Connections
        $report += "Active Connections:`n"
        $report += "-----------------`n"
        foreach ($conn in $activeConns) {
            $report += "Process: $($conn.ProcessName)`n"
            $report += "Remote Address: $($conn.RemoteAddress)`n"
            $report += "ISP: $($conn.ISP)`n"
            $report += "First Seen: $($conn.FirstSeen)`n"
            $report += "Last Seen: $($conn.LastSeen)`n"
            $report += "Status: Active`n`n"
        }
        
        # Inactive Connections
        $report += "Inactive Connections:`n"
        $report += "-------------------`n"
        foreach ($conn in $inactiveConns) {
            $report += "Process: $($conn.ProcessName)`n"
            $report += "Remote Address: $($conn.RemoteAddress)`n"
            $report += "ISP: $($conn.ISP)`n"
            $report += "First Seen: $($conn.FirstSeen)`n"
            $report += "Last Seen: $($conn.LastSeen)`n"
            $report += "Status: Inactive`n`n"
        }
        
        $report | Out-File -FilePath $reportPath
        Write-Host "`nFocus Mode report saved to: $reportPath" -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
}
# Function to export cache
function Export-IPCache {
    try {
        if (Test-Path $cacheFile) {
            $existingCache = ConvertFrom-JsonToHashtable (Get-Content $cacheFile -Raw)
            $initialCount = $existingCache.Count
            
            foreach ($key in $script:ipCache.Keys) {
                if (-not $existingCache.ContainsKey($key)) {
                    $existingCache[$key] = $script:ipCache[$key]
                }
            }
            
            $newEntries = $existingCache.Count - $initialCount
            $existingCache | ConvertTo-Json | Set-Content $cacheFile
            $script:ipCache = $existingCache
            
            Write-Host "`nCache exported successfully to: $cacheFile" -ForegroundColor Green
            Write-Host "Total cached entries: $($existingCache.Count)" -ForegroundColor DarkGray
            if ($newEntries -gt 0) {
                Write-Host "Added $newEntries new entries" -ForegroundColor Cyan
            } else {
                Write-Host "No new entries added" -ForegroundColor DarkGray
            }
        } else {
            $script:ipCache | ConvertTo-Json | Set-Content $cacheFile
            Write-Host "`nNew cache file created: $cacheFile" -ForegroundColor Green
            Write-Host "Exported $($script:ipCache.Count) IP entries" -ForegroundColor DarkGray
        }
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "`nError exporting cache: $_" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

# Function to get ISP information
function Get-IspInfo {
    param($ipAddress)
    if ($script:ipCache.ContainsKey($ipAddress)) {
        return $script:ipCache[$ipAddress]
    }
    if ($ipAddress -match '^(192\.168\.|10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|127\.|::1$)') {
        $ispInfo = "Local Network"
    } else {
        try {
            $result = Invoke-RestMethod -Uri "http://ip-api.com/json/$ipAddress" -TimeoutSec 2
            $ispInfo = if ($result.status -eq "success") { $result.isp } else { "Unknown" }
        } catch {
            $ispInfo = "Lookup Failed"
        }
    }
    $script:ipCache[$ipAddress] = $ispInfo
    return $ispInfo
}

# Function to add current focus view to cache
function Add-FocusToCache {
    if ($script:currentConnections) {
        $addedCount = 0
        foreach ($conn in $script:currentConnections) {
            if (-not $script:ipCache.ContainsKey($conn.RemoteAddress)) {
                $script:ipCache[$conn.RemoteAddress] = $conn.ISP
                $addedCount++
            }
        }
        if ($addedCount -gt 0) {
            Write-Host "`nAdded $addedCount new entries to cache" -ForegroundColor Green
            Export-IPCache
        } else {
            Write-Host "`nNo new entries to add to cache" -ForegroundColor Yellow
        }
        Start-Sleep -Seconds 2
    }
}

# Create a script block to handle key presses
$keyHandler = {
    param([System.ConsoleKeyInfo]$key)
    if ($key.Key -eq [System.ConsoleKey]::P) {
        $script:isPaused = -not $script:isPaused
    }
    elseif ($key.Key -eq [System.ConsoleKey]::E) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = Join-Path $PWD "connections_$timestamp.csv"
        $script:currentConnections | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "`nConnections exported to: $exportPath" -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    elseif ($key.Key -eq [System.ConsoleKey]::S -and $script:focusMode) {
        Save-FocusReport
    }
    elseif ($key.Key -eq [System.ConsoleKey]::C) {
        Export-IPCache
    }
    elseif ($key.KeyChar -eq '+' -and $script:focusMode) {
        Add-FocusToCache
    }
    elseif ($key.Key -eq [System.ConsoleKey]::I) {
        if ($script:inspectMode) {
            $script:inspectMode = $false
            $script:inspectProcess = $null
            Write-Host "`nExiting Inspect Mode" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        } else {
            $script:isPaused = $true
            Clear-Host
            Write-Host "Current Processes:" -ForegroundColor Cyan
            $script:currentConnections | Format-Table -AutoSize
            Write-Host "`nEnter the number of the process to inspect (or 'c' to cancel):" -ForegroundColor Yellow
            $input = Read-Host
            if ($input -ne 'c') {
                try {
                    $number = [int]$input
                    $selectedProcess = $script:currentConnections | Where-Object { $_.Number -eq $number }
                    if ($selectedProcess) {
                        $script:inspectMode = $true
                        $script:inspectProcess = $selectedProcess.ProcessName
                        Write-Host "`nInspect Mode activated for: $($script:inspectProcess)" -ForegroundColor Green
                    } else {
                        Write-Host "`nInvalid number selected" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "`nInvalid input" -ForegroundColor Red
                }
            }
            $script:isPaused = $false
            Start-Sleep -Seconds 1
        }
    }
    elseif ($key.Key -eq [System.ConsoleKey]::F) {
        $script:focusMode = -not $script:focusMode
        if ($script:focusMode) {
            $script:focusStartTime = Get-Date
            $script:focusConnections = @{}
            Write-Host "`nFocus Mode activated (showing uncached IPs only)" -ForegroundColor Green
        } else {
            Save-FocusReport
            Write-Host "`nExiting Focus Mode" -ForegroundColor Yellow
        }
        Start-Sleep -Seconds 1
    }
}

# Main monitoring loop
while ($true) {
    if (-not $script:isPaused) {
        $connections = Get-NetTCPConnection | Where-Object {
            $_.RemoteAddress -ne "0.0.0.0" -and
            $_.RemoteAddress -ne $_.LocalAddress -and
            $_.RemoteAddress -notmatch "^(127\.0\.0\.1|::1|::)$" -and
            $_.LocalAddress -notmatch "^(127\.0\.0\.1|::1|::)$"
        }
        
        Clear-Host
        Write-Host "Monitoring External Connections (Excluding Loopback)...`n"
        Show-CommandMenu
        Write-Host
        if ($script:focusMode) {
            Write-Host "Focus Mode Options:" -ForegroundColor Cyan
            Write-Host "Press '+' to add current view to cache" -ForegroundColor Green
            Write-Host "Press 'S' to save Focus Mode report" -ForegroundColor Green
            Write-Host
        }
        
        if ($script:isPaused) {
            Write-Host "`nPAUSED" -ForegroundColor Yellow
        }
        if ($script:inspectMode) {
            Write-Host "INSPECT MODE: $($script:inspectProcess)" -ForegroundColor Green
        }
        if ($script:focusMode) {
            Write-Host "FOCUS MODE: Showing uncached IPs only" -ForegroundColor Cyan
            if ($script:focusStartTime) {
                $duration = [math]::Round(((Get-Date) - $script:focusStartTime).TotalMinutes, 2)
                Write-Host "Focus Duration: $duration minutes" -ForegroundColor DarkGray
            }
        }
        
        if ($connections) {
            $sortedConnections = $connections | ForEach-Object {
                $process = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
                
                if ($process -and $process.ProcessName -ne "Idle") {
                    $shouldShow = $true
                    if ($script:inspectMode) {
                        $shouldShow = $process.ProcessName -eq $script:inspectProcess
                    }
                    elseif ($script:focusMode) {
                        $shouldShow = -not $script:ipCache.ContainsKey($_.RemoteAddress)
                    }
                    
                    if ($shouldShow) {
                        [PSCustomObject]@{
                            ProcessName = $process.ProcessName
                            OwningProcess = $_.OwningProcess
                            LocalAddress = $_.LocalAddress
                            LocalPort = $_.LocalPort
                            RemoteAddress = $_.RemoteAddress
                            ISP = Get-IspInfo -ipAddress $_.RemoteAddress
                            RemotePort = $_.RemotePort
                            State = $_.State
                        }
                    }
                }
            } | Sort-Object ProcessName

            if ($sortedConnections) {
                $counter = 0
                $script:currentConnections = $sortedConnections | ForEach-Object {
                    $counter++
                    [PSCustomObject]@{
                        'Number' = $counter
                        ProcessName = $_.ProcessName
                        OwningProcess = $_.OwningProcess
                        LocalAddress = $_.LocalAddress
                        LocalPort = $_.LocalPort
                        RemoteAddress = $_.RemoteAddress
                        ISP = $_.ISP
                        RemotePort = $_.RemotePort
                        State = $_.State
                    }
                }

                if ($script:focusMode) {
                    Update-FocusConnections -currentConnections $script:currentConnections
                    Display-FocusConnections
                } else {
                    $script:currentConnections | Format-Table -AutoSize
                }

                Write-Host "`nTotal Active Connections: $counter" -ForegroundColor Cyan
                Write-Host "Cached IP lookups: $($script:ipCache.Count)" -ForegroundColor DarkGray
                if ($script:focusMode) {
                    $activeCount = ($script:focusConnections.Values | Where-Object { $_.Status -eq "Active" }).Count
                    $inactiveCount = ($script:focusConnections.Values | Where-Object { $_.Status -eq "Inactive" }).Count
                    Write-Host "Focus Mode - Active/Inactive Connections: $activeCount/$inactiveCount" -ForegroundColor DarkGray
                }
            }
        } else {
            Write-Host "No external connections found."
        }
    }
    
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        & $keyHandler $key
    }
    
    Start-Sleep -Milliseconds 500
}
