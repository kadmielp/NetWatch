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
$script:showDetails = $false
$script:sessionConnections = @{}
$script:sessionStartTime = Get-Date

# Function to get color for connection state
function Get-StateColor {
    param([string]$state)
    
    switch ($state) {
        "Established" { return "Green" }
        "Listen" { return "Cyan" }
        "TimeWait" { return "Yellow" }
        "CloseWait" { return "DarkYellow" }
        "SynSent" { return "Magenta" }
        "SynReceived" { return "DarkMagenta" }
        "FinWait1" { return "Red" }
        "FinWait2" { return "DarkRed" }
        "LastAck" { return "DarkRed" }
        "Closed" { return "Gray" }
        "Closing" { return "Red" }
        default { return "White" }
    }
}

# Function to display command menu
function Show-CommandMenu {
    $menuItems = @(
        "Press 'P' - Pause/Unpause",
        "Press 'A' - Export active to CSV",
        "Press 'E' - Export session to CSV",
        "Press 'I' - Inspect Mode",
        "Press 'F' - Focus Mode",
        "Press 'C' - Save Cache",
        "Press 'D' - Show Details"
    )
    # First row
    Write-Host "$($menuItems[0]) || $($menuItems[1]) || $($menuItems[2])"
    # Second row
    Write-Host "$($menuItems[3]) || $($menuItems[4]) || $($menuItems[5])"
    # Third row
    Write-Host "$($menuItems[6])"
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
                    ISP = $conn.GeoInfo.ISP
                    Country = $conn.GeoInfo.Country
                    CountryCode = $conn.GeoInfo.CountryCode
                    City = $conn.GeoInfo.City
                    Region = $conn.GeoInfo.Region
                    GeoInfo = $conn.GeoInfo
                    FirstSeen = Get-Date
                    LastSeen = Get-Date
                    Status = "Active"
                }
            } else {
                $script:focusConnections[$connectionKey].LastSeen = Get-Date
                $script:focusConnections[$connectionKey].Status = "Active"
                # Update geolocation if available
                if ($conn.GeoInfo) {
                    $script:focusConnections[$connectionKey].ISP = $conn.GeoInfo.ISP
                    $script:focusConnections[$connectionKey].Country = $conn.GeoInfo.Country
                    $script:focusConnections[$connectionKey].CountryCode = $conn.GeoInfo.CountryCode
                    $script:focusConnections[$connectionKey].City = $conn.GeoInfo.City
                    $script:focusConnections[$connectionKey].Region = $conn.GeoInfo.Region
                    $script:focusConnections[$connectionKey].GeoInfo = $conn.GeoInfo
                }
            }
        }
        
        # Mark inactive connections
        $script:focusConnections.Keys | Where-Object { $_ -notin $currentKeys } | ForEach-Object {
            $script:focusConnections[$_].Status = "Inactive"
        }
    }
}

# Function to update session connections (track all connections during session)
function Update-SessionConnections {
    param (
        [Parameter(Mandatory=$true)]
        [array]$currentConnections
    )
    
    $currentKeys = @()
    
    # Process current connections
    foreach ($conn in $currentConnections) {
        $connectionKey = "$($conn.ProcessName)_$($conn.RemoteAddress):$($conn.RemotePort)"
        $currentKeys += $connectionKey
        
        # Add new connection or update existing one
        if (-not $script:sessionConnections.ContainsKey($connectionKey)) {
            $geoInfo = if ($conn.GeoInfo) { $conn.GeoInfo } else { @{ ISP = ""; Country = ""; CountryCode = ""; City = ""; Region = "" } }
            $script:sessionConnections[$connectionKey] = @{
                ProcessName = $conn.ProcessName
                OwningProcess = $conn.OwningProcess
                ProcessPath = if ($conn.ProcessPath) { $conn.ProcessPath } else { "" }
                LocalAddress = $conn.LocalAddress
                LocalPort = $conn.LocalPort
                RemoteAddress = $conn.RemoteAddress
                RemotePort = $conn.RemotePort
                State = $conn.State
                ISP = if ($geoInfo -is [hashtable]) { $geoInfo.ISP } else { if ($geoInfo) { $geoInfo } else { "" } }
                Country = if ($geoInfo -is [hashtable]) { $geoInfo.Country } else { "" }
                CountryCode = if ($geoInfo -is [hashtable]) { $geoInfo.CountryCode } else { "" }
                City = if ($geoInfo -is [hashtable]) { $geoInfo.City } else { "" }
                Region = if ($geoInfo -is [hashtable]) { $geoInfo.Region } else { "" }
                GeoInfo = $geoInfo
                FirstSeen = Get-Date
                LastSeen = Get-Date
                Status = "Active"
                ConnectionCount = 1
            }
        } else {
            $script:sessionConnections[$connectionKey].LastSeen = Get-Date
            $script:sessionConnections[$connectionKey].Status = "Active"
            $script:sessionConnections[$connectionKey].ConnectionCount++
            # Update connection details if they changed
            if ($conn.State) { $script:sessionConnections[$connectionKey].State = $conn.State }
            if ($conn.ProcessPath) { $script:sessionConnections[$connectionKey].ProcessPath = $conn.ProcessPath }
            # Update geolocation if available
            if ($conn.GeoInfo) {
                $geoInfo = $conn.GeoInfo
                $script:sessionConnections[$connectionKey].ISP = if ($geoInfo -is [hashtable]) { $geoInfo.ISP } else { if ($geoInfo) { $geoInfo } else { "" } }
                $script:sessionConnections[$connectionKey].Country = if ($geoInfo -is [hashtable]) { $geoInfo.Country } else { "" }
                $script:sessionConnections[$connectionKey].CountryCode = if ($geoInfo -is [hashtable]) { $geoInfo.CountryCode } else { "" }
                $script:sessionConnections[$connectionKey].City = if ($geoInfo -is [hashtable]) { $geoInfo.City } else { "" }
                $script:sessionConnections[$connectionKey].Region = if ($geoInfo -is [hashtable]) { $geoInfo.Region } else { "" }
                $script:sessionConnections[$connectionKey].GeoInfo = $geoInfo
            } else {
                # Ensure geoInfo is set even if conn.GeoInfo is null
                $geoInfo = if ($script:sessionConnections[$connectionKey].GeoInfo) { $script:sessionConnections[$connectionKey].GeoInfo } else { @{ ISP = ""; Country = ""; CountryCode = ""; City = ""; Region = "" } }
            }
        }
    }
    
    # Mark inactive connections
    $script:sessionConnections.Keys | Where-Object { $_ -notin $currentKeys } | ForEach-Object {
        $script:sessionConnections[$_].Status = "Inactive"
    }
}

# Function to display connections with colored states
function Write-ColoredConnectionsTable {
    param([array]$connections, [bool]$highlightActive = $false)
    
    if (-not $connections -or $connections.Count -eq 0) { return }
    
    # Print header
    $header = "Number | ProcessName         | LocalAddress   | LocalPort | RemoteAddress     | RemotePort | State        | Country     | City            | ISP"
    Write-Host $header -ForegroundColor Cyan
    Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
    
    foreach ($conn in $connections) {
        $stateColor = Get-StateColor -state $conn.State
        $baseColor = "White"
        
        # In Focus Mode, highlight active vs inactive
        if ($highlightActive) {
            $connectionKey = "$($conn.ProcessName)_$($conn.RemoteAddress)"
            if ($script:focusConnections.ContainsKey($connectionKey)) {
                if ($script:focusConnections[$connectionKey].Status -eq "Active") {
                    $baseColor = "Green"
                } else {
                    $baseColor = "DarkGray"
                }
            }
        }
        
        # Get geolocation display values
        $geoInfo = if ($conn.GeoInfo) { $conn.GeoInfo } else { @{ Country = ""; CountryCode = ""; City = ""; ISP = "" } }
        $countryDisplay = if ($geoInfo.CountryCode) { "$($geoInfo.CountryCode)" } else { "" }
        $cityDisplay = if ($geoInfo.City) { $geoInfo.City } else { "" }
        $ispDisplay = if ($geoInfo.ISP) { $geoInfo.ISP } else { "" }
        
        # Build the line by parts for precise color control
        $beforeState = "{0,-6} | {1,-18} | {2,-13} | {3,-9} | {4,-17} | {5,-10} | " -f `
            $conn.Number,
            $conn.ProcessName,
            $conn.LocalAddress,
            $conn.LocalPort,
            $conn.RemoteAddress,
            $conn.RemotePort
        
        $statePadded = "{0,-11}" -f $conn.State
        $countryPadded = "{0,-11}" -f $countryDisplay
        $cityPadded = "{0,-16}" -f $cityDisplay
        $afterState = " | $countryPadded | $cityPadded | $ispDisplay"
        
        # Write the line with appropriate colors
        Write-Host $beforeState -NoNewline -ForegroundColor $baseColor
        Write-Host $statePadded -NoNewline -ForegroundColor $stateColor
        Write-Host $afterState -ForegroundColor $baseColor
        
        # Show process path if details are enabled
        if ($script:showDetails) {
            if ($conn.ProcessPath) {
                $pathDisplay = "       -> $($conn.ProcessPath)"
                # Truncate if too long
                if ($pathDisplay.Length -gt 120) {
                    $pathDisplay = $pathDisplay.Substring(0, 117) + "..."
                }
                Write-Host $pathDisplay -ForegroundColor DarkGray
            } else {
                # Show indicator that path couldn't be retrieved
                Write-Host "       -> (Path not accessible)" -ForegroundColor DarkGray
            }
        }
    }
}

# Function to display focus connections
function Display-FocusConnections {
    if ($script:focusMode) {
        # Get connections from current view and merge with focus tracking
        if ($script:currentConnections) {
            Write-ColoredConnectionsTable -connections $script:currentConnections -highlightActive $true
        } else {
            # Fallback: Display from focusConnections hashtable
            $activeConns = $script:focusConnections.Values | 
                Where-Object { $_.Status -eq "Active" } |
                Select-Object @{N='ProcessName';E={$_.ProcessName}},
                            @{N='RemoteAddress';E={$_.RemoteAddress}},
                            @{N='ISP';E={if ($_.GeoInfo) { $_.GeoInfo.ISP } else { $_.ISP }}},
                            @{N='Country';E={if ($_.GeoInfo) { "$($_.GeoInfo.Country) ($($_.GeoInfo.CountryCode))" } else { "" }}},
                            @{N='City';E={if ($_.GeoInfo) { $_.GeoInfo.City } else { "" }}},
                            @{N='Region';E={if ($_.GeoInfo) { $_.GeoInfo.Region } else { "" }}},
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
                            @{N='ISP';E={if ($_.GeoInfo) { $_.GeoInfo.ISP } else { $_.ISP }}},
                            @{N='Country';E={if ($_.GeoInfo) { "$($_.GeoInfo.Country) ($($_.GeoInfo.CountryCode))" } else { "" }}},
                            @{N='City';E={if ($_.GeoInfo) { $_.GeoInfo.City } else { "" }}},
                            @{N='Region';E={if ($_.GeoInfo) { $_.GeoInfo.Region } else { "" }}},
                            @{N='FirstSeen';E={$_.FirstSeen}},
                            @{N='LastSeen';E={$_.LastSeen}}
            
            if ($inactiveConns) {
                Write-Host "`nInactive Connections:" -ForegroundColor DarkGray
                $inactiveConns | Format-Table -AutoSize
            }
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
            if ($conn.GeoInfo) {
                $report += "ISP: $($conn.GeoInfo.ISP)`n"
                if ($conn.GeoInfo.Country) { $report += "Country: $($conn.GeoInfo.Country)`n" }
                if ($conn.GeoInfo.CountryCode) { $report += "Country Code: $($conn.GeoInfo.CountryCode)`n" }
                if ($conn.GeoInfo.City) { $report += "City: $($conn.GeoInfo.City)`n" }
                if ($conn.GeoInfo.Region) { $report += "Region: $($conn.GeoInfo.Region)`n" }
            } else {
                $report += "ISP: $($conn.ISP)`n"
            }
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
            if ($conn.GeoInfo) {
                $report += "ISP: $($conn.GeoInfo.ISP)`n"
                if ($conn.GeoInfo.Country) { $report += "Country: $($conn.GeoInfo.Country)`n" }
                if ($conn.GeoInfo.CountryCode) { $report += "Country Code: $($conn.GeoInfo.CountryCode)`n" }
                if ($conn.GeoInfo.City) { $report += "City: $($conn.GeoInfo.City)`n" }
                if ($conn.GeoInfo.Region) { $report += "Region: $($conn.GeoInfo.Region)`n" }
            } else {
                $report += "ISP: $($conn.ISP)`n"
            }
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

# Function to get geolocation information (Country, City, Region, ISP)
function Get-IspInfo {
    param($ipAddress)
    
    # Check cache first
    if ($script:ipCache.ContainsKey($ipAddress)) {
        $cached = $script:ipCache[$ipAddress]
        # Handle backward compatibility: if it's a string (old format), return as ISP-only object
        if ($cached -is [string]) {
            return @{
                ISP = $cached
                Country = ""
                CountryCode = ""
                City = ""
                Region = ""
            }
        }
        # Handle PSCustomObject (from JSON deserialization) - convert to hashtable
        if ($cached -is [PSCustomObject]) {
            return @{
                ISP = $cached.ISP
                Country = $cached.Country
                CountryCode = $cached.CountryCode
                City = $cached.City
                Region = $cached.Region
            }
        }
        # Already a hashtable
        return $cached
    }
    
    # Check if it's a local network address
    if ($ipAddress -match '^(192\.168\.|10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|127\.|::1$)') {
        $geoInfo = @{
            ISP = "Local Network"
            Country = "Local"
            CountryCode = "LOC"
            City = ""
            Region = ""
        }
    } else {
        try {
            $result = Invoke-RestMethod -Uri "http://ip-api.com/json/$ipAddress" -TimeoutSec 2
            if ($result.status -eq "success") {
                $geoInfo = @{
                    ISP = $result.isp
                    Country = $result.country
                    CountryCode = $result.countryCode
                    City = $result.city
                    Region = $result.regionName
                }
            } else {
                $geoInfo = @{
                    ISP = "Unknown"
                    Country = ""
                    CountryCode = ""
                    City = ""
                    Region = ""
                }
            }
        } catch {
            $geoInfo = @{
                ISP = "Lookup Failed"
                Country = ""
                CountryCode = ""
                City = ""
                Region = ""
            }
        }
    }
    
    $script:ipCache[$ipAddress] = $geoInfo
    return $geoInfo
}

# Helper function to format geolocation for display
function Format-GeoInfo {
    param($geoInfo)
    
    if ($geoInfo -is [string]) {
        return $geoInfo
    }
    
    $parts = @()
    if ($geoInfo.City) { $parts += $geoInfo.City }
    if ($geoInfo.Region) { $parts += $geoInfo.Region }
    if ($geoInfo.Country) {
        if ($geoInfo.CountryCode) {
            $parts += "$($geoInfo.Country) ($($geoInfo.CountryCode))"
        } else {
            $parts += $geoInfo.Country
        }
    }
    if ($geoInfo.ISP) { $parts += $geoInfo.ISP }
    
    return ($parts -join ", ")
}

# Function to show interactive process selector
function Show-ProcessSelector {
    param([array]$processes)
    
    if (-not $processes -or $processes.Count -eq 0) {
        Write-Host "`nNo processes available to inspect" -ForegroundColor Yellow
        return $null
    }
    
    # Get unique processes grouped by ProcessName
    $uniqueProcesses = $processes | Group-Object ProcessName | ForEach-Object {
        [PSCustomObject]@{
            ProcessName = $_.Name
            Count = $_.Count
            Number = ($_.Group | Select-Object -First 1).Number
        }
    } | Sort-Object ProcessName
    
    $selectedIndex = 0
    $maxIndex = $uniqueProcesses.Count - 1
    
    # Clear screen and show menu
    Clear-Host
    Write-Host "Select Process to Inspect (Use Up/Down arrows, Enter to select, Esc to cancel)" -ForegroundColor Cyan
    Write-Host ("-" * 70) -ForegroundColor DarkGray
    
    $startTop = [Console]::CursorTop
    
    while ($true) {
        # Save current cursor position for menu items
        $menuStartTop = $startTop
        
        # Display all processes with highlighting
        for ($i = 0; $i -le $maxIndex; $i++) {
            $proc = $uniqueProcesses[$i]
            if ($i -eq $selectedIndex) {
                Write-Host (" > {0,-25} ({1} connection(s))" -f $proc.ProcessName, $proc.Count) -ForegroundColor Green -BackgroundColor DarkGray
            } else {
                Write-Host ("   {0,-25} ({1} connection(s))" -f $proc.ProcessName, $proc.Count) -ForegroundColor White
            }
        }
        
        # Read key
        $key = [Console]::ReadKey($true)
        
        # Handle key presses
        if ($key.Key -eq [System.ConsoleKey]::UpArrow) {
            $selectedIndex = if ($selectedIndex -gt 0) { $selectedIndex - 1 } else { $maxIndex }
            # Move cursor back to redraw menu
            [Console]::SetCursorPosition(0, $menuStartTop)
        }
        elseif ($key.Key -eq [System.ConsoleKey]::DownArrow) {
            $selectedIndex = if ($selectedIndex -lt $maxIndex) { $selectedIndex + 1 } else { 0 }
            # Move cursor back to redraw menu
            [Console]::SetCursorPosition(0, $menuStartTop)
        }
        elseif ($key.Key -eq [System.ConsoleKey]::Enter) {
            # Clear menu area
            [Console]::SetCursorPosition(0, $menuStartTop)
            for ($i = 0; $i -le $maxIndex; $i++) {
                Write-Host (" " * 70)
            }
            [Console]::SetCursorPosition(0, $menuStartTop)
            return $uniqueProcesses[$selectedIndex].ProcessName
        }
        elseif ($key.Key -eq [System.ConsoleKey]::Escape) {
            # Clear menu area
            [Console]::SetCursorPosition(0, $menuStartTop)
            for ($i = 0; $i -le $maxIndex; $i++) {
                Write-Host (" " * 70)
            }
            [Console]::SetCursorPosition(0, $menuStartTop)
            return $null
        }
        else {
            # Move cursor back for redraw if invalid key
            [Console]::SetCursorPosition(0, $menuStartTop)
        }
    }
}

# Function to add current focus view to cache
function Add-FocusToCache {
    if ($script:currentConnections) {
        $addedCount = 0
        foreach ($conn in $script:currentConnections) {
            if (-not $script:ipCache.ContainsKey($conn.RemoteAddress)) {
                $script:ipCache[$conn.RemoteAddress] = if ($conn.GeoInfo) { $conn.GeoInfo } else { @{ ISP = "Unknown" } }
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
    elseif ($key.Key -eq [System.ConsoleKey]::A) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = Join-Path $PWD "connections_$timestamp.csv"
        # Create expanded connection objects with geolocation columns for CSV export
        $exportData = $script:currentConnections | ForEach-Object {
            $geoInfo = if ($_.GeoInfo) { $_.GeoInfo } else { @{ ISP = ""; Country = ""; CountryCode = ""; City = ""; Region = "" } }
            [PSCustomObject]@{
                Number = $_.Number
                ProcessName = $_.ProcessName
                OwningProcess = $_.OwningProcess
                ProcessPath = if ($_.ProcessPath) { $_.ProcessPath } else { "" }
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort = $_.RemotePort
                State = $_.State
                ISP = if ($geoInfo -is [hashtable]) { $geoInfo.ISP } else { $geoInfo }
                Country = if ($geoInfo -is [hashtable]) { $geoInfo.Country } else { "" }
                CountryCode = if ($geoInfo -is [hashtable]) { $geoInfo.CountryCode } else { "" }
                City = if ($geoInfo -is [hashtable]) { $geoInfo.City } else { "" }
                Region = if ($geoInfo -is [hashtable]) { $geoInfo.Region } else { "" }
            }
        }
        $exportData | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "`nConnections exported to: $exportPath" -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    elseif ($key.Key -eq [System.ConsoleKey]::E) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = Join-Path $PWD "session_connections_$timestamp.csv"
        
        if ($script:sessionConnections.Count -eq 0) {
            Write-Host "`nNo session connections to export" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        } else {
            $exportData = $script:sessionConnections.Values | ForEach-Object {
                $geoInfo = if ($_.GeoInfo) { $_.GeoInfo } else { @{ ISP = ""; Country = ""; CountryCode = ""; City = ""; Region = "" } }
                [PSCustomObject]@{
                    ProcessName = $_.ProcessName
                    OwningProcess = $_.OwningProcess
                    ProcessPath = $_.ProcessPath
                    LocalAddress = $_.LocalAddress
                    LocalPort = $_.LocalPort
                    RemoteAddress = $_.RemoteAddress
                    RemotePort = $_.RemotePort
                    State = $_.State
                    ISP = if ($geoInfo -is [hashtable]) { $geoInfo.ISP } else { if ($geoInfo) { $geoInfo } else { "" } }
                    Country = if ($geoInfo -is [hashtable]) { $geoInfo.Country } else { "" }
                    CountryCode = if ($geoInfo -is [hashtable]) { $geoInfo.CountryCode } else { "" }
                    City = if ($geoInfo -is [hashtable]) { $geoInfo.City } else { "" }
                    Region = if ($geoInfo -is [hashtable]) { $geoInfo.Region } else { "" }
                    FirstSeen = $_.FirstSeen
                    LastSeen = $_.LastSeen
                    Status = $_.Status
                    ConnectionCount = $_.ConnectionCount
                }
            }
            $exportData | Export-Csv -Path $exportPath -NoTypeInformation
            $sessionDuration = [math]::Round(((Get-Date) - $script:sessionStartTime).TotalMinutes, 2)
            Write-Host "`nSession connections exported to: $exportPath" -ForegroundColor Green
            Write-Host "Total connections tracked: $($script:sessionConnections.Count)" -ForegroundColor DarkGray
            Write-Host "Session duration: $sessionDuration minutes" -ForegroundColor DarkGray
            Start-Sleep -Seconds 2
        }
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
            $selectedProcessName = Show-ProcessSelector -processes $script:currentConnections
            if ($selectedProcessName) {
                $script:inspectMode = $true
                $script:inspectProcess = $selectedProcessName
                Write-Host "`nInspect Mode activated for: $($script:inspectProcess)" -ForegroundColor Green
                Start-Sleep -Seconds 1
            } else {
                Write-Host "`nInspect Mode cancelled" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            $script:isPaused = $false
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
    elseif ($key.Key -eq [System.ConsoleKey]::D) {
        $script:showDetails = -not $script:showDetails
        if ($script:showDetails) {
            Write-Host "`nDetails enabled (showing process file paths)" -ForegroundColor Green
        } else {
            Write-Host "`nDetails disabled" -ForegroundColor Yellow
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
                        $geoInfo = Get-IspInfo -ipAddress $_.RemoteAddress
                        # Get process path
                        $processPath = ""
                        try {
                            # Try Path property first (PowerShell 5.1+)
                            if ($process.Path) {
                                $processPath = $process.Path
                            }
                        } catch {
                            # Path property not available, try MainModule
                            try {
                                $mainModule = $process.MainModule
                                if ($mainModule -and $mainModule.FileName) {
                                    $processPath = $mainModule.FileName
                                }
                            } catch {
                                # MainModule access denied or not available (common for system processes)
                                $processPath = ""
                            }
                        }
                        
                        # If still empty, try alternative method
                        if ([string]::IsNullOrWhiteSpace($processPath)) {
                            try {
                                $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $($process.Id)" -ErrorAction SilentlyContinue
                                if ($proc -and $proc.ExecutablePath) {
                                    $processPath = $proc.ExecutablePath
                                }
                            } catch {
                                # Couldn't get path - leave empty
                            }
                        }
                        
                        [PSCustomObject]@{
                            ProcessName = $process.ProcessName
                            OwningProcess = $_.OwningProcess
                            ProcessPath = $processPath
                            LocalAddress = $_.LocalAddress
                            LocalPort = $_.LocalPort
                            RemoteAddress = $_.RemoteAddress
                            GeoInfo = $geoInfo
                            ISP = if ($geoInfo -is [hashtable]) { $geoInfo.ISP } else { $geoInfo }
                            RemotePort = $_.RemotePort
                            State = $_.State
                        }
                    }
                }
            } | Sort-Object ProcessName, @{Expression={if ($_.GeoInfo -is [hashtable]) { $_.GeoInfo.ISP } else { $_.ISP }}}

            if ($sortedConnections) {
                $counter = 0
                $script:currentConnections = $sortedConnections | ForEach-Object {
                    $counter++
                    [PSCustomObject]@{
                        'Number' = $counter
                        ProcessName = $_.ProcessName
                        OwningProcess = $_.OwningProcess
                        ProcessPath = if ($_.ProcessPath) { $_.ProcessPath } else { "" }
                        LocalAddress = $_.LocalAddress
                        LocalPort = $_.LocalPort
                        RemoteAddress = $_.RemoteAddress
                        GeoInfo = $_.GeoInfo
                        ISP = $_.ISP
                        RemotePort = $_.RemotePort
                        State = $_.State
                    }
                }

                # Update session connections tracking
                Update-SessionConnections -currentConnections $script:currentConnections

                if ($script:focusMode) {
                    Update-FocusConnections -currentConnections $script:currentConnections
                    Display-FocusConnections
                } else {
                    Write-ColoredConnectionsTable -connections $script:currentConnections -highlightActive $false
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