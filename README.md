# NetWatch ![Top Language](https://img.shields.io/github/languages/top/kadmielp/NetWatch) ![Last Commit](https://img.shields.io/github/last-commit/kadmielp/NetWatch) ![License](https://img.shields.io/github/license/kadmielp/NetWatch)



This PowerShell script, `NetWatch`, monitors outbound network connections on your Windows system, providing real-time information about external connections, along with additional features like Inspect and Focus modes.

## Features

- **Real-Time Connection Monitoring:** Displays active external network connections, excluding loopback addresses.
- **IP Geolocation Lookup:** Retrieves ISP information for remote IP addresses.
- **IP Address Caching:** Caches ISP lookups to minimize API calls and improve performance.
- **Pause/Unpause:** Ability to pause and resume connection monitoring.
- **Export Active to CSV:** Export currently active connections to a CSV file.
- **Export Session to CSV:** Export all connections seen during the session (active and inactive) to a CSV file.
- **Save Cache:** Save the cached IP lookups to a JSON file.
- **Inspect Mode:** Inspect connections associated with a specific process.
- **Focus Mode:** Track and display a history of all non-cached connections established during a session, with active and inactive states.
- **Focus Report:** Save a detailed report of the Focus Mode session, including active and inactive connection history.

## How to Use

1.  **Download:** Clone or download the `netwatch.ps1` script to your local machine.
2.  **Run:** Open PowerShell as an administrator and navigate to the directory where the script is saved.
3.  **Execute:** Run the script using:

    ```powershell
    .\netwatch.ps1
    ```
4.  **Interact:**
    -   **`P`**: Pause/Unpause monitoring.
    -   **`A`**: Export currently active connections to a CSV file.
    -   **`E`**: Export all session connections (active and inactive) to a CSV file.
    -   **`I`**: Enter Inspect Mode to view connections for a specific process.
    -   **`F`**: Toggle Focus Mode to track uncached connections.
    -   **`C`**: Save IP cache to `ip_cache.json`.
    -   **`D`**: Toggle showing process file paths.
    -   **`+`** (while in focus mode): Add current view to cache.
    -   **`S`** (while in focus mode): Save focus mode report.

## Export Features

-   **Export Active (`A`):** Exports only the connections that are currently active at the time of export.
-   **Export Session (`E`):** Exports all unique connections that have been seen during the entire session, including both active and inactive connections. Includes timestamps (FirstSeen, LastSeen), status, and connection count.

## Focus Mode Details

-   **Tracking:** Focus mode displays an accumulating history of unique, non-cached connections.
-   **Status:** Connections are marked as "Active" (in normal color) if currently active or "Inactive" (in red) if no longer active.
-   **Reporting:** Focus Mode session reports are automatically saved upon exit, listing all tracked connections along with their active/inactive status.

## Cache Details

-   The script uses `ip-api.com` for IP lookups.
-   Lookups are cached locally in `ip_cache.json` to minimize requests and improve performance.

## Note

-   The script requires internet access to perform IP lookups for external addresses.
-   The script requires administrator privileges to retrieve network connection information.

## License

This project is available under the [MIT License](LICENSE)

## Contributing

Contributions are welcome! Please submit pull requests with improvements or bug fixes.
