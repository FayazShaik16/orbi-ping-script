# Netgear Orbi 771/770 Mesh Testbed - 6 Ping Terminal Launcher & Live Excel Reporter

Cross-platform testbed orchestration tool that opens **6 concurrent terminal sessions** (tabs/windows) on the **Main Client** (connected to the Base Node RBE771 via LAN cable) and continuously records data into a **single Master Excel workbook with 6 distinct sheets**.

---

## Architecture Overview

```
                      [ INTERNET (8.8.8.8) ]
                                |
                   +--------------------------+
                   |   Base Node (Orbi 771)   |
                   +--------------------------+
                     /                      \
      (Wireless MLO Backhaul)        (Wireless MLO Backhaul)
                   /                          \
+-------------------------+        +-------------------------+
|  Satellite 1 (Orbi 770) |        |  Satellite 2 (Orbi 770) |
+-------------------------+        +-------------------------+
             |                                  |
         (LAN Cable)                        (LAN Cable)
             |                                  |
     [ Client 2 Laptop ]                [ Client 3 Laptop ]

===============================================================
                MAIN CLIENT (Connected to Base Node)
===============================================================
  Tab 1: Base Client -> 8.8.8.8 (Internet)
  Tab 2: Base Client -> Base Node IP (10.168.168.1)
  Tab 3: Base Client -> Satellite-1 IP (10.168.168.196)
  Tab 4: Base Client -> Satellite-3 IP (10.168.168.40)
  Tab 5: PC 2 (wired to Sat-1, via SSH) -> 8.8.8.8 (Internet)
  Tab 6: PC 3 (wired to Sat-3, via SSH) -> 8.8.8.8 (Internet)
===============================================================
             || (Simultaneous & Real-Time Sync)
             \/
[ Master Excel File: orbi_mesh_ping_<timestamp>.xlsx ]
  - Sheet 1: Base_Internet
  - Sheet 2: Base_Self
  - Sheet 3: Base_Satellite1
  - Sheet 4: Base_Satellite2
  - Sheet 5: Satellite1_Internet
  - Sheet 6: Satellite2_Internet
```

---

## Excel Structure & Exact Layout

### 2-Row Sheet Header Summary:
Each sheet features an executive summary at the very top (pinned via freeze panes):
- **Row 1**: `SheetName (TestCase): <SheetName> (<TestCase Details>)`
- **Row 2**: `Packet Loss : <Percentage>%` &nbsp;|&nbsp; `Packet Loss Count : <Total Loss Count>` &nbsp;|&nbsp; `Packet Loss Hour(s) : <Hour(s) with Loss>` (or `None`)
*(Updates dynamically in real time as pings arrive; highlights in green if None, bold red if losses occurred)*

### Column Headers (Row 4) & Excel AutoFilter:
Excel **AutoFilter** is automatically enabled across row 4 headers (`A4:F`), enabling 1-click dropdown filtering by hour, node type, packet loss, or TTL.

| Column Name | Description | Example |
| :--- | :--- | :--- |
| **`Time stamp`** | Millisecond-accurate timestamp | `2026-10-05 18:25:01.345` |
| **`Hour`** | Test hour window identifier for 1-click Excel filtering | `Hour 1`, `Hour 2`... |
| **`Mesh Node Type`** | Node context (`Base`, `Satellite1`, `Satellite2`) | `Base` |
| **`PacketLoss`** | `0` if ping reply received; incrementing integer counter if timed out / dropped | `0` or `1`, `2`, `3`... |
| **`TTL`** | Time-To-Live integer value from the ping response (`NA` on failure) | `117` or `64` |
| **`String`** | Full raw ping output string returned | `64 bytes from 8.8.8.8: seq=1 ttl=117 time=14.2 ms` |

### Hourly Separator Banner Rows:
At the start of every 1-hour window, a styled divider banner row is inserted across the data table:
- **Visual Styling**: Dark navy fill (`#1F4E79`), bold white text, distinct borders.
- **Content**: `=== HOUR <N> (<Start_Time> - <End_Time>) ===`
- **Filter-Compatible**: Unmerged cells ensure full compatibility with Excel AutoFilter, sorting, and keyboard navigation.

### Hourly Breakdown Table & Bar Graph (Bottom of Sheet):
At the bottom of each sheet, a 1-hour breakdown table and a native Excel Bar Chart are automatically generated:
- **Hourly Breakdown Table**:
  - `Hour Window` (e.g. `Hour 1 (18:00 - 19:00)`, `Hour 2 (19:00 - 20:00)`)
  - `Total Pings`
  - `Loss Count`
  - `Avg Packet Loss (%)`
- **Native Excel Bar Graph**:
  - Visualizes the average packet loss percentage for each 1-hour interval.
  - **Fixed Y-Axis Scale**: Ranged from **0% to 100%** with **10%** major step intervals (`0%`, `10%`, `20%` ... `100%`) for standardized, consistent visual comparisons across all sheets and tests.
  - Placed right below the hourly table for executive reporting and visual analysis.

---

## Instantaneous Persistence Guarantee (Crash-Proof)
To protect against system crashes, power loss, or sudden disconnects:
1. **Direct Disk Fsyncing**: Every terminal immediately writes each ping row to a session journal CSV with `flush()` and `os.fsync()`. Even if the power is cut at any millisecond, 100% of data up to that exact instant is physically committed to storage.
2. **Atomic Excel Saving**: A background sync worker updates the master Excel file in real-time every second using atomic file replacement (`.tmp` swap to `.xlsx`). The `.xlsx` file is never corrupted mid-write.
3. **Emergency Recovery Command**: If the machine reboots mid-test, you can regenerate/verify the full Excel workbook from the crash journals at any time:
   ```bash
   python3 orbi_mesh_ping.py --sync-excel ./ping_logs/session_<timestamp>
   ```

---

## Multi-OS Terminal Support
- **macOS**: Opens 6 tabs in `Terminal.app` via AppleScript.
- **Linux**: Opens 6 tabs in `gnome-terminal`, or creates a `tmux` session with 6 windows, or `xterm`.
- **Windows**: Opens 6 tabs in **Windows Terminal** (`wt.exe`) or separate `cmd.exe` windows.

---

## How to Run

### 1. Interactive Mode
Run without arguments to configure parameters step-by-step:
```bash
python3 orbi_mesh_ping.py
```

### 2. Command-Line Execution
```bash
# Run for 2 hours:
python3 orbi_mesh_ping.py --duration 2h

# Run for 1.5 hours with 100ms interval:
python3 orbi_mesh_ping.py --duration 1.5h --interval 0.1

# Run for 30 minutes:
python3 orbi_mesh_ping.py --duration 30m

# Run continuously until Ctrl+C:
python3 orbi_mesh_ping.py --duration 0
```

### Full Example with Options:
```bash
python orbi_mesh_ping.py \
  --os windows \
  --duration 2h \
  --interval 0.1 \
  --base-ip 10.168.168.1 \
  --sat1-ip 10.168.168.196 \
  --sat2-ip 10.168.168.40 \
  --client2-ip 10.168.168.49 \
  --client3-ip 10.168.168.188 \
  --ssh-user Administrator \
  --log-dir ./ping_logs
```

### All CLI Options:
```
  --os {macos,linux,windows}   Operating System of Main Client (default: auto-detected)
  --interval INTERVAL          Ping interval in seconds (default: 0.1 = 100ms)
  --duration DURATION          Ping duration: e.g. 2h (2 hours), 1.5h, 30m, 120s, or 0 (continuous)
  --sat1-ip SAT1_IP            IP address of Satellite-1 (default: 10.168.168.196)
  --sat2-ip SAT2_IP            IP address of Satellite-3 (default: 10.168.168.40)
  --base-ip BASE_IP            IP address of Base Node (default: 10.168.168.1)
  --client2-ip CLIENT2_IP      IP of PC 2 wired to Sat-1 (default: 10.168.168.49)
  --client3-ip CLIENT3_IP      IP of PC 3 wired to Sat-3 (default: 10.168.168.188)
  --client-ip CLIENT_IP        Self IP to ping (defaults to auto-detected local IP)
  --internet-ip INTERNET_IP    Internet target to ping (default: 8.8.8.8)
  --ssh-user SSH_USER          SSH username for remote client PCs
  --ssh-port SSH_PORT          SSH port (default: 22)
  --ssh-key SSH_KEY            Path to SSH private key file
  --ssh-pass SSH_PASS          SSH password (optional)
  --ssh-remote-os {windows,linux} Remote OS of clients (default: windows)
  --log-dir LOG_DIR            Folder to store session journals and Excel reports
  --excel-file EXCEL_FILE      Custom filename/path for output Excel file
  --sync-excel JOURNAL_DIR     Rebuild/recover Excel report from an existing journal folder
```
