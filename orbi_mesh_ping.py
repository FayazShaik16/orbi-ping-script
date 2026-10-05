#!/usr/bin/env python3
"""
Netgear Orbi 771/770 Mesh Testbed - 6 Ping Terminal Orchestrator
----------------------------------------------------------------
Launches 6 terminal tabs / windows on the Main Client (connected to Base Node 771):
1) Base Client -> 8.8.8.8 (Internet) every 100ms
2) Base Client -> Itself (Local Client IP) every 100ms
3) Base Client -> Satellite-1 IP every 100ms
4) Base Client -> Satellite-2 IP every 100ms
5) Satellite-1 (via SSH) -> 8.8.8.8 every 100ms
6) Satellite-2 (via SSH) -> 8.8.8.8 every 100ms

Data Persistence & Excel Architecture:
- Creates a SINGLE master Excel file with 6 distinct sheets.
- Sheet Columns: Time stamp | Hour | Mesh Node Type | PacketLoss | TTL | String
- Instantaneous persistence: Every single ping in every terminal is immediately
  flushed and fsynced to disk in its session journal CSV.
- Simultaneously, a background synchronization worker syncs incoming ping rows
  into the master Excel file in real time using atomic replacement.
- Supports macOS (Terminal.app tabs), Linux (gnome-terminal tabs / tmux), and
  Windows (Windows Terminal wt.exe / cmd).
"""

import argparse
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import threading
import time

from excel_sync import ExcelSyncManager, SHEET_CONFIG, run_standalone_sync

# ANSI Color Codes
COLOR_RESET = "\033[0m"
COLOR_GREEN = "\033[92m"
COLOR_RED = "\033[91m"
COLOR_YELLOW = "\033[93m"
COLOR_CYAN = "\033[96m"
COLOR_BOLD = "\033[1m"


def get_default_local_ip():
    """Attempt to detect local IP routed toward default gateway."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def launch_macos_terminal(commands, titles=None):
    """
    Launches commands in macOS Terminal.app without requiring Accessibility/keystroke permissions.
    """
    print(f"[*] Launching {len(commands)} sessions in macOS Terminal.app...")

    def escape_applescript(cmd):
        return cmd.replace('\\', '\\\\').replace('"', '\\"')

    # Native AppleScript 'do script' creates a terminal window/tab reliably without System Events
    script_lines = [
        'tell application "Terminal"',
        'activate',
    ]
    for cmd in commands:
        c_esc = escape_applescript(cmd)
        script_lines.append(f'do script "{c_esc}"')
    script_lines.append('end tell')

    apple_script = '\n'.join(script_lines)

    try:
        subprocess.run(["osascript", "-e", apple_script], check=True)
    except Exception as e:
        print(f"[!] AppleScript launch error: {e}. Falling back to bash scripts...")
        for cmd in commands:
            subprocess.Popen(["bash", "-c", cmd])


def launch_linux_terminal(commands, titles):
    """
    Launches commands on Linux using gnome-terminal (tabs), tmux, or xterm.
    """
    # 1. Try gnome-terminal with tabs
    if shutil.which("gnome-terminal"):
        print("[*] Launching with gnome-terminal tabs...")
        args = ["gnome-terminal"]
        for idx, (cmd, title) in enumerate(zip(commands, titles)):
            args += ["--tab", f"--title={title}", "--", "bash", "-c", f"{cmd}; exec bash"]
        subprocess.Popen(args)
        return

    # 2. Try tmux
    if shutil.which("tmux"):
        session_name = f"orbi_ping_{int(time.time())}"
        print(f"[*] Launching with tmux session '{session_name}'...")
        subprocess.run(["tmux", "new-session", "-d", "-s", session_name, "-n", titles[0], commands[0]])
        for cmd, title in zip(commands[1:], titles[1:]):
            subprocess.run(["tmux", "new-window", "-t", session_name, "-n", title, cmd])
        print(f"[*] Attach to sessions using: tmux attach -t {session_name}")
        if shutil.which("xterm"):
            subprocess.Popen(["xterm", "-e", f"tmux attach -t {session_name}"])
        return

    # 3. Fallback to individual windows (xfce4-terminal, konsole, or xterm)
    term_bin = shutil.which("xfce4-terminal") or shutil.which("konsole") or shutil.which("xterm")
    if term_bin:
        print(f"[*] Launching windows using {term_bin}...")
        for cmd in commands:
            subprocess.Popen([term_bin, "-e", f"bash -c '{cmd}; exec bash'"])
        return

    print("[!] No supported terminal emulator detected. Run commands manually:")
    for cmd in commands:
        print(f"    {cmd}")


def launch_windows_terminal(commands, titles):
    """
    Launches commands on Windows using Windows Terminal (wt.exe) or cmd.exe.
    """
    wt_path = shutil.which("wt.exe") or shutil.which("wt")
    if wt_path:
        print("[*] Launching with Windows Terminal (wt.exe) tabs...")
        wt_args = ["wt.exe"]
        for idx, (cmd, title) in enumerate(zip(commands, titles)):
            if idx > 0:
                wt_args.append(";")
                wt_args.append("new-tab")
            wt_args += ["--title", f'"{title}"', "cmd.exe", "/k", cmd]
        subprocess.Popen(" ".join(wt_args), shell=True)
        return

    print("[*] Launching standard cmd.exe windows...")
    for cmd, title in zip(commands, titles):
        subprocess.Popen(f'start "{title}" cmd.exe /k {cmd}', shell=True)


def parse_duration(val_str):
    """
    Parses human-friendly duration strings:
    - '2h' or '2 hours' -> 7200 seconds
    - '1.5h' -> 5400 seconds
    - '30m' or '30 mins' -> 1800 seconds
    - '120s' or '120' -> 120 seconds
    - '0' -> continuous
    """
    if val_str is None:
        return 60
    val = str(val_str).strip().lower()
    if not val or val == "0":
        return 0

    if val.endswith("h") or "hour" in val or "hr" in val:
        num = float(re.sub(r"[^\d.]", "", val))
        return int(num * 3600)
    elif val.endswith("m") or "min" in val:
        num = float(re.sub(r"[^\d.]", "", val))
        return int(num * 60)
    elif val.endswith("s") or "sec" in val:
        num = float(re.sub(r"[^\d.]", "", val))
        return int(num)
    else:
        try:
            return int(float(val))
        except ValueError:
            raise argparse.ArgumentTypeError(f"Invalid duration format: '{val_str}'. Use e.g. 2h, 30m, 120s, or 3600.")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Launch 6 Ping Terminals and real-time synchronize results to a 6-sheet Excel file.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    detected_os = platform.system().lower()
    if "darwin" in detected_os:
        default_os = "macos"
    elif "windows" in detected_os:
        default_os = "windows"
    else:
        default_os = "linux"

    parser.add_argument(
        "--os",
        dest="target_os",
        choices=["macos", "linux", "windows"],
        default=default_os,
        help="Operating System of this Main Client machine.",
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=0.1,
        help="Ping interval in seconds (default: 0.1s = 100ms).",
    )
    parser.add_argument(
        "--duration",
        type=parse_duration,
        default="60",
        help="Ping duration. Supports e.g. 2h (2 hours), 30m (30 mins), 120s, or 0 for continuous.",
    )
    parser.add_argument(
        "--sat1-ip",
        type=str,
        default="192.168.1.10",
        help="IP address of Satellite-1 (RBE770).",
    )
    parser.add_argument(
        "--sat2-ip",
        type=str,
        default="192.168.1.11",
        help="IP address of Satellite-2 (RBE770).",
    )
    parser.add_argument(
        "--base-ip",
        type=str,
        default="192.168.1.1",
        help="IP address of Base Node (RBE771).",
    )
    parser.add_argument(
        "--client-ip",
        type=str,
        default=None,
        help="Self IP to ping (defaults to auto-detected local IP).",
    )
    parser.add_argument(
        "--internet-ip",
        type=str,
        default="8.8.8.8",
        help="Internet target to ping.",
    )

    # SSH settings
    parser.add_argument(
        "--ssh-user",
        type=str,
        default="root",
        help="SSH username for Satellite nodes.",
    )
    parser.add_argument(
        "--ssh-port",
        type=int,
        default=22,
        help="SSH port for Satellite nodes.",
    )
    parser.add_argument(
        "--ssh-key",
        type=str,
        default=None,
        help="Path to SSH private key file (optional).",
    )
    parser.add_argument(
        "--ssh-pass",
        type=str,
        default=None,
        help="SSH password (optional).",
    )

    # Output options
    parser.add_argument(
        "--log-dir",
        type=str,
        default="./ping_logs",
        help="Root directory to store timestamped logs and session journals.",
    )
    parser.add_argument(
        "--excel-file",
        type=str,
        default=None,
        help="Custom path for the master Excel file (default: saved inside the session folder).",
    )

    # Recovery option
    parser.add_argument(
        "--sync-excel",
        type=str,
        default=None,
        metavar="JOURNAL_DIR",
        help="Rebuild or synchronize an Excel file from an existing session's CSV journal directory.",
    )

    return parser.parse_args()


def prompt_interactive(args):
    """Interactive console configuration wizard."""
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")
    print(f"   {COLOR_BOLD}NETGEAR ORBI 771/770 MESH TESTBED - 6 PING TERMINALS + EXCEL{COLOR_RESET}")
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")

    print(f"\n[1] Selected OS: {COLOR_YELLOW}{args.target_os.upper()}{COLOR_RESET} (Detected: {platform.system()})")
    choice = input("    Change OS? [m=macOS, l=Linux, w=Windows, Enter to keep]: ").strip().lower()
    if choice == "m":
        args.target_os = "macos"
    elif choice == "l":
        args.target_os = "linux"
    elif choice == "w":
        args.target_os = "windows"

    dur_str = f"{args.duration}s" if args.duration < 3600 else f"{args.duration / 3600:.1f}h ({args.duration}s)"
    dur_input = input(f"\n[2] Ping Duration (e.g. 2h, 30m, 300s, 0=continuous) [{dur_str}]: ").strip()
    if dur_input:
        try:
            args.duration = parse_duration(dur_input)
        except Exception as e:
            print(f"    [!] {e}. Keeping default {args.duration}s.")

    int_input = input(f"\n[3] Ping Interval [{args.interval}s (100ms)]: ").strip()
    if int_input:
        try:
            args.interval = float(int_input)
        except ValueError:
            pass

    sat1_input = input(f"\n[4] Satellite-1 IP [{args.sat1_ip}]: ").strip()
    if sat1_input:
        args.sat1_ip = sat1_input

    sat2_input = input(f"[5] Satellite-2 IP [{args.sat2_ip}]: ").strip()
    if sat2_input:
        args.sat2_ip = sat2_input

    ssh_user_input = input(f"\n[6] Satellite SSH Username [{args.ssh_user}]: ").strip()
    if ssh_user_input:
        args.ssh_user = ssh_user_input

    ssh_pass_input = input(f"[7] Satellite SSH Password (leave blank if SSH key / manual login): ").strip()
    if ssh_pass_input:
        args.ssh_pass = ssh_pass_input

    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}\n")


def live_sync_loop(excel_manager, stop_event, update_interval=1.0):
    """
    Background worker thread that periodically syncs incoming CSV rows
    into the master Excel file atomically.
    """
    while not stop_event.is_set():
        try:
            excel_manager.sync_once()
        except Exception as e:
            sys.stderr.write(f"[LiveSync Error] {e}\n")
        time.sleep(update_interval)

    # Final sync pass on exit
    try:
        excel_manager.finalize()
    except Exception as e:
        sys.stderr.write(f"[Finalize Error] {e}\n")


def format_time_str(seconds):
    """Formats seconds to e.g. 1h 20m 30s or 45s."""
    s = int(seconds)
    if s >= 3600:
        h = s // 3600
        m = (s % 3600) // 60
        sec = s % 60
        return f"{h}h {m}m {sec}s"
    elif s >= 60:
        m = s // 60
        sec = s % 60
        return f"{m}m {sec}s"
    else:
        return f"{s}s"


def print_live_dashboard(excel_manager, duration, start_time):
    """
    Prints a live terminal dashboard showing current row and loss counts across all 6 sheets.
    """
    elapsed = int(time.time() - start_time)
    if duration > 0:
        dur_str = f"{format_time_str(elapsed)} / {format_time_str(duration)}"
    else:
        dur_str = f"{format_time_str(elapsed)} (continuous)"
    print(f"\r{COLOR_BOLD}[Live Status {dur_str}]{COLOR_RESET} ", end="")
    sheet_snippets = []
    for s_name in ["Base_Internet", "Base_Self", "Base_Satellite1", "Base_Satellite2", "Satellite1_Internet", "Satellite2_Internet"]:
        st = excel_manager.stats.get(s_name, {"total": 0, "losses": 0})
        short_name = s_name.replace("Base_", "B_").replace("Satellite", "Sat")
        loss_color = COLOR_RED if st["losses"] > 0 else COLOR_GREEN
        sheet_snippets.append(f"{short_name}: {st['total']}p ({loss_color}{st['losses']}d{COLOR_RESET})")
    print(" | ".join(sheet_snippets), end="", flush=True)


def main():
    args = parse_args()

    # Recovery mode if --sync-excel is passed
    if args.sync_excel:
        target_excel = args.excel_file or os.path.join(args.sync_excel, "recovered_ping_results.xlsx")
        run_standalone_sync(args.sync_excel, target_excel)
        return

    # Interactive mode if launched with no CLI parameters
    if len(sys.argv) == 1:
        prompt_interactive(args)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    local_ip = args.client_ip if args.client_ip else get_default_local_ip()

    # Setup Session and Filepaths
    ts_tag = time.strftime("%Y%m%d_%H%M%S")
    session_dir = os.path.abspath(os.path.join(args.log_dir, f"session_{ts_tag}"))
    os.makedirs(session_dir, exist_ok=True)

    excel_file = args.excel_file or os.path.join(session_dir, f"orbi_mesh_ping_{ts_tag}.xlsx")
    excel_file = os.path.abspath(excel_file)

    # Map sheets to individual CSV journals
    journal_map = {}
    for cfg in SHEET_CONFIG:
        journal_map[cfg["name"]] = os.path.join(session_dir, f"{cfg['name']}.csv")

    print(f"{COLOR_BOLD}{COLOR_GREEN}[+] Initializing Master Excel File with 6 Sheets...{COLOR_RESET}")
    print(f"    Excel File  : {excel_file}")
    print(f"    Session Dir : {session_dir}")

    # Initialize Excel Sync Manager
    excel_manager = ExcelSyncManager(excel_file, journal_map)

    # Define Worker Configurations
    worker_script = os.path.join(base_dir, "ping_worker.py")
    py_exec = sys.executable

    tab_definitions = [
        # Tab 1: Local Base Client -> Internet
        {
            "sheet": "Base_Internet",
            "node_type": "Base",
            "title": "Tab1: Base -> Internet",
            "target": args.internet_ip,
            "ssh_host": None,
        },
        # Tab 2: Local Base Client -> Itself
        {
            "sheet": "Base_Self",
            "node_type": "Base",
            "title": "Tab2: Base -> Self",
            "target": local_ip,
            "ssh_host": None,
        },
        # Tab 3: Local Base Client -> Satellite-1
        {
            "sheet": "Base_Satellite1",
            "node_type": "Base",
            "title": "Tab3: Base -> Satellite1",
            "target": args.sat1_ip,
            "ssh_host": None,
        },
        # Tab 4: Local Base Client -> Satellite-2
        {
            "sheet": "Base_Satellite2",
            "node_type": "Base",
            "title": "Tab4: Base -> Satellite2",
            "target": args.sat2_ip,
            "ssh_host": None,
        },
        # Tab 5: Satellite-1 (via SSH) -> Internet
        {
            "sheet": "Satellite1_Internet",
            "node_type": "Satellite1",
            "title": "Tab5: Sat1 (SSH) -> Internet",
            "target": args.internet_ip,
            "ssh_host": args.sat1_ip,
        },
        # Tab 6: Satellite-2 (via SSH) -> Internet
        {
            "sheet": "Satellite2_Internet",
            "node_type": "Satellite2",
            "title": "Tab6: Sat2 (SSH) -> Internet",
            "target": args.internet_ip,
            "ssh_host": args.sat2_ip,
        },
    ]

    commands = []
    titles = []

    for tab in tab_definitions:
        sheet = tab["sheet"]
        csv_journal = journal_map[sheet]
        cmd = [
            f'"{py_exec}"',
            f'"{worker_script}"',
            f'--sheet "{sheet}"',
            f'--node-type "{tab["node_type"]}"',
            f'--target "{tab["target"]}"',
            f'--interval {args.interval}',
            f'--duration {args.duration}',
            f'--journal "{csv_journal}"',
        ]
        if tab["ssh_host"]:
            cmd += [
                f'--ssh-host "{tab["ssh_host"]}"',
                f'--ssh-user "{args.ssh_user}"',
                f'--ssh-port {args.ssh_port}',
            ]
            if args.ssh_key:
                cmd.append(f'--ssh-key "{args.ssh_key}"')
            if args.ssh_pass:
                cmd.append(f'--ssh-pass "{args.ssh_pass}"')

        full_cmd_str = " ".join(cmd)
        commands.append(full_cmd_str)
        titles.append(tab["title"])

    print(f"\n[*] Launching 6 concurrent terminal tabs on {args.target_os.upper()}...")
    if args.target_os == "macos":
        launch_macos_terminal(commands)
    elif args.target_os == "linux":
        launch_linux_terminal(commands, titles)
    elif args.target_os == "windows":
        launch_windows_terminal(commands, titles)

    # Start Background Excel Sync Thread
    stop_sync_event = threading.Event()
    sync_thread = threading.Thread(
        target=live_sync_loop,
        args=(excel_manager, stop_sync_event, 1.0),
        daemon=True,
    )
    sync_thread.start()

    print(f"{COLOR_BOLD}{COLOR_GREEN}[+] All 6 terminal sessions have been initiated!{COLOR_RESET}")
    print(f"[*] Real-time synchronization to Excel is active.")
    print(f"[*] Press Ctrl+C in this launcher window at any time to finish and finalize.\n")

    start_time = time.time()
    try:
        while True:
            time.sleep(1.0)
            print_live_dashboard(excel_manager, args.duration, start_time)
            if args.duration > 0 and (time.time() - start_time) >= (args.duration + 2):
                break
    except KeyboardInterrupt:
        print(f"\n\n{COLOR_YELLOW}[!] Test run interrupted by user.{COLOR_RESET}")
    finally:
        print(f"\n\n[*] Finalizing Excel sync and ensuring all rows are committed...")
        stop_sync_event.set()
        sync_thread.join(timeout=5.0)

        # Print Final Statistics
        print(f"\n{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")
        print(f"               {COLOR_BOLD}FINAL MESH PING TEST RESULTS{COLOR_RESET}")
        print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")
        print(f"{'Sheet Name':<25} {'Mesh Node':<14} {'Total Pings':<14} {'Losses':<10} {'Loss Hour(s)':<22}")
        print("-" * 88)
        for cfg in SHEET_CONFIG:
            s_name = cfg["name"]
            st = excel_manager.stats.get(s_name, {"total": 0, "losses": 0, "loss_hours": set()})
            loss_str = f"{COLOR_RED}{st['losses']}{COLOR_RESET}" if st["losses"] > 0 else f"{COLOR_GREEN}0{COLOR_RESET}"
            loss_hrs = sorted(st.get("loss_hours", []))
            hrs_str = ", ".join(f"Hour {h+1}" for h in loss_hrs) if loss_hrs else "None"
            hrs_color = f"{COLOR_RED}{hrs_str}{COLOR_RESET}" if loss_hrs else f"{COLOR_GREEN}None{COLOR_RESET}"
            print(f"{s_name:<25} {cfg['node_type']:<14} {st['total']:<14} {loss_str:<10} {hrs_color:<22}")
        print("-" * 88)

        print(f"\n{COLOR_BOLD}{COLOR_GREEN}[✓] Single Master Excel File:{COLOR_RESET} {excel_file}")
        print(f"{COLOR_BOLD}[✓] Raw Flushed CSV Journals:{COLOR_RESET} {session_dir}\n")


if __name__ == "__main__":
    main()
