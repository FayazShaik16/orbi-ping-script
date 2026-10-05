#!/usr/bin/env python3
"""
Ping Worker for Orbi Mesh Testbed
---------------------------------
Executes high-frequency pings (default: 100ms) either locally or via SSH,
streams live output to terminal stdout, tracks sequential packet loss counters,
and immediately flushes & fsyncs every row into a session CSV journal.
"""

import argparse
import csv
from datetime import datetime
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import time


# ANSI Terminal Colors
COLOR_RESET = "\033[0m"
COLOR_GREEN = "\033[92m"
COLOR_RED = "\033[91m"
COLOR_YELLOW = "\033[93m"
COLOR_CYAN = "\033[96m"
COLOR_BOLD = "\033[1m"


def extract_ttl(text):
    """Extracts integer TTL from ping response line, or returns 'NA' if absent/dropped."""
    if not text:
        return "NA"
    m = re.search(r'\bttl[=\s:]+(\d+)', text, re.IGNORECASE)
    return int(m.group(1)) if m else "NA"


def write_journal_row(journal_file, timestamp_str, node_type, loss_str, ttl_val, ping_str):
    """
    Appends a row to the journal CSV and immediately forces disk sync (flush + fsync).
    Ensures 0 data loss even in the event of an abrupt system power loss or process kill.
    """
    try:
        with open(journal_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([timestamp_str, node_type, loss_str, ttl_val, ping_str])
            f.flush()
            os.fsync(f.fileno())
    except Exception as e:
        sys.stderr.write(f"[Journal Error] Could not write row: {e}\n")


def run_local_ping_worker(sheet_name, node_type, target, interval, duration, journal_file):
    """
    Executes local pings from Main Client to target at precise intervals.
    """
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")
    print(f" {COLOR_BOLD}ORBI MESH PING WORKER - LOCAL STREAM{COLOR_RESET}")
    print(f" Sheet Name     : {COLOR_YELLOW}{sheet_name}{COLOR_RESET}")
    print(f" Mesh Node Type : {COLOR_YELLOW}{node_type}{COLOR_RESET}")
    print(f" Target         : {COLOR_GREEN}{target}{COLOR_RESET}")
    print(f" Interval       : {interval * 1000:.0f} ms")
    print(f" Duration       : {f'{duration}s' if duration > 0 else 'Continuous (Ctrl+C to stop)'}")
    print(f" Journal Path   : {journal_file}")
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}\n")

    os_type = platform.system().lower()
    start_time = time.time()
    transmitted = 0
    received = 0
    loss_counter = 0

    try:
        while True:
            now = time.time()
            if duration > 0 and (now - start_time) >= duration:
                break

            transmitted += 1
            ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

            # Choose single ping command based on OS
            if "windows" in os_type:
                cmd = ["ping", "-n", "1", "-w", "1000", target]
            elif "darwin" in os_type:
                cmd = ["ping", "-c", "1", "-W", "1000", target]
            else:
                cmd = ["ping", "-c", "1", "-W", "1", target]

            t_send = time.time()
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            elapsed_ms = (time.time() - t_send) * 1000.0

            # Extract raw output string
            stdout_clean = proc.stdout.strip()
            stderr_clean = proc.stderr.strip()

            is_success = (proc.returncode == 0)
            # Find relevant reply line if multiple lines returned
            reply_line = ""
            for line in stdout_clean.splitlines():
                if any(k in line.lower() for k in ["bytes from", "reply from", "ttl="]):
                    reply_line = line.strip()
                    break

            if not reply_line:
                reply_line = stdout_clean.replace("\r", " ").replace("\n", " | ") if stdout_clean else stderr_clean
            if not reply_line:
                reply_line = f"Ping {target} returned code {proc.returncode}"

            ttl_val = extract_ttl(reply_line) if is_success else "NA"

            if is_success:
                received += 1
                loss_field = "NA"
                display_str = f"[{ts_str}] Reply from {target}: seq={transmitted} time={elapsed_ms:.1f}ms TTL={ttl_val} (Loss: NA)"
                print(f"{COLOR_GREEN}{display_str}{COLOR_RESET}")
            else:
                loss_counter += 1
                loss_field = str(loss_counter)
                display_str = f"[{ts_str}] Request timed out / drop (seq={transmitted}) [Packet Loss #{loss_counter}] TTL=NA"
                print(f"{COLOR_RED}{display_str}{COLOR_RESET}")

            # Commit to journal immediately
            write_journal_row(journal_file, ts_str, node_type, loss_field, ttl_val, reply_line)

            # High precision interval sleep
            spent = time.time() - t_send
            sleep_rem = interval - spent
            if sleep_rem > 0:
                time.sleep(sleep_rem)

    except KeyboardInterrupt:
        print(f"\n{COLOR_YELLOW}[!] Interrupted by user.{COLOR_RESET}")
    finally:
        total_time = time.time() - start_time
        loss_pct = ((transmitted - received) / transmitted * 100.0) if transmitted > 0 else 0.0
        print(f"\n{COLOR_BOLD}--- {sheet_name} Summary ---{COLOR_RESET}")
        print(f"Transmitted: {transmitted} | Received: {received} | Lost: {loss_counter} ({loss_pct:.1f}%) | Time: {total_time:.1f}s")
        print("Journal saved. Press Enter to exit.")
        try:
            input()
        except Exception:
            pass


def run_ssh_ping_worker(sheet_name, node_type, ssh_host, ssh_user, ssh_port, ssh_key, ssh_pass, target, interval, duration, journal_file):
    """
    Executes remote pings from Satellite node to target over SSH, streaming and timestamping
    every line on the local client machine, tracking packet loss, and writing to journal.
    """
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}")
    print(f" {COLOR_BOLD}ORBI MESH PING WORKER - REMOTE SSH STREAM{COLOR_RESET}")
    print(f" Sheet Name     : {COLOR_YELLOW}{sheet_name}{COLOR_RESET}")
    print(f" Mesh Node Type : {COLOR_YELLOW}{node_type}{COLOR_RESET}")
    print(f" SSH Target     : {COLOR_GREEN}{ssh_user}@{ssh_host}:{ssh_port}{COLOR_RESET}")
    print(f" Remote Ping    : {COLOR_GREEN}{target}{COLOR_RESET} every {interval * 1000:.0f} ms")
    print(f" Duration       : {f'{duration}s' if duration > 0 else 'Continuous (Ctrl+C to stop)'}")
    print(f" Journal Path   : {journal_file}")
    print(f"{COLOR_BOLD}{COLOR_CYAN}================================================================{COLOR_RESET}\n")

    count_param = f"-c {int(duration / interval)}" if duration > 0 else ""
    remote_cmd = f"ping -i {interval} {count_param} {target}"

    ssh_args = [
        "ssh",
        "-p", str(ssh_port),
        "-o", "StrictHostKeyChecking=no",
        "-o", "UserKnownHostsFile=/dev/null",
        "-o", "ConnectTimeout=10",
    ]
    if ssh_key:
        ssh_args += ["-i", ssh_key]

    target_str = f"{ssh_user}@{ssh_host}" if ssh_user else ssh_host

    # Check sshpass
    if ssh_pass and shutil.which("sshpass"):
        full_cmd = ["sshpass", "-p", ssh_pass] + ssh_args + [target_str, remote_cmd]
    else:
        full_cmd = ssh_args + [target_str, remote_cmd]

    print(f"[*] Connecting to {target_str}...")
    start_time = time.time()
    transmitted = 0
    received = 0
    loss_counter = 0

    try:
        proc = subprocess.Popen(
            full_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
            universal_newlines=True,
        )

        for raw_line in iter(proc.stdout.readline, ""):
            line = raw_line.strip()
            if not line:
                continue

            ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            lower_line = line.lower()

            # Ignore initial ping header line e.g. "PING 8.8.8.8 (8.8.8.8): 56 data bytes"
            if lower_line.startswith("ping ") and "bytes" in lower_line:
                print(f"[{ts_str}] (SSH Remote) {line}")
                continue

            # Summary footer lines
            if "packets transmitted" in lower_line or "round-trip" in lower_line or "rtt min" in lower_line:
                print(f"[{ts_str}] (SSH Summary) {line}")
                continue

            transmitted += 1

            # Check if this line represents successful ping
            is_success = ("bytes from" in lower_line or "ttl=" in lower_line or "time=" in lower_line)

            ttl_val = extract_ttl(line) if is_success else "NA"

            if is_success:
                received += 1
                loss_field = "NA"
                display_str = f"[{ts_str}] Reply from {target}: seq={transmitted} {line} TTL={ttl_val} (Loss: NA)"
                print(f"{COLOR_GREEN}{display_str}{COLOR_RESET}")
            else:
                loss_counter += 1
                loss_field = str(loss_counter)
                display_str = f"[{ts_str}] Packet Loss #{loss_counter}: {line} TTL=NA"
                print(f"{COLOR_RED}{display_str}{COLOR_RESET}")

            # Append to journal immediately with fsync
            write_journal_row(journal_file, ts_str, node_type, loss_field, ttl_val, line)

        proc.wait()
        if proc.returncode != 0:
            err = proc.stderr.read().strip()
            if err:
                ts_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                loss_counter += 1
                print(f"{COLOR_RED}[{ts_str}] SSH Error: {err}{COLOR_RESET}")
                write_journal_row(journal_file, ts_str, node_type, str(loss_counter), "NA", f"SSH Error: {err}")

    except KeyboardInterrupt:
        print(f"\n{COLOR_YELLOW}[!] Interrupted by user.{COLOR_RESET}")
    finally:
        total_time = time.time() - start_time
        loss_pct = ((transmitted - received) / transmitted * 100.0) if transmitted > 0 else 0.0
        print(f"\n{COLOR_BOLD}--- {sheet_name} Summary ---{COLOR_RESET}")
        print(f"Transmitted: {transmitted} | Received: {received} | Lost: {loss_counter} ({loss_pct:.1f}%) | Time: {total_time:.1f}s")
        print("Journal saved. Press Enter to exit.")
        try:
            input()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(description="Worker process for Orbi Mesh Ping Terminal")
    parser.add_argument("--sheet", required=True, help="Target Sheet name in Excel")
    parser.add_argument("--node-type", required=True, choices=["Base", "Satellite1", "Satellite2"], help="Mesh Node Type")
    parser.add_argument("--target", required=True, help="IP to ping (e.g. 8.8.8.8)")
    parser.add_argument("--interval", type=float, default=0.1, help="Ping interval in seconds")
    parser.add_argument("--duration", type=int, default=60, help="Duration in seconds (0=continuous)")
    parser.add_argument("--journal", required=True, help="Path to CSV journal file")

    # SSH options
    parser.add_argument("--ssh-host", default=None, help="Remote SSH host if remote ping")
    parser.add_argument("--ssh-user", default="root", help="SSH username")
    parser.add_argument("--ssh-port", type=int, default=22, help="SSH port")
    parser.add_argument("--ssh-key", default=None, help="SSH private key path")
    parser.add_argument("--ssh-pass", default=None, help="SSH password")

    args = parser.parse_args()

    # Ensure journal file directory exists and header is written if empty
    os.makedirs(os.path.dirname(os.path.abspath(args.journal)), exist_ok=True)
    if not os.path.exists(args.journal) or os.path.getsize(args.journal) == 0:
        with open(args.journal, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Time stamp", "Mesh Node Type", "PacketLoss", "TTL", "String"])
            f.flush()
            os.fsync(f.fileno())

    if args.ssh_host:
        run_ssh_ping_worker(
            sheet_name=args.sheet,
            node_type=args.node_type,
            ssh_host=args.ssh_host,
            ssh_user=args.ssh_user,
            ssh_port=args.ssh_port,
            ssh_key=args.ssh_key,
            ssh_pass=args.ssh_pass,
            target=args.target,
            interval=args.interval,
            duration=args.duration,
            journal_file=args.journal,
        )
    else:
        run_local_ping_worker(
            sheet_name=args.sheet,
            node_type=args.node_type,
            target=args.target,
            interval=args.interval,
            duration=args.duration,
            journal_file=args.journal,
        )


if __name__ == "__main__":
    main()
