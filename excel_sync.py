#!/usr/bin/env python3
"""
Excel Synchronization Manager for Orbi Mesh Ping Testbed
---------------------------------------------------------
Creates and maintains a single Excel (.xlsx) file with 6 distinct sheets.
Continuously reads append-only CSV journals from the 6 concurrent terminal
streams and synchronizes rows into the Excel sheets in real time.

Features:
- Crash-safe atomic writing (writes to temp file and atomically swaps)
- Graceful handling of Excel file locks if opened by user
- Professional styling (navy header, bold white text, column widths, freeze pane)
- Standalone CLI mode to reconstruct or repair the Excel file from CSV journals
"""

import argparse
import csv
import os
import shutil
import sys
import tempfile
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SHEET_CONFIG = [
    {"name": "Base_Internet", "node_type": "Base", "title": "Base Node Client -> Internet (8.8.8.8)"},
    {"name": "Base_Self", "node_type": "Base", "title": "Base Node Client -> Self (Local IP)"},
    {"name": "Base_Satellite1", "node_type": "Base", "title": "Base Node Client -> Satellite-1 IP"},
    {"name": "Base_Satellite2", "node_type": "Base", "title": "Base Node Client -> Satellite-2 IP"},
    {"name": "Satellite1_Internet", "node_type": "Satellite1", "title": "Satellite-1 (SSH) -> Internet (8.8.8.8)"},
    {"name": "Satellite2_Internet", "node_type": "Satellite2", "title": "Satellite-2 (SSH) -> Internet (8.8.8.8)"},
]

HEADER_COLUMNS = ["Time stamp", "Mesh Node Type", "PacketLoss", "TTL", "String"]


class ExcelSyncManager:
    def __init__(self, excel_path, journal_map):
        """
        :param excel_path: Target .xlsx filepath
        :param journal_map: Dict mapping sheet_name -> csv_journal_path
        """
        self.excel_path = os.path.abspath(excel_path)
        self.journal_map = journal_map
        self.offsets = {sheet: 0 for sheet in journal_map}
        self.wb = None
        self.stats = {sheet: {"total": 0, "losses": 0} for sheet in journal_map}
        self.init_workbook()

    def init_workbook(self):
        """Initializes the Excel workbook with 6 formatted sheets."""
        os.makedirs(os.path.dirname(self.excel_path), exist_ok=True)

        if os.path.exists(self.excel_path):
            try:
                self.wb = openpyxl.load_workbook(self.excel_path)
            except Exception:
                self.wb = self._create_fresh_workbook()
        else:
            self.wb = self._create_fresh_workbook()

        self._ensure_sheets_exist()
        self._atomic_save()

    def _create_fresh_workbook(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default 'Sheet'
        return wb

    def _ensure_sheets_exist(self):
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        label_font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
        val_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )

        for cfg in SHEET_CONFIG:
            s_name = cfg["name"]
            if s_name in self.wb.sheetnames:
                # If sheet exists with old layout (A1 is not SheetName (TestCase):), remove and recreate
                if self.wb[s_name]["A1"].value != "SheetName (TestCase):":
                    self.wb.remove(self.wb[s_name])

            if s_name not in self.wb.sheetnames:
                ws = self.wb.create_sheet(title=s_name)

                # Summary Row 1: SheetName (TestCase)
                ws["A1"] = "SheetName (TestCase):"
                ws["B1"] = f"{s_name} ({cfg['title']})"
                ws["A1"].font = label_font
                ws["B1"].font = val_font

                # Summary Row 2: Packet Loss Percentage & Packet Loss Count
                ws["A2"] = "Packet Loss :"
                ws["B2"] = "0.00%"
                ws["C2"] = "Packet Loss Count :"
                ws["D2"] = 0
                ws["A2"].font = label_font
                ws["B2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")
                ws["C2"].font = label_font
                ws["D2"].font = val_font

                # Row 3 is a blank row separator
                ws.append([])

                # Row 4: Data Table Headers
                ws.append(HEADER_COLUMNS)

                # Format Row 4 header cells
                for col_idx in range(1, len(HEADER_COLUMNS) + 1):
                    cell = ws.cell(row=4, column=col_idx)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center" if col_idx in [2, 3, 4] else "left")

                ws.column_dimensions["A"].width = 25
                ws.column_dimensions["B"].width = 18
                ws.column_dimensions["C"].width = 15
                ws.column_dimensions["D"].width = 12
                ws.column_dimensions["E"].width = 65
                ws.freeze_panes = "A5"

    def sync_once(self):
        """
        Polls each journal CSV, reads newly appended lines, appends to the
        respective sheet, updates the 2-row summary, and atomically saves the Excel file.
        Returns the number of new rows added in this cycle.
        """
        total_new_rows = 0

        for sheet_name, csv_path in self.journal_map.items():
            if not os.path.exists(csv_path):
                continue

            ws = self.wb[sheet_name]
            current_offset = self.offsets[sheet_name]

            try:
                with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
                    f.seek(current_offset)
                    # If this is the very first read from offset 0, skip CSV header line
                    if current_offset == 0:
                        first_line = f.readline()
                        if "Time stamp" in first_line:
                            pass
                        else:
                            # Not a header, rewind
                            f.seek(0)

                    reader = csv.reader(f)
                    for row in reader:
                        if not row or len(row) < 4:
                            continue

                        # Parse row: supports 5-col [timestamp, node, loss, ttl, string]
                        # or 4-col legacy [timestamp, node, loss, string]
                        if len(row) >= 5:
                            ts_val = row[0]
                            node_val = row[1]
                            loss_val = row[2]
                            ttl_val = row[3]
                            str_val = row[4]
                        else:
                            ts_val = row[0]
                            node_val = row[1]
                            loss_val = row[2]
                            str_val = row[3]
                            import re
                            m = re.search(r'\bttl[=\s:]+(\d+)', str_val, re.IGNORECASE)
                            ttl_val = int(m.group(1)) if m else "NA"

                        # Try to format TTL as int if numeric
                        try:
                            ttl_out = int(ttl_val)
                        except (ValueError, TypeError):
                            ttl_out = ttl_val

                        ws.append([ts_val, node_val, loss_val, ttl_out, str_val])
                        total_new_rows += 1

                        # Update in-memory stats
                        self.stats[sheet_name]["total"] += 1
                        loss_clean = str(loss_val).strip().upper()
                        if loss_clean != "NA" and loss_clean != "":
                            self.stats[sheet_name]["losses"] += 1

                    self.offsets[sheet_name] = f.tell()

                # Update the 2-row summary at the top of the sheet
                total_cnt = self.stats[sheet_name]["total"]
                loss_cnt = self.stats[sheet_name]["losses"]
                loss_pct = (loss_cnt / total_cnt * 100.0) if total_cnt > 0 else 0.0

                ws["B2"].value = f"{loss_pct:.2f}%"
                ws["D2"].value = loss_cnt

                if loss_cnt > 0:
                    ws["B2"].font = Font(name="Calibri", size=11, bold=True, color="C0392B")
                    ws["D2"].font = Font(name="Calibri", size=11, bold=True, color="C0392B")
                else:
                    ws["B2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")
                    ws["D2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")

            except Exception as e:
                sys.stderr.write(f"[Sync Warning] Reading {csv_path}: {e}\n")

        if total_new_rows > 0:
            self._atomic_save()

        return total_new_rows

    def _atomic_save(self):
        """
        Saves the workbook to a temporary file first, then atomically replaces
        the target file. Prevents corrupt files if interrupted mid-save.
        """
        temp_dir = os.path.dirname(self.excel_path)
        temp_file = os.path.join(temp_dir, f".tmp_{os.path.basename(self.excel_path)}_{os.getpid()}")
        try:
            self.wb.save(temp_file)
            os.replace(temp_file, self.excel_path)
        except PermissionError:
            # File may be locked by user opening it in Microsoft Excel
            pass
        except Exception as e:
            sys.stderr.write(f"[Save Error] Could not save Excel file: {e}\n")
        finally:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass

    def finalize(self):
        """Performs a thorough final sync and logs summary."""
        self.sync_once()
        self._atomic_save()
        return self.stats


def run_standalone_sync(journal_dir, excel_path):
    """Rebuilds or synchronizes the Excel file from an existing session's CSV journals."""
    journal_map = {}
    for cfg in SHEET_CONFIG:
        csv_file = os.path.join(journal_dir, f"{cfg['name']}.csv")
        if os.path.exists(csv_file):
            journal_map[cfg["name"]] = csv_file

    if not journal_map:
        print(f"[!] No journal CSV files found in {journal_dir}")
        return

    print(f"[*] Synchronizing {len(journal_map)} sheets into: {excel_path}...")
    manager = ExcelSyncManager(excel_path, journal_map)
    manager.sync_once()
    stats = manager.finalize()

    print("[+] Synchronization complete!\n")
    print(f"{'Sheet Name':<25} {'Total Pings':<15} {'Packet Losses':<15}")
    print("-" * 55)
    for s_name, data in stats.items():
        print(f"{s_name:<25} {data['total']:<15} {data['losses']:<15}")
    print("-" * 55)


def main():
    parser = argparse.ArgumentParser(description="Synchronize Orbi Mesh CSV journals to Master Excel file.")
    parser.add_argument("--journal-dir", required=True, help="Directory containing sheet CSV files")
    parser.add_argument("--excel-file", required=True, help="Target Excel (.xlsx) file")
    args = parser.parse_args()

    run_standalone_sync(args.journal_dir, args.excel_file)


if __name__ == "__main__":
    main()
