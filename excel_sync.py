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
from datetime import datetime, timedelta
import os
import shutil
import sys
import tempfile
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
try:
    from openpyxl.drawing.image import Image as OpenpyxlImage
    from PIL import Image as PILImage, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

SHEET_CONFIG = [
    {"name": "Base_Internet", "node_type": "Base", "title": "Base Node Client -> Internet (8.8.8.8)"},
    {"name": "Base_Self", "node_type": "Base", "title": "Base Node Client -> Base Node IP (10.168.168.1)"},
    {"name": "Base_Satellite1", "node_type": "Base", "title": "Base Node Client -> Satellite-1 IP"},
    {"name": "Base_Satellite2", "node_type": "Base", "title": "Base Node Client -> Satellite-2 IP"},
    {"name": "Satellite1_Internet", "node_type": "Satellite1", "title": "Satellite-1 (SSH) -> Internet (8.8.8.8)"},
    {"name": "Satellite2_Internet", "node_type": "Satellite2", "title": "Satellite-2 (SSH) -> Internet (8.8.8.8)"},
]

HEADER_COLUMNS = ["Time stamp", "Hour", "Mesh Node Type", "PacketLoss", "TTL", "String"]


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
        self.stats = {sheet: {"total": 0, "losses": 0, "loss_hours": set()} for sheet in journal_map}
        self.sheet_t0 = {}
        self.sheet_current_hour = {}

        # Pre-allocate reusable styles for high performance
        self.align_center = Alignment(horizontal="center", vertical="center")
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.font_loss = Font(name="Calibri", size=11, bold=True, color="C0392B")

        self.sep_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        self.sep_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.sep_border = Border(
            left=Side(style="thin", color="1B365D"),
            right=Side(style="thin", color="1B365D"),
            top=Side(style="medium", color="0D1E3A"),
            bottom=Side(style="medium", color="0D1E3A"),
        )

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
                ws_test = self.wb[s_name]
                # If sheet exists with old layout (A1 is not SheetName:, B4 is not Hour, or E2 is not Packet Loss Hour(s) :), remove and recreate
                if ws_test["A1"].value != "SheetName (TestCase):" or ws_test["B4"].value != "Hour" or ws_test["E2"].value != "Packet Loss Hour(s) :":
                    self.wb.remove(ws_test)

            if s_name not in self.wb.sheetnames:
                ws = self.wb.create_sheet(title=s_name)

                # Summary Row 1: SheetName (TestCase)
                ws["A1"] = "SheetName (TestCase):"
                ws["B1"] = f"{s_name} ({cfg['title']})"
                ws["A1"].font = label_font
                ws["B1"].font = val_font

                # Summary Row 2: Packet Loss Percentage, Packet Loss Count & Packet Loss Hour(s)
                ws["A2"] = "Packet Loss :"
                ws["B2"] = "0.00%"
                ws["C2"] = "Packet Loss Count :"
                ws["D2"] = 0
                ws["E2"] = "Packet Loss Hour(s) :"
                ws["F2"] = "None"
                ws["A2"].font = label_font
                ws["B2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")
                ws["C2"].font = label_font
                ws["D2"].font = val_font
                ws["E2"].font = label_font
                ws["F2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")

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
                    cell.alignment = Alignment(horizontal="center" if col_idx in [2, 3, 4, 5] else "left")

                ws.column_dimensions["A"].width = 25
                ws.column_dimensions["B"].width = 14
                ws.column_dimensions["C"].width = 18
                ws.column_dimensions["D"].width = 15
                ws.column_dimensions["E"].width = 22
                ws.column_dimensions["F"].width = 65
                ws.column_dimensions["G"].width = 3
                ws.freeze_panes = "A5"
                ws.auto_filter.ref = "A4:F4"

    def _detect_current_hour(self, ws):
        """Scans backward from ws.max_row to detect the last hour index present in column B."""
        if ws.max_row >= 5:
            for r in range(ws.max_row, 4, -1):
                val = ws.cell(row=r, column=2).value
                if val and isinstance(val, str) and val.startswith("Hour "):
                    try:
                        return int(val.replace("Hour ", "").strip()) - 1
                    except ValueError:
                        pass
        return -1

    def _get_sheet_t0(self, csv_path):
        """Retrieves start timestamp from the first data row of the journal CSV."""
        if not os.path.exists(csv_path):
            return None
        try:
            with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
                reader = csv.reader(f)
                first_row = next(reader, None)
                if first_row and "Time stamp" in first_row[0]:
                    first_row = next(reader, None)
                if first_row and len(first_row) >= 1:
                    ts_str = first_row[0].strip()
                    return datetime.strptime(ts_str.split(".")[0], "%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
        return None

    def sync_once(self):
        """
        Polls each journal CSV, reads newly appended lines, appends to the
        respective sheet with hourly separators and auto-filter, updates the
        2-row summary, and atomically saves the Excel file.
        Returns the number of new rows added in this cycle.
        """
        total_new_rows = 0

        for sheet_name, csv_path in self.journal_map.items():
            if not os.path.exists(csv_path):
                continue

            ws = self.wb[sheet_name]
            current_offset = self.offsets[sheet_name]

            if sheet_name not in self.sheet_t0 or self.sheet_t0[sheet_name] is None:
                self.sheet_t0[sheet_name] = self._get_sheet_t0(csv_path)

            if sheet_name not in self.sheet_current_hour:
                self.sheet_current_hour[sheet_name] = self._detect_current_hour(ws)

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
                        if not row or len(row) < 3:
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

                        # Parse packet loss: 0 on success, positive integer on loss/failure
                        loss_clean = str(loss_val).strip().upper()
                        if loss_clean in ["NA", "", "0"]:
                            loss_out = 0
                        else:
                            try:
                                loss_out = int(loss_clean)
                            except ValueError:
                                loss_out = loss_clean

                        # Parse timestamp to calculate hour index
                        try:
                            t = datetime.strptime(ts_val.split(".")[0], "%Y-%m-%d %H:%M:%S")
                        except Exception:
                            t = None

                        if self.sheet_t0.get(sheet_name) is None and t is not None:
                            self.sheet_t0[sheet_name] = t

                        t0 = self.sheet_t0.get(sheet_name)
                        if t0 and t:
                            diff_sec = max(0.0, (t - t0).total_seconds())
                            hour_idx = int(diff_sec // 3600)
                        else:
                            hour_idx = 0

                        hour_label = f"Hour {hour_idx + 1}"

                        # Insert hourly separator banner when entering a new hour
                        last_hour = self.sheet_current_hour.get(sheet_name, -1)
                        if hour_idx > last_hour:
                            if t0:
                                start_hr = t0 + timedelta(hours=hour_idx)
                                end_hr = t0 + timedelta(hours=hour_idx + 1)
                                time_range = f"{start_hr.strftime('%H:%M:%S')} - {end_hr.strftime('%H:%M:%S')}"
                            else:
                                time_range = ts_val

                            sep_title = f"=== HOUR {hour_idx + 1} ({time_range}) ==="
                            sep_row = [
                                sep_title,
                                hour_label,
                                "---",
                                "---",
                                "---",
                                f"=== HOUR {hour_idx + 1} START ==="
                            ]
                            ws.append(sep_row)
                            sep_row_idx = ws.max_row
                            ws.row_dimensions[sep_row_idx].height = 22

                            for c_idx in range(1, len(HEADER_COLUMNS) + 1):
                                cell = ws.cell(row=sep_row_idx, column=c_idx)
                                cell.fill = self.sep_fill
                                cell.font = self.sep_font
                                cell.alignment = self.align_center
                                cell.border = self.sep_border

                            self.sheet_current_hour[sheet_name] = hour_idx

                        # Append data row: [Time stamp, Hour, Mesh Node Type, PacketLoss, TTL, String]
                        ws.append([ts_val, hour_label, node_val, loss_out, ttl_out, str_val])
                        curr_row = ws.max_row

                        ws.cell(row=curr_row, column=2).alignment = self.align_center
                        ws.cell(row=curr_row, column=3).alignment = self.align_center
                        cell_loss = ws.cell(row=curr_row, column=4)
                        cell_loss.alignment = self.align_center
                        if loss_out != 0 and loss_clean != "0":
                            cell_loss.font = self.font_loss
                        ws.cell(row=curr_row, column=5).alignment = self.align_center

                        total_new_rows += 1

                        # Update in-memory stats
                        self.stats[sheet_name]["total"] += 1
                        if loss_out != 0 and loss_clean != "0":
                            self.stats[sheet_name]["losses"] += 1
                            self.stats[sheet_name]["loss_hours"].add(hour_idx)

                    self.offsets[sheet_name] = f.tell()

                # Enable auto-filter covering all current data rows
                if ws.max_row >= 4:
                    ws.auto_filter.ref = f"A4:F{ws.max_row}"

                # Update the 2-row summary at the top of the sheet
                total_cnt = self.stats[sheet_name]["total"]
                loss_cnt = self.stats[sheet_name]["losses"]
                loss_pct = (loss_cnt / total_cnt * 100.0) if total_cnt > 0 else 0.0

                ws["B2"].value = f"{loss_pct:.2f}%"
                ws["D2"].value = loss_cnt

                loss_hours = sorted(self.stats[sheet_name]["loss_hours"])
                if not loss_hours:
                    ws["F2"].value = "None"
                    ws["F2"].font = Font(name="Calibri", size=11, bold=True, color="27AE60")
                else:
                    t0 = self.sheet_t0.get(sheet_name)
                    if t0:
                        desc_items = []
                        for h_idx in loss_hours:
                            s_t = (t0 + timedelta(hours=h_idx)).strftime('%H:%M')
                            e_t = (t0 + timedelta(hours=h_idx + 1)).strftime('%H:%M')
                            desc_items.append(f"Hour {h_idx + 1} ({s_t} - {e_t})")
                        detailed_str = ", ".join(desc_items)
                        if len(detailed_str) <= 60:
                            loss_hours_str = detailed_str
                        else:
                            loss_hours_str = ", ".join(f"Hour {h_idx + 1}" for h_idx in loss_hours)
                    else:
                        loss_hours_str = ", ".join(f"Hour {h_idx + 1}" for h_idx in loss_hours)

                    ws["F2"].value = loss_hours_str
                    ws["F2"].font = Font(name="Calibri", size=11, bold=True, color="C0392B")

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
        """
        Performs final sync, appends the Hourly Packet Loss Breakdown table,
        and generates a native Bar Chart showing average packet loss for each 1 hour
        anchored at H1 (top of sheet, beside the executive packet loss stats).
        """
        self.sync_once()

        for sheet_name, csv_path in self.journal_map.items():
            if not os.path.exists(csv_path):
                continue

            ws = self.wb[sheet_name]

            # Lock auto-filter strictly to data rows before appending the summary table!
            data_end_row = ws.max_row
            if data_end_row >= 4:
                ws.auto_filter.ref = f"A4:F{data_end_row}"

            hourly_buckets = self._compute_hourly_buckets(csv_path)

            if not hourly_buckets:
                continue

            # Add spacing and Hourly Summary section at bottom of sheet
            ws.append([])
            ws.append(["--- HOURLY PACKET LOSS BREAKDOWN (1 HOUR INTERVALS) ---", "", "", ""])
            title_row = ws.max_row
            ws.cell(row=title_row, column=1).font = Font(name="Calibri", size=11, bold=True, color="1F4E79")

            # Table Header
            table_header = ["Hour Window", "Total Pings", "Loss Count", "Avg Packet Loss (%)"]
            ws.append(table_header)
            header_row = ws.max_row

            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            for col_idx in range(1, 5):
                cell = ws.cell(row=header_row, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center" if col_idx in [2, 3, 4] else "left")

            # Append hourly data rows
            data_start_row = header_row + 1
            for label, total, losses, loss_pct in hourly_buckets:
                ws.append([label, total, losses, loss_pct])
                curr_row = ws.max_row
                ws.cell(row=curr_row, column=1).alignment = Alignment(horizontal="left")
                ws.cell(row=curr_row, column=2).alignment = Alignment(horizontal="center")
                ws.cell(row=curr_row, column=3).alignment = Alignment(horizontal="center")
                cell_pct = ws.cell(row=curr_row, column=4)
                cell_pct.alignment = Alignment(horizontal="center")
                if losses > 0:
                    cell_pct.font = Font(name="Calibri", size=11, bold=True, color="C0392B")
                else:
                    cell_pct.font = Font(name="Calibri", size=11, bold=True, color="27AE60")

            summary_end_row = ws.max_row

            # Build and insert Bar Chart
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = f"Average Packet Loss per 1 Hour - {sheet_name}"
            chart.y_axis.title = "Packet Loss (%)"
            chart.x_axis.title = "1-Hour Time Window"
            chart.legend = None
            chart.width = 18
            chart.height = 11

            # Exact Y-Axis Range: Min 0, Max 100, Major Steps 10, Minor Steps 1
            chart.y_axis.scaling.min = 0
            chart.y_axis.scaling.max = 100
            chart.y_axis.majorUnit = 10
            chart.y_axis.minorUnit = 1
            chart.y_axis.tickLblPos = "nextTo"
            chart.y_axis.delete = False
            chart.y_axis.number_format = '0'

            # Value labels on bars
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True

            data_ref = Reference(ws, min_col=4, min_row=header_row, max_row=summary_end_row)
            cats_ref = Reference(ws, min_col=1, min_row=data_start_row, max_row=summary_end_row)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)

            # Clear previous charts to prevent duplicates on reruns
            ws._charts = []

            # Place Bar Chart at H1 at the top of the sheet, directly beside the executive packet loss stats
            chart_cell = "H1"
            ws.add_chart(chart, chart_cell)

        self._atomic_save()
        return self.stats

    def _render_barchart_image(self, hourly_data, sheet_name):
        """
        Renders a crisp, high-resolution PNG bar chart of hourly packet loss with
        a fixed 0% - 100% Y-axis (10% steps) and embeds it into the spreadsheet.
        """
        w, h = 900, 480
        img = PILImage.new("RGB", (w, h), color="#FFFFFF")
        draw = ImageDraw.Draw(img)

        # Header bar
        draw.rectangle([(0, 0), (w, 55)], fill="#1F4E79")
        draw.text((25, 18), f"Average Packet Loss per 1 Hour - {sheet_name}", fill="#FFFFFF")

        ml = 75
        mr = 35
        mt = 85
        mb = 80
        pw = w - ml - mr
        ph = h - mt - mb

        # Y-Axis 0% to 100% with 10% steps
        for pct in range(0, 101, 10):
            y = mt + ph - int(pct / 100.0 * ph)
            draw.line([(ml, y), (w - mr, y)], fill="#EEEEEE" if pct > 0 else "#333333", width=1)
            draw.text((ml - 48, y - 7), f"{pct:>3}%", fill="#555555")

        draw.line([(ml, mt), (ml, mt + ph)], fill="#333333", width=2)
        draw.line([(ml, mt + ph), (w - mr, mt + ph)], fill="#333333", width=2)

        draw.text((ml, mt - 22), "Packet Loss (%)", fill="#1F4E79")
        draw.text((w // 2 - 40, mt + ph + 45), "1-Hour Time Window", fill="#1F4E79")

        n = len(hourly_data)
        if n > 0:
            bar_w = min(75, max(24, int(pw / (n * 1.6))))
            step = pw / n
            for idx, (label, total, loss, pct) in enumerate(hourly_data):
                xc = ml + idx * step + step / 2
                x1 = xc - bar_w / 2
                x2 = xc + bar_w / 2
                bh = int(min(pct, 100.0) / 100.0 * ph)
                y1 = mt + ph - bh
                y2 = mt + ph

                bar_color = "#C0392B" if pct > 0 else "#27AE60"
                draw.rectangle([(x1, y1), (x2, y2)], fill=bar_color, outline="#2C3E50", width=1)

                val_str = f"{pct:.1f}%"
                draw.text((x1, max(mt, y1 - 16)), val_str, fill=bar_color)
                draw.text((x1, mt + ph + 8), label, fill="#222222")

        temp_img = os.path.join(tempfile.gettempdir(), f"chart_{sheet_name}_{os.getpid()}.png")
        img.save(temp_img)
        return temp_img

    def _compute_hourly_buckets(self, csv_path):
        """
        Parses all rows in a journal CSV and groups them into 1-hour interval buckets
        starting from the first ping's timestamp.
        """
        from datetime import datetime, timedelta

        rows = []
        try:
            with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
                reader = csv.reader(f)
                header = next(reader, None)
                for r in reader:
                    if not r or len(r) < 3:
                        continue
                    ts_str = r[0]
                    if len(r) >= 6 and str(r[1]).strip().lower().startswith("hour"):
                        loss_str = str(r[3]).strip().upper()
                    else:
                        loss_str = str(r[2]).strip().upper()
                    is_loss = (loss_str not in ["NA", "", "0"])
                    rows.append((ts_str, is_loss))
        except Exception:
            return []

        if not rows:
            return []

        t0 = None
        for ts_str, _ in rows:
            try:
                t0 = datetime.strptime(ts_str.split(".")[0], "%Y-%m-%d %H:%M:%S")
                break
            except Exception:
                continue

        if not t0:
            return []

        buckets = []
        for ts_str, is_loss in rows:
            try:
                t = datetime.strptime(ts_str.split(".")[0], "%Y-%m-%d %H:%M:%S")
            except Exception:
                continue

            diff_sec = (t - t0).total_seconds()
            if diff_sec < 0:
                diff_sec = 0
            hour_idx = int(diff_sec // 3600)

            while len(buckets) <= hour_idx:
                idx = len(buckets)
                start_hr = t0 + timedelta(hours=idx)
                end_hr = t0 + timedelta(hours=idx + 1)
                b_name = f"Hour {idx + 1} ({start_hr.strftime('%H:%M')} - {end_hr.strftime('%H:%M')})"
                buckets.append({"label": b_name, "total": 0, "loss": 0})

            buckets[hour_idx]["total"] += 1
            if is_loss:
                buckets[hour_idx]["loss"] += 1

        results = []
        for b in buckets:
            pct = (b["loss"] / b["total"] * 100.0) if b["total"] > 0 else 0.0
            results.append((b["label"], b["total"], b["loss"], round(pct, 2)))

        return results


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

    # If rebuilding from journals, remove stale existing file to prevent row duplication
    if os.path.exists(excel_path):
        try:
            os.remove(excel_path)
        except Exception:
            pass

    print(f"[*] Synchronizing {len(journal_map)} sheets into: {excel_path}...")
    manager = ExcelSyncManager(excel_path, journal_map)
    stats = manager.finalize()

    print("[+] Synchronization complete!\n")
    print(f"{'Sheet Name':<25} {'Total Pings':<14} {'Packet Losses':<14} {'Loss Hour(s)':<25}")
    print("-" * 80)
    for s_name, data in stats.items():
        loss_hrs = sorted(data.get("loss_hours", []))
        hrs_str = ", ".join(f"Hour {h+1}" for h in loss_hrs) if loss_hrs else "None"
        print(f"{s_name:<25} {data['total']:<14} {data['losses']:<14} {hrs_str:<25}")
    print("-" * 80)


def main():
    parser = argparse.ArgumentParser(description="Synchronize Orbi Mesh CSV journals to Master Excel file.")
    parser.add_argument("--journal-dir", required=True, help="Directory containing sheet CSV files")
    parser.add_argument("--excel-file", required=True, help="Target Excel (.xlsx) file")
    args = parser.parse_args()

    run_standalone_sync(args.journal_dir, args.excel_file)


if __name__ == "__main__":
    main()
