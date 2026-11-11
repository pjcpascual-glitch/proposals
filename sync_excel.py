"""
Offer & Tender Tracker — Python Sync & Database Helper Script
=============================================================
This script provides headless sync, validation, JSON database export/import,
and an optional local HTTP server for the Offer Tracker Dashboard.

Usage:
    py sync_excel.py --stats          # Display summary statistics & overdue tenders
    py sync_excel.py --export-json    # Export Excel data to offer_tracker_backup.json
    py sync_excel.py --import-json    # Merge JSON backup changes into Excel
    py sync_excel.py --sync           # Bidirectional sync between Excel and JSON
    py sync_excel.py --serve          # Run local server with CORS & save API on port 8000
"""

import sys
import os
import json
import datetime
import http.server
import socketserver
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is required. Run 'pip install openpyxl'")
    sys.exit(1)

EXCEL_FILE = "REG_16_Tender List_v1.xlsx"
DATA_SHEET = "REG_16_Tender List_v1"
BACKUP_FILE = "offer_tracker_backup.json"
PORT = 8000

CURRENCY_RATES = {
    'EUR': 1.0,
    'USD': 1.15,
    'PHP': 70.0,
    'STN': 24.5,
    'CFA': 655.957,
    'FCFA': 655.957,
    'BOB': 7.5
}

def clean_val(v):
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return v

def get_last_data_row(ws):
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(row=r, column=col).value is not None for col in range(1, 15)):
            return r
    return 2

def load_tenders_from_excel(filepath=EXCEL_FILE):
    if not os.path.exists(filepath):
        print(f"Error: Excel file {filepath} not found.")
        return []

    wb = openpyxl.load_workbook(filepath, data_only=True)
    if DATA_SHEET not in wb.sheetnames:
        print(f"Error: Sheet '{DATA_SHEET}' not found in {filepath}.")
        return []

    ws = wb[DATA_SHEET]
    headers = [str(ws.cell(2, c).value or '').strip() for c in range(1, ws.max_column + 1)]
    
    last_r = get_last_data_row(ws)
    tenders = []
    
    for r in range(3, last_r + 1):
        tender_id = ws.cell(r, 1).value
        client = ws.cell(r, 3).value
        desc = ws.cell(r, 6).value
        
        # Skip completely empty rows
        if not tender_id and not client and not desc:
            continue
            
        row_dict = {
            '_excel_row': r,
            'tender_id': tender_id,
            'firm': clean_val(ws.cell(r, 2).value),
            'client': clean_val(ws.cell(r, 3).value),
            'new_client': clean_val(ws.cell(r, 4).value),
            'country': clean_val(ws.cell(r, 5).value),
            'description': clean_val(ws.cell(r, 6).value),
            'type': clean_val(ws.cell(r, 7).value),
            'responsible': clean_val(ws.cell(r, 8).value),
            'value': clean_val(ws.cell(r, 9).value),
            'currency': clean_val(ws.cell(r, 10).value),
            'eur': clean_val(ws.cell(r, 11).value),
            'date': clean_val(ws.cell(r, 12).value),
            'exp_date': clean_val(ws.cell(r, 13).value),
            'status': clean_val(ws.cell(r, 14).value),
            'comments': clean_val(ws.cell(r, 15).value),
            'implementation': clean_val(ws.cell(r, 16).value)
        }
        tenders.append(row_dict)

    return tenders

def find_dashboard_html():
    for name in ["ADM_APP_ATI_000 Proposals.html", "Offer Dashboard.html"]:
        if os.path.exists(name):
            return name
    for f in os.listdir("."):
        if f.endswith(".html") and not f.startswith("."):
            return f
    return "ADM_APP_ATI_000 Proposals.html"

def update_html_snapshot(html_path=None, excel_path=EXCEL_FILE):
    """Bakes the latest Excel data directly into the HTML file for zero-click sharing with executives."""
    import re
    if not html_path:
        html_path = find_dashboard_html()
    if not os.path.exists(html_path):
        return
    tenders = load_tenders_from_excel(excel_path)
    if not tenders:
        return
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    json_str = json.dumps(tenders, ensure_ascii=False)
    pattern = r'(<script id="prebaked-data" type="application/json">\s*)(.*?)(\s*</script>)'
    if re.search(pattern, html, re.DOTALL):
        html = re.sub(pattern, lambda m: f'{m.group(1)}{json_str}{m.group(3)}', html, flags=re.DOTALL)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"Baked {len(tenders)} proposals into {html_path} (Ready for zero-click sharing).")
    else:
        print(f"Warning: prebaked-data block not found in {html_path}")

def export_to_json(filepath=EXCEL_FILE, backup_file=BACKUP_FILE):
    tenders = load_tenders_from_excel(filepath)
    payload = {
        'version': '2.0',
        'exported_at': datetime.datetime.now().isoformat(),
        'source_file': filepath,
        'count': len(tenders),
        'tenders': tenders
    }
    with open(backup_file, 'w', encoding='utf-8') as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    print(f"Exported {len(tenders)} proposals to {backup_file}")
    update_html_snapshot()
    return payload

def compute_eur(value, currency):
    if not value or not isinstance(value, (int, float)):
        return None
    cur = str(currency or 'EUR').strip().upper()
    rate = CURRENCY_RATES.get(cur, 1.0)
    if rate > 0:
        return round(value / rate, 2)
    return value

def append_tender_to_excel(tender, filepath=EXCEL_FILE):
    wb = openpyxl.load_workbook(filepath)
    ws = wb[DATA_SHEET]
    last_r = get_last_data_row(ws)
    new_r = last_r + 1
    
    # 1: Tender Id.
    ws.cell(new_r, 1, tender.get('tender_id'))
    # 2: Firm
    ws.cell(new_r, 2, tender.get('firm') or 'ATI Consultores')
    # 3: Client
    ws.cell(new_r, 3, tender.get('client'))
    # 4: New Client
    ws.cell(new_r, 4, tender.get('new_client') or 'No')
    # 5: Country
    ws.cell(new_r, 5, tender.get('country'))
    # 6: Description
    ws.cell(new_r, 6, tender.get('description'))
    # 7: Type
    ws.cell(new_r, 7, tender.get('type'))
    # 8: Responsible
    ws.cell(new_r, 8, tender.get('responsible'))
    
    # 9: Value
    val = tender.get('value')
    if val is not None and str(val).strip() != '':
        try:
            ws.cell(new_r, 9, float(val))
        except ValueError:
            ws.cell(new_r, 9, val)
    
    # 10: Currency
    cur = (tender.get('currency') or 'EUR').strip().upper()
    ws.cell(new_r, 10, cur)
    
    # 11: EUR Formula or Value
    if cur == 'EUR':
        ws.cell(new_r, 11, f'=I{new_r}')
    elif cur == 'USD':
        ws.cell(new_r, 11, f'=I{new_r}/1.15')
    elif cur == 'PHP':
        ws.cell(new_r, 11, f'=I{new_r}/70')
    else:
        eur_val = tender.get('eur')
        if eur_val:
            ws.cell(new_r, 11, float(eur_val))
            
    # 12: Date
    d_str = tender.get('date')
    if d_str:
        try:
            ws.cell(new_r, 12, datetime.datetime.strptime(d_str, "%Y-%m-%d"))
        except ValueError:
            ws.cell(new_r, 12, d_str)
            
    # 13: Exp. Date
    exp_str = tender.get('exp_date')
    if exp_str:
        try:
            ws.cell(new_r, 13, datetime.datetime.strptime(exp_str, "%Y-%m-%d"))
        except ValueError:
            ws.cell(new_r, 13, exp_str)
            
    # 14: Status
    ws.cell(new_r, 14, tender.get('status') or 'Preliminary')
    # 15: Comments
    ws.cell(new_r, 15, tender.get('comments'))
    # 16: Implementation
    ws.cell(new_r, 16, tender.get('implementation'))
    
    wb.save(filepath)
    print(f"Successfully added tender ID {tender.get('tender_id')} at row {new_r} in {filepath}")
    return new_r

def show_stats():
    tenders = load_tenders_from_excel()
    print(f"\n==========================================")
    print(f"   ATI OFFER TRACKER — DASHBOARD SUMMARY  ")
    print(f"==========================================")
    print(f"Total Proposals in Excel: {len(tenders)}")
    
    status_counts = {}
    users_stats = {}
    overdue_count = 0
    today = datetime.date.today()
    
    for t in tenders:
        stat = (t.get('status') or 'Unknown').strip().title()
        status_counts[stat] = status_counts.get(stat, 0) + 1
        
        resp = (t.get('responsible') or 'Unassigned').strip()
        if resp not in users_stats:
            users_stats[resp] = {'total': 0, 'won': 0, 'lost': 0, 'active': 0}
            
        users_stats[resp]['total'] += 1
        if stat == 'Accepted':
            users_stats[resp]['won'] += 1
        elif stat in ['Rejected', 'Cancelled']:
            users_stats[resp]['lost'] += 1
        elif stat in ['Issued', 'Preliminary']:
            users_stats[resp]['active'] += 1
            
        exp_d = t.get('exp_date')
        if exp_d and stat in ['Issued', 'Preliminary']:
            try:
                exp_date_obj = datetime.datetime.strptime(exp_d[:10], "%Y-%m-%d").date()
                if exp_date_obj < today:
                    overdue_count += 1
            except Exception:
                pass

    print("\n[Status Breakdown]")
    for st, count in sorted(status_counts.items(), key=lambda x: -x[1]):
        print(f"  • {st:15s}: {count:4d}")

    print(f"\n[Overdue for Follow-up (Expired active proposals)]: {overdue_count}")

    print("\n[User Performance & Success Rates (Option A: Won / Total Sent)]")
    print(f"  {'User':<8} {'Total':<7} {'Won':<6} {'Active':<8} {'Win Rate (%)':<16}")
    print("  " + "-" * 45)
    for u, st in sorted(users_stats.items(), key=lambda x: -(x[1]['won'] / x[1]['total'] if x[1]['total'] > 0 else 0)):
        rate = (st['won'] / st['total'] * 100) if st['total'] > 0 else 0.0
        print(f"  {u:<8} {st['total']:<7} {st['won']:<6} {st['active']:<8} {rate:6.1f}%")
    print(f"==========================================\n")

class EnhancedRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/save_tender':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            try:
                data = json.loads(body.decode('utf-8'))
                new_row = append_tender_to_excel(data)
                export_to_json() # also keep backup json synchronized
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'ok', 'row': new_row}).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

def run_server():
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), EnhancedRequestHandler) as httpd:
        print(f"\n========================================================")
        print(f"  ATI Offer Tracker Development & Sync Server")
        print(f"  Dashboard: http://localhost:{PORT}/Offer%20Dashboard.html")
        print(f"  Data Sync: Auto-linking with {EXCEL_FILE}")
        print(f"========================================================")
        print("Press CTRL+C to stop.\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped safely.")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        if cmd == '--stats':
            show_stats()
        elif cmd == '--export-json':
            export_to_json()
        elif cmd == '--bake':
            update_html_snapshot()
        elif cmd == '--sync':
            export_to_json()
            show_stats()
        elif cmd == '--serve':
            run_server()
        else:
            print(f"Unknown command: {cmd}")
            print("Options: --stats, --export-json, --bake, --sync, --serve")
    else:
        show_stats()

