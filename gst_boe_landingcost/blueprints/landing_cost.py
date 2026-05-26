import base64
import json
import os
import re
import urllib.request
import urllib.error
from datetime import date
from flask import Blueprint, request, jsonify, render_template

landing_bp = Blueprint('landing', __name__, url_prefix='/landing')

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/"
    "models/gemini-2.5-flash:generateContent?key={key}"
)

def fmt_inr(n: float) -> str: return "₹" + f"{int(round(n)):,}"
def fmt_usd(n: float) -> str: return f"${n:.2f}"
def fmt_cur(n: float, currency: str) -> str: return f"¥{n:.2f}" if currency == "RMB" else f"${n:.2f}"

def calc_landed(items, rate, bank, ship, duty, trans):
    total_addl = bank + ship + duty + trans
    inv_usd    = sum(it["qty"] * it["unitPrice"] for it in items)
    inv_inr    = inv_usd * rate
    grand      = inv_inr + total_addl

    result = []
    for it in items:
        item_inr   = it["qty"] * it["unitPrice"] * rate
        share      = (item_inr / inv_inr) if inv_inr > 0 else 0
        addl_share = total_addl * share
        total_item = item_inr + addl_share
        per_unit   = (total_item / it["qty"]) if it["qty"] > 0 else 0
        result.append({
            **it,
            "item_inr":   round(item_inr, 2),
            "share":      round(share * 100, 2),
            "addl_share": round(addl_share, 2),
            "total_item": round(total_item, 2),
            "per_unit":   round(per_unit, 2),
            "bank_s":     round(bank * share, 2),
            "ship_s":     round(ship * share, 2),
            "duty_s":     round(duty * share, 2),
            "trans_s":    round(trans * share, 2),
        })

    return {
        "items":       result,
        "inv_usd":     round(inv_usd, 2),
        "inv_inr":     round(inv_inr, 2),
        "total_addl":  round(total_addl, 2),
        "grand":       round(grand, 2),
        "bank":        round(bank, 2),
        "ship":        round(ship, 2),
        "duty":        round(duty, 2),
        "trans":       round(trans, 2),
        "rate":        rate,
    }

@landing_bp.route('/')
def index():
    return render_template('landing_cost.html')

@landing_bp.route('/api/calculate', methods=['POST'])
def api_calculate():
    data     = request.get_json(force=True)
    items    = data.get('items', [])
    currency = data.get('currency', 'USD')
    rate     = float(data.get('rmbRate', 11.5)) if currency == 'RMB' else float(data.get('rate', 84))
    bank  = float(data.get('bank', 0))
    ship  = float(data.get('ship', 0))
    duty  = float(data.get('duty', 0))
    trans = float(data.get('trans', 0))
    return jsonify(calc_landed(items, rate, bank, ship, duty, trans))

@landing_bp.route('/api/parse-invoice', methods=['POST'])
def api_parse_invoice():
    api_key = request.form.get('api_key', '').strip()
    if not api_key: return jsonify({'error': 'Gemini API key is required'}), 400
    f = request.files.get('file')
    if not f: return jsonify({'error': 'No file uploaded'}), 400

    ext_map = {'.pdf': 'application/pdf', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png'}
    ext   = os.path.splitext(f.filename.lower())[1]
    mime  = ext_map.get(ext, 'application/octet-stream')
    b64   = base64.b64encode(f.read()).decode()

    prompt = """Extract all line items from this commercial invoice.
Return ONLY valid JSON, no markdown, no explanation:
{"invoiceRef": "invoice number or empty string", "invoiceDate": "date in YYYY-MM-DD format or empty string", "currency": "USD or RMB", "items": [{"name": "item description", "qty": number, "unitPrice": number}]}
Rules:
- unitPrice = per unit price in the invoice currency
- If only total price given, divide by qty to get unitPrice
- qty = numeric value only
- currency: return "RMB" if invoice is in CNY/RMB/Yuan, otherwise "USD"
- DO NOT invent data"""

    payload = json.dumps({"contents": [{"parts": [{"inline_data": {"mime_type": mime, "data": b64}}, {"text": prompt}]}]}).encode()
    url = GEMINI_URL.format(key=api_key)
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp: result = json.loads(resp.read())
        raw = result["candidates"][0]["content"]["parts"][0]["text"]
        raw = raw.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(raw)
        return jsonify(parsed)
    except Exception as exc:
        return jsonify({'error': str(exc)}), 500

@landing_bp.route('/api/report', methods=['POST'])
def api_report():
    data = request.get_json(force=True)
    items = data.get('items', [])
    currency = data.get('currency', 'USD')
    inv_ref = data.get('invoiceRef', '')
    inv_date = data.get('invoiceDate', '')
    rate = float(data.get('rmbRate', 11.5)) if currency == 'RMB' else float(data.get('rate', 84))
    bank = float(data.get('bank', 0))
    ship = float(data.get('ship', 0))
    duty = float(data.get('duty', 0))
    trans = float(data.get('trans', 0))

    calc = calc_landed(items, rate, bank, ship, duty, trans)
    c = calc

    cur_label = "RMB" if currency == "RMB" else "USD"
    rate_label = f"1 {cur_label} = ₹{rate:.2f}"
    report_date = date.today().strftime("%d-%m-%Y")
    
    bank_ship = bank + ship
    duty_trans = duty + trans
    
    pct_inv = (c['inv_inr'] / c['grand'] * 100) if c['grand'] > 0 else 0
    pct_bs = (bank_ship / c['grand'] * 100) if c['grand'] > 0 else 0
    pct_dt = (duty_trans / c['grand'] * 100) if c['grand'] > 0 else 0
    
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><title>Landed Cost Report</title>
<style>
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 11px; color: #111; background: #fff; padding: 20px; margin: 0; }}
  @page {{ size: A4 landscape; margin: 10mm; }}
  .header-box {{ text-align: center; margin-bottom: 20px; border-bottom: 2px solid #222; padding-bottom: 10px; }}
  .header-box h1 {{ margin: 0; font-size: 18px; text-transform: uppercase; letter-spacing: 1px; }}
  .header-box h2 {{ margin: 4px 0; font-size: 14px; color: #444; }}
  .header-box p {{ margin: 2px 0; color: #666; font-size: 10px; }}
  
  .meta-grid {{ display: flex; justify-content: space-between; margin-bottom: 15px; font-weight: bold; border: 1px solid #ddd; padding: 8px; background: #f9f9f9; }}
  
  .summary-grid {{ display: flex; gap: 10px; margin-bottom: 15px; text-align: center; }}
  .summary-box {{ flex: 1; border: 1px solid #aaa; padding: 8px; }}
  .summary-label {{ font-size: 9px; text-transform: uppercase; color: #555; margin-bottom: 4px; display: block; }}
  .summary-val {{ font-size: 13px; font-weight: bold; }}
  
  table {{ width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 20px; }}
  th, td {{ padding: 6px 4px; border: 1px solid #888; text-align: right; }}
  th {{ background: #eee; font-weight: bold; text-align: center; }}
  td:nth-child(2) {{ text-align: left; }}
  .grand-row {{ background: #ddd; font-weight: bold; }}
  
  .footer-grid {{ display: flex; gap: 10px; text-align: center; margin-bottom: 10px; }}
  .footer-box {{ flex: 1; border: 2px solid #555; padding: 8px; }}
  .footer-title {{ font-size: 10px; font-weight: bold; margin-bottom: 4px; text-transform: uppercase; }}
  .footer-amt {{ font-size: 14px; font-weight: bold; color: #000; }}
  .footer-pct {{ font-size: 12px; color: #444; margin-top: 2px; }}
  
  .btn-print {{ display: block; width: 200px; margin: 0 auto 20px; padding: 10px; background: #000; color: #fff; text-align: center; cursor: pointer; text-decoration: none; font-weight: bold; border-radius: 4px; }}
  @media print {{ .no-print {{ display: none !important; }} body {{ padding: 0; }} }}
</style>
</head>
<body>
<div class="no-print" style="text-align:center;">
  <button class="btn-print" onclick="window.print()">🖨 Print / Save as PDF</button>
</div>

<div class="header-box">
  <h1>IMPORT COST CALCULATOR</h1>
  <h2>China → India Landed Cost Report</h2>
  <p>Item-Wise Breakdown with Proportional Cost Allocation</p>
</div>

<div class="meta-grid">
  <div>Report Date: {report_date} {f"| Invoice Ref: {inv_ref}" if inv_ref else ""}</div>
  <div>Currency: {cur_label} | {rate_label}</div>
</div>

<div class="summary-grid">
  <div class="summary-box"><span class="summary-label">Invoice Value ({cur_label})</span><span class="summary-val">{fmt_cur(c['inv_usd'], currency)}</span></div>
  <div class="summary-box"><span class="summary-label">Invoice Value (INR)</span><span class="summary-val">{fmt_inr(c['inv_inr'])}</span></div>
  <div class="summary-box"><span class="summary-label">Bank Charges</span><span class="summary-val">{fmt_inr(bank)}</span></div>
  <div class="summary-box"><span class="summary-label">Shipping Cost</span><span class="summary-val">{fmt_inr(ship)}</span></div>
  <div class="summary-box"><span class="summary-label">Custom Duty</span><span class="summary-val">{fmt_inr(duty)}</span></div>
  <div class="summary-box"><span class="summary-label">Local Transport</span><span class="summary-val">{fmt_inr(trans)}</span></div>
</div>

<table>
  <tr>
    <th>#</th>
    <th>Item / Description</th>
    <th>Qty</th>
    <th>Unit Price ({cur_label})</th>
    <th>Invoice Value (INR)</th>
    <th>Share</th>
    <th>Bank Charges</th>
    <th>Shipping</th>
    <th>Custom Duty</th>
    <th>Local Trans.</th>
    <th>Total Addl.</th>
    <th>TOTAL LANDED (INR)</th>
    <th>PER UNIT (INR)</th>
  </tr>
"""
    for i, it in enumerate(c['items']):
        html += f"""
  <tr>
    <td style="text-align:center;">{i+1}</td>
    <td>{it.get('name', f'Item {i+1}')}</td>
    <td style="text-align:center;">{it['qty']}</td>
    <td>{fmt_cur(it['unitPrice'], currency)}</td>
    <td>{fmt_inr(it['item_inr'])}</td>
    <td>{it['share']:.1f}%</td>
    <td>{fmt_inr(it['bank_s'])}</td>
    <td>{fmt_inr(it['ship_s'])}</td>
    <td>{fmt_inr(it['duty_s'])}</td>
    <td>{fmt_inr(it['trans_s'])}</td>
    <td>{fmt_inr(it['addl_share'])}</td>
    <td style="font-weight:bold;">{fmt_inr(it['total_item'])}</td>
    <td style="font-weight:bold;">{fmt_inr(it['per_unit'])}</td>
  </tr>
"""
        
    html += f"""
  <tr class="grand-row">
    <td colspan="4" style="text-align:right;">GRAND TOTAL</td>
    <td>{fmt_inr(c['inv_inr'])}</td>
    <td>100%</td>
    <td>{fmt_inr(bank)}</td>
    <td>{fmt_inr(ship)}</td>
    <td>{fmt_inr(duty)}</td>
    <td>{fmt_inr(trans)}</td>
    <td>{fmt_inr(c['total_addl'])}</td>
    <td>{fmt_inr(c['grand'])}</td>
    <td></td>
  </tr>
</table>

<h3 style="margin-bottom:8px; font-size:12px; text-transform:uppercase; border-bottom:1px solid #ccc; padding-bottom:4px;">TOTAL LANDED COST</h3>
<div class="footer-grid">
  <div class="footer-box" style="background:#f9f9f9;">
    <div class="footer-title">Invoice Value</div>
    <div class="footer-amt">{fmt_inr(c['inv_inr'])}</div>
    <div class="footer-pct">{pct_inv:.1f}%</div>
  </div>
  <div class="footer-box" style="background:#fff3e0;">
    <div class="footer-title">Bank & Shipping</div>
    <div class="footer-amt">{fmt_inr(bank_ship)}</div>
    <div class="footer-pct">{pct_bs:.1f}%</div>
  </div>
  <div class="footer-box" style="background:#e8f5e9;">
    <div class="footer-title">Custom & Local</div>
    <div class="footer-amt">{fmt_inr(duty_trans)}</div>
    <div class="footer-pct">{pct_dt:.1f}%</div>
  </div>
  <div class="footer-box" style="background:#e3f2fd; border-color:#1565c0;">
    <div class="footer-title">Factory Delivered</div>
    <div class="footer-amt">{fmt_inr(c['grand'])}</div>
    <div class="footer-pct">100.0%</div>
  </div>
</div>
<p style="text-align:center; font-size:9px; color:#666; margin-top:20px;">
  Additional costs allocated proportionally by invoice value share.<br>
  Total Landed: {fmt_inr(c['grand'])} | Invoice: {fmt_cur(c['inv_usd'], currency)} @ ₹{rate:.2f}
</p>
</body>
</html>
"""
    return html, 200, {'Content-Type': 'text/html; charset=utf-8'}
