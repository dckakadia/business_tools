import base64
import json
import os
import re
import google.generativeai as genai
from flask import Blueprint, request, jsonify, render_template

boe_bp = Blueprint('boe', __name__, url_prefix='/boe')

@boe_bp.route('/')
def index():
    return render_template('boe_calculator.html')

@boe_bp.route('/scan-invoice', methods=['POST'])
def scan_invoice():
    try:
        api_key = request.form.get('api_key', '').strip()
        if not api_key: return jsonify({'error': 'Gemini API key is required'}), 400

        file = request.files.get('file')
        if not file: return jsonify({'error': 'No file uploaded'}), 400

        filename = file.filename.lower()
        file_bytes = file.read()
        mime_type = file.mimetype or 'application/octet-stream'

        if filename.endswith('.pdf'): mime_type = 'application/pdf'
        elif filename.endswith(('.jpg', '.jpeg')): mime_type = 'image/jpeg'
        elif filename.endswith('.png'): mime_type = 'image/png'

        selected_model = request.form.get('model', 'gemini-2.5-flash').strip()
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(selected_model)

        prompt = """You are an expert invoice data extractor for Indian import/customs calculations.
Carefully analyze this Proforma Invoice (PI) document and extract ALL data.
Return ONLY a valid JSON object with this exact structure:
{
  "items": [
    {
      "name": "exact product/item description",
      "qty": 10,
      "unit_price_usd": 121.64
    }
  ],
  "exchange_rate": 87.20,
  "exchange_rate_date": "2025-07-29",
  "boe_number": "",
  "boe_date": "",
  "port_code": "",
  "boe_taxable_amount": 0,
  "boe_gst_amount": 0
}
Rules:
- Extract EVERY line item from the invoice table
- qty must be a number
- unit_price_usd must be a number in USD only
- exchange_rate: fill only if explicitly stated, else null
- exchange_rate_date: YYYY-MM-DD format if found, else null
- boe_number, boe_date, port_code: fill only if found, else empty string
- boe_taxable_amount, boe_gst_amount: fill only if found, else 0
- DO NOT invent data"""

        image_part = {'inline_data': {'mime_type': mime_type, 'data': base64.b64encode(file_bytes).decode('utf-8')}}
        response = model.generate_content([prompt, image_part])
        raw_text = response.text.strip()

        clean = re.sub(r'^```(?:json)?\s*', '', raw_text, flags=re.MULTILINE)
        clean = re.sub(r'\s*```$', '', clean, flags=re.MULTILINE).strip()

        try:
            data = json.loads(clean)
        except json.JSONDecodeError:
            match = re.search(r'\{[\s\S]+\}', clean)
            if match: data = json.loads(match.group())
            else: return jsonify({'error': 'Gemini returned non-JSON response', 'raw_text': raw_text}), 500

        data['raw_text'] = raw_text
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@boe_bp.route('/scan-boe', methods=['POST'])
def scan_boe():
    try:
        api_key = request.form.get('api_key', '').strip()
        if not api_key: return jsonify({'error': 'Gemini API key is required'}), 400
        file = request.files.get('file')
        if not file: return jsonify({'error': 'No file uploaded'}), 400

        filename = file.filename.lower()
        file_bytes = file.read()
        mime_type = 'application/pdf' if filename.endswith('.pdf') else 'image/jpeg' if filename.endswith(('.jpg','.jpeg')) else 'image/png'

        selected_model = request.form.get('model', 'gemini-2.5-flash').strip()
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(selected_model)

        prompt = """You are an expert at reading Indian Customs Bill of Entry (BOE) documents.
Extract the following fields and return ONLY a valid JSON object:
{
  "boe_number": "BE No from top header",
  "boe_date": "YYYY-MM-DD",
  "port_code": "Port Code",
  "tot_ass_val": 0.00,
  "bcd_amount": 0.00,
  "sws_amount": 0.00,
  "boe_gst_amount": 0.00,
  "exchange_rate": null,
  "exchange_rate_date": null
}
CRITICAL EXTRACTION RULES:
1. boe_number: "BE No" in top header
2. boe_date: "BE Date", YYYY-MM-DD
3. port_code: "Port Code"
4. Section "C. DUTY SUMMARY": tot_ass_val ("18.TOT.ASS VAL"), bcd_amount ("1.BCD"), sws_amount ("3.SWS")
5. boe_gst_amount: "7.IGST" column value
6. exchange_rate: "H. PROCESSING DETAILS" e.g. "1 USD=88.7INR" -> 88.7
7. exchange_rate_date: Date of "Assessment" in H. PROCESSING DETAILS
DO NOT compute boe_taxable_amount yourself."""

        image_part = {'inline_data': {'mime_type': mime_type, 'data': base64.b64encode(file_bytes).decode('utf-8')}}
        response = model.generate_content([prompt, image_part])
        raw_text = response.text.strip()

        clean = re.sub(r'^```(?:json)?\s*', '', raw_text, flags=re.MULTILINE)
        clean = re.sub(r'\s*```$', '', clean, flags=re.MULTILINE).strip()
        try:
            data = json.loads(clean)
        except json.JSONDecodeError:
            match = re.search(r'\{[\s\S]+\}', clean)
            if match: data = json.loads(match.group())
            else: return jsonify({'error': 'Gemini returned non-JSON response', 'raw_text': raw_text}), 500

        tot_ass_val = float(data.get('tot_ass_val') or 0)
        bcd_amount  = float(data.get('bcd_amount')  or 0)
        sws_amount  = float(data.get('sws_amount')  or 0)
        data['boe_taxable_amount'] = round(tot_ass_val + bcd_amount + sws_amount, 2)
        data['_formula'] = f"{tot_ass_val} (ASS VAL) + {bcd_amount} (BCD) + {sws_amount} (SWS) = {data['boe_taxable_amount']}"

        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@boe_bp.route('/scan-challan', methods=['POST'])
def scan_challan():
    try:
        api_key = request.form.get('api_key', '').strip()
        if not api_key: return jsonify({'error': 'Gemini API key is required'}), 400
        file = request.files.get('file')
        if not file: return jsonify({'error': 'No file uploaded'}), 400

        filename = file.filename.lower()
        file_bytes = file.read()
        mime_type = 'application/pdf' if filename.endswith('.pdf') else 'image/jpeg' if filename.endswith(('.jpg','.jpeg')) else 'image/png'

        selected_model = request.form.get('model', 'gemini-2.5-flash').strip()
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(selected_model)

        prompt = """You are an expert at reading Indian ICEGATE Customs Challan payment receipts.
Extract the following fields from this challan document and return ONLY a valid JSON object:
{
  "challan_date": "YYYY-MM-DD",
  "challan_amount": 0.00
}
Rules:
- challan_date: payment date in YYYY-MM-DD
- challan_amount: TOTAL amount paid on this challan in INR
- DO NOT invent data"""

        image_part = {'inline_data': {'mime_type': mime_type, 'data': base64.b64encode(file_bytes).decode('utf-8')}}
        response = model.generate_content([prompt, image_part])
        raw_text = response.text.strip()

        clean = re.sub(r'^```(?:json)?\s*', '', raw_text, flags=re.MULTILINE)
        clean = re.sub(r'\s*```$', '', clean, flags=re.MULTILINE).strip()
        try:
            data = json.loads(clean)
        except json.JSONDecodeError:
            match = re.search(r'\{[\s\S]+\}', clean)
            if match: data = json.loads(match.group())
            else: return jsonify({'error': 'Gemini returned non-JSON response', 'raw_text': raw_text}), 500

        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
