#!/usr/bin/env python3
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import io, os, zipfile, re
from datetime import datetime

app = Flask(__name__)
CORS(app)
BASE = os.path.dirname(os.path.abspath(__file__))

TEMPLATES = {
    'china':  os.path.join(BASE, 'template_china.xlsx'),
    'india':  os.path.join(BASE, 'template_india.xlsx'),
    'nasn':   os.path.join(BASE, 'template_nasn.xlsx'),
    'mexico': os.path.join(BASE, 'template_mexico.xlsx'),
}

@app.route('/')
def index():
    return send_from_directory(BASE, 'invoice_generator.html')

def fmt_ordinal(date_str):
    d = datetime.strptime(date_str, '%Y-%m-%d')
    day = d.day
    suf = {1:'st',2:'nd',3:'rd'}.get(day if day<20 else day%10, 'th')
    return f"{day}{suf}.{d.strftime('%b')}.{d.year}"

def fmt_upper(date_str):
    d = datetime.strptime(date_str, '%Y-%m-%d')
    return f"{d.day}-{d.strftime('%b').upper()}-{d.year}"

def set_number_cell(content, cell_ref, value):
    pattern = rf'(<c r="{cell_ref}"[^>]*>(?:<f>[^<]*</f>)?)<v>[^<]*</v>'
    new = re.sub(pattern, rf'\g<1><v>{value}</v>', content)
    if new == content:
        new = re.sub(rf'(<c r="{cell_ref}"[^>]*>)<v>[^<]*</v>', rf'\g<1><v>{value}</v>', content)
    return new

def set_inline_str_cell(content, cell_ref, value):
    escaped = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    pattern = rf'(<c r="{cell_ref}"[^>]*t="inlineStr"[^>]*><is><t>)[^<]*(</t></is></c>)'
    new = re.sub(pattern, rf'\g<1>{escaped}\g<2>', content)
    if new == content:
        pattern2 = rf'(<c r="{cell_ref}"[^>]*><is><t>)[^<]*(</t></is></c>)'
        new = re.sub(pattern2, rf'\g<1>{escaped}\g<2>', content)
    return new

def generate_xlsx(template_path, updates_num, updates_str):
    out = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                # sheet1, sheet2 둘 다 수정 (NASN은 PL도 inlineStr)
                if item in ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']:
                    content = data.decode('utf-8')
                    for cell_ref, value in updates_str.items():
                        content = set_inline_str_cell(content, cell_ref, value)
                    # 수량/단가는 sheet1(Invoice)만
                    if item == 'xl/worksheets/sheet1.xml':
                        for cell_ref, value in updates_num.items():
                            content = set_number_cell(content, cell_ref, value)
                    data = content.encode('utf-8')
                zout.writestr(item, data)
    out.seek(0)
    return out

def get_updates(data):
    cid = data.get('client_id')
    items = data.get('items', [])
    inv_no = data.get('inv_no', 'TS-000000-01')
    date_str = data.get('date', '2026-01-01')
    updates_num, updates_str = {}, {}

    if cid == 'china':
        updates_str['F3'] = f"INVOICE NO : {inv_no}"
        updates_str['F4'] = f"INVOICE DATE : {fmt_ordinal(date_str)}"
        if len(items) >= 1:
            updates_num['G23'] = items[0]['qty']; updates_num['H23'] = items[0]['uprice']
        if len(items) >= 2:
            updates_num['G25'] = items[1]['qty']; updates_num['H25'] = items[1]['uprice']
    elif cid == 'india':
        updates_str['F3'] = f"INVOICE NO : {inv_no}"
        updates_str['F4'] = f"INVOICE DATE : {fmt_ordinal(date_str)}"
        if len(items) >= 1:
            updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
        if len(items) >= 2:
            updates_num['G25'] = items[1]['qty']; updates_num['H25'] = items[1]['uprice']
    elif cid == 'nasn':
        updates_str['F3'] = f"INVOICE NO : {inv_no}"
        updates_str['F4'] = f"INVOICE DATE : {fmt_upper(date_str)}"
        if len(items) >= 1:
            updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
    elif cid == 'mexico':
        updates_str['F3'] = f"INVOICE NO : {inv_no}"
        updates_str['F4'] = f"INVOICE DATE : {fmt_ordinal(date_str)}"
        if len(items) >= 1:
            updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
        if len(items) >= 2:
            updates_num['G26'] = items[1]['qty']; updates_num['H26'] = items[1]['uprice']

    return updates_num, updates_str

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    cid = data.get('client_id')
    if cid not in TEMPLATES:
        return jsonify({'error': f'Unknown client: {cid}'}), 400
    inv_no = data.get('inv_no', 'TS-000000-01')
    updates_num, updates_str = get_updates(data)
    result = generate_xlsx(TEMPLATES[cid], updates_num, updates_str)
    return send_file(result, as_attachment=True,
                     download_name=f"{inv_no}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    print("✓ Invoice Generator Server 시작")
    app.run(host='0.0.0.0', port=5050, debug=False)
