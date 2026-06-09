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

SS_INDEX = {
    'china':  {'inv_no': 76, 'inv_date': 77},
    'india':  {'inv_no': 69, 'inv_date': 70},
    'nasn':   {'inv_no': 83, 'inv_date': 82},
    'mexico': {'inv_no': 64, 'inv_date': 65},
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

def update_shared_string(ss_xml, index, new_value):
    items = list(re.finditer(r'<si>(.*?)</si>', ss_xml, re.DOTALL))
    if index >= len(items):
        return ss_xml
    match = items[index]
    old_si = match.group(0)
    new_si = re.sub(r'(<t>)[^<]*(</t>)', rf'\g<1>{new_value}\g<2>', old_si, count=1)
    return ss_xml[:match.start()] + new_si + ss_xml[match.end():]

def add_full_calc(wb_xml):
    """workbook.xml에 fullCalcOnLoad 추가 → 파일 열 때 수식 자동 재계산"""
    if 'fullCalcOnLoad' in wb_xml:
        return wb_xml
    return re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb_xml)

def generate_xlsx(template_path, cid, inv_no, inv_date, updates_num):
    idx = SS_INDEX[cid]
    out = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)

                if item == 'xl/workbook.xml':
                    content = data.decode('utf-8')
                    content = add_full_calc(content)
                    data = content.encode('utf-8')

                elif item == 'xl/sharedStrings.xml':
                    ss = data.decode('utf-8')
                    ss = update_shared_string(ss, idx['inv_no'], f"INVOICE NO : {inv_no}")
                    ss = update_shared_string(ss, idx['inv_date'], f"INVOICE DATE : {inv_date}")
                    data = ss.encode('utf-8')

                elif item == 'xl/worksheets/sheet1.xml':
                    content = data.decode('utf-8')
                    for cell_ref, value in updates_num.items():
                        content = set_number_cell(content, cell_ref, value)
                    data = content.encode('utf-8')

                zout.writestr(item, data)
    out.seek(0)
    return out

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    cid = data.get('client_id')
    if cid not in TEMPLATES:
        return jsonify({'error': f'Unknown client: {cid}'}), 400

    items = data.get('items', [])
    inv_no = data.get('inv_no', 'TS-000000-01')
    date_str = data.get('date', '2026-01-01')
    inv_date = fmt_upper(date_str) if cid == 'nasn' else fmt_ordinal(date_str)

    updates_num = {}
    if cid == 'china':
        if len(items) >= 1: updates_num['G23'] = items[0]['qty']; updates_num['H23'] = items[0]['uprice']
        if len(items) >= 2: updates_num['G25'] = items[1]['qty']; updates_num['H25'] = items[1]['uprice']
    elif cid == 'india':
        if len(items) >= 1: updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
        if len(items) >= 2: updates_num['G25'] = items[1]['qty']; updates_num['H25'] = items[1]['uprice']
    elif cid == 'nasn':
        if len(items) >= 1: updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
    elif cid == 'mexico':
        if len(items) >= 1: updates_num['G24'] = items[0]['qty']; updates_num['H24'] = items[0]['uprice']
        if len(items) >= 2: updates_num['G26'] = items[1]['qty']; updates_num['H26'] = items[1]['uprice']

    result = generate_xlsx(TEMPLATES[cid], cid, inv_no, inv_date, updates_num)
    return send_file(result, as_attachment=True,
                     download_name=f"{inv_no}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'version': 'v5-fullcalc'})

if __name__ == '__main__':
    print("✓ Invoice Generator Server 시작")
    app.run(host='0.0.0.0', port=5050, debug=False)
