#!/usr/bin/env python3
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import io, os, zipfile, re, math
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
    # 수식 제거하고 값으로 교체
    pattern = rf'<c r="{cell_ref}"([^>]*)><f>[^<]*</f><v>[^<]*</v></c>'
    replacement = rf'<c r="{cell_ref}"\1><v>{value}</v></c>'
    new = re.sub(pattern, replacement, content)
    if new == content:
        pattern2 = rf'(<c r="{cell_ref}"[^>]*>(?:<f>[^<]*</f>)?)<v>[^<]*</v>'
        new = re.sub(pattern2, rf'\g<1><v>{value}</v>', content)
    if new == content:
        new = re.sub(rf'(<c r="{cell_ref}"[^>]*>)<v>[^<]*</v>', rf'\g<1><v>{value}</v>', content)
    return new

def remove_formula_set_value(content, cell_ref, value):
    """수식 셀을 값 셀로 교체"""
    # <c r="G23" s="X"><f>수식</f><v>이전값</v></c> → <c r="G23" s="X"><v>새값</v></c>
    pattern = rf'<c r="{cell_ref}"([^>]*)>(<f>[^<]*</f>)?<v>[^<]*</v></c>'
    replacement = rf'<c r="{cell_ref}"\1><v>{value}</v></c>'
    new = re.sub(pattern, replacement, content)
    if new == content:
        # 값이 없는 경우
        pattern2 = rf'<c r="{cell_ref}"([^>]*)><f>[^<]*</f></c>'
        replacement2 = rf'<c r="{cell_ref}"\1><v>{value}</v></c>'
        new = re.sub(pattern2, replacement2, content)
    return new

def update_shared_string(ss_xml, index, new_value):
    items = list(re.finditer(r'<si>(.*?)</si>', ss_xml, re.DOTALL))
    if index >= len(items):
        return ss_xml
    match = items[index]
    old_si = match.group(0)
    new_si = re.sub(r'(<t>)[^<]*(</t>)', rf'\g<1>{new_value}\g<2>', old_si, count=1)
    return ss_xml[:match.start()] + new_si + ss_xml[match.end():]

def calc_pl_values(cid, items_data):
    """PL 시트에 직접 써야 할 수량 값 계산"""
    pl = {}
    if cid == 'china':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 12
        pallets = max(1, math.ceil(boxes / 27))
        pl['G23'] = q1       # BG6900Z000 qty
        pl['G24'] = q2       # BG69003800 qty
        pl['G25'] = trays    # PE Tray
        pl['G26'] = boxes    # Box
        pl['G27'] = pallets  # Pallet

    elif cid == 'india':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 12
        pallets = max(1, math.ceil(boxes / 27))
        pl['F22'] = q1       # SM100 qty
        pl['F23'] = q2       # SM110 qty
        pl['F24'] = trays    # PE Tray
        pl['F25'] = boxes    # Box
        pl['F26'] = pallets  # Pallet

    elif cid == 'nasn':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        boxes = math.ceil(q1 / 1034)
        trays = boxes * 11
        pallets = max(1, math.ceil(boxes / 27))
        pl['F22'] = q1       # SNA120 qty
        pl['F23'] = trays    # PE Tray
        pl['F24'] = boxes    # Box
        pl['F25'] = pallets  # Pallet

    elif cid == 'mexico':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 11
        pallets = max(1, math.ceil(boxes / 27))
        pl['F21'] = q1       # SM100 qty
        pl['F22'] = q2       # SM110 qty
        pl['F23'] = trays    # PE Tray
        pl['F24'] = boxes    # Box
        pl['F25'] = pallets  # Pallet

    return pl

def generate_xlsx(template_path, cid, inv_no, inv_date, updates_num, items_data):
    idx = SS_INDEX[cid]
    pl_values = calc_pl_values(cid, items_data)

    out = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)

                if item == 'xl/workbook.xml':
                    content = data.decode('utf-8')
                    if 'fullCalcOnLoad' not in content:
                        content = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', content)
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

                elif item == 'xl/worksheets/sheet2.xml':
                    content = data.decode('utf-8')
                    for cell_ref, value in pl_values.items():
                        content = remove_formula_set_value(content, cell_ref, value)
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

    items_data = data.get('items', [])
    inv_no = data.get('inv_no', 'TS-000000-01')
    date_str = data.get('date', '2026-01-01')
    inv_date = fmt_upper(date_str) if cid == 'nasn' else fmt_ordinal(date_str)

    updates_num = {}
    if cid == 'china':
        if len(items_data) >= 1: updates_num['G23'] = items_data[0]['qty']; updates_num['H23'] = items_data[0]['uprice']
        if len(items_data) >= 2: updates_num['G25'] = items_data[1]['qty']; updates_num['H25'] = items_data[1]['uprice']
    elif cid == 'india':
        if len(items_data) >= 1: updates_num['G24'] = items_data[0]['qty']; updates_num['H24'] = items_data[0]['uprice']
        if len(items_data) >= 2: updates_num['G25'] = items_data[1]['qty']; updates_num['H25'] = items_data[1]['uprice']
    elif cid == 'nasn':
        if len(items_data) >= 1: updates_num['G24'] = items_data[0]['qty']; updates_num['H24'] = items_data[0]['uprice']
    elif cid == 'mexico':
        if len(items_data) >= 1: updates_num['G24'] = items_data[0]['qty']; updates_num['H24'] = items_data[0]['uprice']
        if len(items_data) >= 2: updates_num['G26'] = items_data[1]['qty']; updates_num['H26'] = items_data[1]['uprice']

    result = generate_xlsx(TEMPLATES[cid], cid, inv_no, inv_date, updates_num, items_data)
    return send_file(result, as_attachment=True,
                     download_name=f"{inv_no}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'version': 'v6-pl-direct'})

if __name__ == '__main__':
    print("✓ Invoice Generator Server 시작")
    app.run(host='0.0.0.0', port=5050, debug=False)
