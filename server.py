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

W = {'sensor': 5.2, 'tray_std': 155, 'tray_india': 190,
     'box_china': 100, 'box_india': 850, 'box_nasn': 800, 'box_mexico': 800,
     'pallet': 7200}

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

def set_cell_value(content, cell_ref, value):
    pattern = rf'<c r="{cell_ref}"([^>]*)>(?:<f>[^<]*</f>)?<v>[^<]*</v></c>'
    new = re.sub(pattern, rf'<c r="{cell_ref}"\1><v>{value}</v></c>', content)
    if new == content:
        pattern2 = rf'<c r="{cell_ref}"([^>]*)><f>[^<]*</f></c>'
        new = re.sub(pattern2, rf'<c r="{cell_ref}"\1><v>{value}</v></c>', content)
    if new == content:
        pattern3 = rf'(<c r="{cell_ref}"[^>]*>)<v>[^<]*</v>'
        new = re.sub(pattern3, rf'\g<1><v>{value}</v>', content)
    return new

def update_shared_string(ss_xml, index, new_value):
    items = list(re.finditer(r'<si>(.*?)</si>', ss_xml, re.DOTALL))
    if index >= len(items):
        return ss_xml
    match = items[index]
    old_si = match.group(0)
    new_si = re.sub(r'(<t[^>]*>)[^<]*(</t>)', rf'\g<1>{new_value}\g<2>', old_si, count=1)
    return ss_xml[:match.start()] + new_si + ss_xml[match.end():]

def fmt_usd(v):
    return f"{v:,.2f}"

def calc_all(cid, items_data):
    """Invoice sheet1, PL sheet2, sharedStrings 업데이트 값 모두 계산"""
    sheet1_num = {}   # 숫자 셀
    sheet2_num = {}   # PL 숫자 셀
    ss_updates = {}   # shared string 인덱스: 새 텍스트

    if cid == 'china':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 12
        pallets = max(1, math.ceil(boxes / 27))
        gross_g = q1*W['sensor'] + q2*W['sensor'] + trays*W['tray_std'] + boxes*W['box_china'] + pallets*W['pallet']
        gross_kg = round(gross_g / 1000, 3)
        total = round(q1*u1 + q2*u2, 2)
        sheet1_num.update({'G23': q1, 'H23': u1, 'G25': q2, 'H25': u2})
        sheet2_num.update({
            'G23': q1, 'G24': q2, 'G25': trays, 'G26': boxes, 'G27': pallets,
            'I23': round(q1*W['sensor'],1), 'I24': round(q2*W['sensor'],1),
            'I25': round(trays*W['tray_std'],1), 'I26': round(boxes*W['box_china'],1),
            'I27': round(pallets*W['pallet'],1), 'I28': round(gross_g,1)
        })
        ss_updates[78] = f" - Total gross weight : {gross_kg}kg"
        ss_updates[80] = f" - Pallet amount : {pallets}EA"
        ss_updates[81] = f"TOTAL AMOUNT : BOX ({boxes}EA), Pallet ({pallets}EA)   USD {fmt_usd(total)}"
        # C33 Box amount (idx 찾기 필요 - 일단 직접 처리)

    elif cid == 'india':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 12
        pallets = max(1, math.ceil(boxes / 27))
        gross_g = q1*W['sensor'] + q2*W['sensor'] + trays*W['tray_india'] + boxes*W['box_india'] + pallets*W['pallet']
        gross_kg = round(gross_g / 1000, 0)
        total = round(q1*u1 + q2*u2, 2)
        sheet1_num.update({'G24': q1, 'H24': u1, 'G25': q2, 'H25': u2})
        sheet2_num.update({
            'F22': q1, 'F23': q2, 'F24': trays, 'F25': boxes, 'F26': pallets,
            'H22': round(q1*W['sensor'],1), 'H23': round(q2*W['sensor'],1),
            'H24': round(trays*W['tray_india'],1), 'H25': round(boxes*W['box_india'],1),
            'H26': round(pallets*W['pallet'],1), 'H27': round(gross_g,1),
            'G29': pallets, 'H29': round(gross_g/1000, 3)
        })
        ss_updates[67] = f" - Total gross weight : {int(gross_kg)}kg"
        ss_updates[68] = f" - BOX AMOUNT : {boxes}EA"
        ss_updates[71] = f"TOTAL AMOUNT : PALLET({pallets}EA)  BOX ({boxes}EA)   USD ${fmt_usd(total)}"

    elif cid == 'nasn':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        boxes = math.ceil(q1 / 1034)
        trays = boxes * 11
        pallets = max(1, math.ceil(boxes / 27))
        gross_g = q1*W['sensor'] + trays*W['tray_std'] + boxes*W['box_nasn'] + pallets*W['pallet']
        gross_kg = round(gross_g / 1000, 0)
        total = round(q1*u1, 2)
        sheet1_num.update({'G24': q1, 'H24': u1})
        # NASN: D30=boxes, D32=gross_kg, D36=pallets, F36=boxes, H36=total(수식→값)
        sheet1_num.update({'D30': boxes, 'D32': int(gross_kg), 'D36': pallets, 'F36': boxes})
        sheet2_num.update({
            'F22': q1, 'F23': trays, 'F24': boxes, 'F25': pallets,
            'H22': round(q1*W['sensor'],1), 'H23': round(trays*W['tray_std'],1),
            'H24': round(boxes*W['box_nasn'],1), 'H25': round(pallets*W['pallet'],1),
            'H26': round(gross_g,1), 'G28': pallets, 'H28': round(gross_g/1000,3)
        })

    elif cid == 'mexico':
        q1 = items_data[0]['qty'] if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty'] if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes = math.ceil((q1 + q2) / 1034)
        trays = boxes * 11
        pallets = max(1, math.ceil(boxes / 27))
        gross_g = q1*W['sensor'] + q2*W['sensor'] + trays*W['tray_std'] + boxes*W['box_mexico'] + pallets*W['pallet']
        gross_kg = round(gross_g / 1000, 0)
        total = round(q1*u1 + q2*u2, 2)
        sheet1_num.update({'G24': q1, 'H24': u1, 'G26': q2, 'H26': u2})
        sheet2_num.update({
            'F21': q1, 'F22': q2, 'F23': trays, 'F24': boxes, 'F25': pallets,
            'H21': round(q1*W['sensor'],1), 'H22': round(q2*W['sensor'],1),
            'H23': round(trays*W['tray_std'],1), 'H24': round(boxes*W['box_mexico'],1),
            'H25': round(pallets*W['pallet'],1), 'H26': round(gross_g,1)
        })
        ss_updates[61] = f" - BOX AMOUNT : {boxes}EA"
        ss_updates[62] = f" - Total gross weight : {int(gross_kg)}kg"
        ss_updates[63] = f"TOTAL AMOUNT : Pallet ({pallets}EA), BOX ({boxes}EA)   USD {fmt_usd(total)}"

    return sheet1_num, sheet2_num, ss_updates

def generate_xlsx(template_path, cid, inv_no, inv_date, items_data):
    idx = SS_INDEX[cid]
    sheet1_num, sheet2_num, ss_updates = calc_all(cid, items_data)

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
                    for ss_idx, ss_val in ss_updates.items():
                        ss = update_shared_string(ss, ss_idx, ss_val)
                    data = ss.encode('utf-8')

                elif item == 'xl/worksheets/sheet1.xml':
                    content = data.decode('utf-8')
                    for cell_ref, value in sheet1_num.items():
                        content = set_cell_value(content, cell_ref, value)
                    data = content.encode('utf-8')

                elif item == 'xl/worksheets/sheet2.xml':
                    content = data.decode('utf-8')
                    for cell_ref, value in sheet2_num.items():
                        content = set_cell_value(content, cell_ref, value)
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
    result = generate_xlsx(TEMPLATES[cid], cid, inv_no, inv_date, items_data)
    return send_file(result, as_attachment=True,
                     download_name=f"{inv_no}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'version': 'v8-total-amount'})

if __name__ == '__main__':
    print("✓ Invoice Generator Server 시작")
    app.run(host='0.0.0.0', port=5050, debug=False)
