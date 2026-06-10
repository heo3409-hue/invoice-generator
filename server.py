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

W_SENSOR = 5.2
W_TRAY   = 155
W_BOX    = 800
W_PALLET = 7200

def ceildiv(a, b):
    return math.ceil(a / b) if b else 0

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
    """수식 포함 여부 상관없이 값으로 교체"""
    # 수식+값 셀
    pattern = rf'<c r="{cell_ref}"([^>]*)>(?:<f>[^<]*</f>)?<v>[^<]*</v></c>'
    new = re.sub(pattern, rf'<c r="{cell_ref}"\1><v>{value}</v></c>', content)
    if new == content:
        # 수식만 있고 값 없는 셀
        pattern2 = rf'<c r="{cell_ref}"([^>]*)><f>[^<]*</f></c>'
        new = re.sub(pattern2, rf'<c r="{cell_ref}"\1><v>{value}</v></c>', content)
    if new == content:
        # 값만 있는 셀
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
    sheet1_num = {}
    sheet2_num = {}
    ss_updates = {}

    if cid == 'china':
        q1 = items_data[0]['qty']    if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty']    if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes   = ceildiv(q1 + q2, 1034)
        trays   = ceildiv(boxes, 12)
        pallets = ceildiv(boxes, 27)
        gw1     = round(q1 * W_SENSOR, 1)
        gw2     = round(q2 * W_SENSOR, 1)
        gw_tr   = round(trays * W_TRAY, 1)
        gw_bx   = round(boxes * W_BOX, 1)
        gw_pl   = round(pallets * W_PALLET, 1)
        total_gw= round(gw1+gw2+gw_tr+gw_bx+gw_pl, 1)
        amt1    = round(q1 * u1, 2)
        amt2    = round(q2 * u2, 2)
        total   = round(amt1 + amt2, 2)

        sheet1_num.update({
            'G23': q1, 'H23': u1, 'I23': amt1,   # BG6900Z000
            'G25': q2, 'H25': u2,                 # BG69003800 (I25는 수식없음)
        })
        # I25도 있으면 값으로
        sheet1_num['I25'] = round(q2 * u2, 2)

        sheet2_num.update({
            'G23': q1,   'G24': q2,   'G25': trays, 'G26': boxes, 'G27': pallets,
            'I23': gw1,  'I24': gw2,  'I25': gw_tr, 'I26': gw_bx,'I27': gw_pl,
            'I28': total_gw
        })
        ss_updates[78] = f" - Total gross weight : {round(total_gw/1000, 3)}kg"
        ss_updates[79] = f" - Box amount : {boxes}EA"
        ss_updates[80] = f" - Pallet amount : {pallets}EA"
        ss_updates[81] = f"TOTAL AMOUNT : BOX ({boxes}EA), Pallet ({pallets}EA)   USD {fmt_usd(total)}"

    elif cid == 'india':
        q1 = items_data[0]['qty']    if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty']    if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes   = ceildiv(q1 + q2, 1034)
        trays   = ceildiv(boxes, 12)
        pallets = ceildiv(boxes, 27)
        gw1     = round(q1 * W_SENSOR, 1)
        gw2     = round(q2 * W_SENSOR, 1)
        gw_tr   = round(trays * W_TRAY, 1)
        gw_bx   = round(boxes * W_BOX, 1)
        gw_pl   = round(pallets * W_PALLET, 1)
        total_gw= round(gw1+gw2+gw_tr+gw_bx+gw_pl, 1)
        amt1    = round(q1 * u1, 2)
        amt2    = round(q2 * u2, 2)
        total   = round(amt1 + amt2, 2)

        sheet1_num.update({
            'G24': q1, 'H24': u1, 'I24': amt1,
            'G25': q2, 'H25': u2, 'I25': amt2,
        })
        sheet2_num.update({
            'F22': q1,  'F23': q2,  'F24': trays, 'F25': boxes, 'F26': pallets,
            'H22': gw1, 'H23': gw2, 'H24': gw_tr, 'H25': gw_bx,'H26': gw_pl,
            'H27': total_gw,
            'G29': pallets, 'H29': round(total_gw/1000, 3)
        })
        ss_updates[67] = f" - Total gross weight : {round(total_gw/1000)}kg"
        ss_updates[68] = f" - BOX AMOUNT : {boxes}EA"
        ss_updates[71] = f"TOTAL AMOUNT : PALLET({pallets}EA)  BOX ({boxes}EA)   USD ${fmt_usd(total)}"

    elif cid == 'nasn':
        q1 = items_data[0]['qty']    if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        boxes   = ceildiv(q1, 1034)
        trays   = ceildiv(boxes, 12)
        pallets = ceildiv(boxes, 27)
        gw1     = round(q1 * W_SENSOR, 1)
        gw_tr   = round(trays * W_TRAY, 1)
        gw_bx   = round(boxes * W_BOX, 1)
        gw_pl   = round(pallets * W_PALLET, 1)
        total_gw= round(gw1+gw_tr+gw_bx+gw_pl, 1)
        amt1    = round(q1 * u1, 2)

        sheet1_num.update({
            'G24': q1, 'H24': u1, 'I24': amt1,
            'D30': boxes, 'D32': round(total_gw/1000),
            'D36': pallets, 'F36': boxes, 'I36': amt1
        })
        sheet2_num.update({
            'F22': q1,  'F23': trays, 'F24': boxes, 'F25': pallets,
            'H22': gw1, 'H23': gw_tr, 'H24': gw_bx, 'H25': gw_pl,
            'H26': total_gw,
            'G28': pallets, 'H28': round(total_gw/1000, 3)
        })

    elif cid == 'mexico':
        q1 = items_data[0]['qty']    if len(items_data) >= 1 else 0
        u1 = items_data[0]['uprice'] if len(items_data) >= 1 else 0
        q2 = items_data[1]['qty']    if len(items_data) >= 2 else 0
        u2 = items_data[1]['uprice'] if len(items_data) >= 2 else 0
        boxes   = ceildiv(q1 + q2, 1034)
        trays   = ceildiv(boxes, 12)
        pallets = ceildiv(boxes, 27)
        gw1     = round(q1 * W_SENSOR, 1)
        gw2     = round(q2 * W_SENSOR, 1)
        gw_tr   = round(trays * W_TRAY, 1)
        gw_bx   = round(boxes * W_BOX, 1)
        gw_pl   = round(pallets * W_PALLET, 1)
        total_gw= round(gw1+gw2+gw_tr+gw_bx+gw_pl, 1)
        amt1    = round(q1 * u1, 2)
        amt2    = round(q2 * u2, 2)
        total   = round(amt1 + amt2, 2)

        sheet1_num.update({
            'G24': q1, 'H24': u1, 'I24': amt1,
            'G26': q2, 'H26': u2, 'I26': amt2,
        })
        sheet2_num.update({
            'F21': q1,  'F22': q2,  'F23': trays, 'F24': boxes, 'F25': pallets,
            'H21': gw1, 'H22': gw2, 'H23': gw_tr, 'H24': gw_bx,'H25': gw_pl,
            'H26': total_gw
        })
        ss_updates[61] = f" - BOX AMOUNT : {boxes}EA"
        ss_updates[62] = f" - Total gross weight : {round(total_gw/1000)}kg"
        ss_updates[63] = f"TOTAL AMOUNT : Pallet ({pallets}EA), BOX ({boxes}EA)   USD {fmt_usd(total)}"

    return sheet1_num, sheet2_num, ss_updates

def generate_xlsx(template_path, cid, inv_no, inv_date, items_data):
    idx = SS_INDEX[cid]
    sheet1_num, sheet2_num, ss_updates = calc_all(cid, items_data)

    out = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():

                # calcChain.xml 제거 → 팝업 오류 해결
                if item == 'xl/calcChain.xml':
                    continue

                data = zin.read(item)

                if item == 'xl/workbook.xml':
                    content = data.decode('utf-8')
                    if 'fullCalcOnLoad' not in content:
                        content = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', content)
                    data = content.encode('utf-8')

                elif item == '[Content_Types].xml':
                    # calcChain 참조도 제거
                    content = data.decode('utf-8')
                    content = re.sub(r'<Override[^>]*calcChain[^>]*/>', '', content)
                    data = content.encode('utf-8')

                elif item == 'xl/sharedStrings.xml':
                    ss = data.decode('utf-8')
                    ss = update_shared_string(ss, idx['inv_no'],   f"INVOICE NO : {inv_no}")
                    ss = update_shared_string(ss, idx['inv_date'], f"INVOICE DATE : {inv_date}")
                    for si, val in ss_updates.items():
                        ss = update_shared_string(ss, si, val)
                    data = ss.encode('utf-8')

                elif item == 'xl/worksheets/sheet1.xml':
                    content = data.decode('utf-8')
                    for ref, val in sheet1_num.items():
                        content = set_cell_value(content, ref, val)
                    data = content.encode('utf-8')

                elif item == 'xl/worksheets/sheet2.xml':
                    content = data.decode('utf-8')
                    for ref, val in sheet2_num.items():
                        content = set_cell_value(content, ref, val)
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
    inv_no   = data.get('inv_no', 'TS-000000-01')
    date_str = data.get('date', '2026-01-01')
    inv_date = fmt_upper(date_str) if cid == 'nasn' else fmt_ordinal(date_str)
    result = generate_xlsx(TEMPLATES[cid], cid, inv_no, inv_date, items_data)
    return send_file(result, as_attachment=True,
                     download_name=f"{inv_no}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'version': 'v10-no-calcchain'})

if __name__ == '__main__':
    print("✓ Invoice Generator Server 시작")
    app.run(host='0.0.0.0', port=5050, debug=False)
