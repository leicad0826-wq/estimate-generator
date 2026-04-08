import streamlit as st
import subprocess, os, shutil, glob, zipfile, re, tempfile
from lxml import etree
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime, timedelta
from PIL import Image
import io

st.set_page_config(page_title="最終見積書 自動生成", page_icon="📄", layout="centered")

# ============================================================
# スタイル
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap');

html, body, [class*="css"] { font-family: 'Noto Sans JP', sans-serif; }

.main { background: #f7f8fa; }

.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 16px;
    padding: 40px 32px;
    color: white;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -20%;
    width: 400px;
    height: 400px;
    background: radial-gradient(circle, rgba(99,179,237,0.15) 0%, transparent 70%);
    border-radius: 50%;
}
.hero h1 { font-size: 1.8rem; font-weight: 700; margin: 0 0 8px 0; letter-spacing: -0.5px; }
.hero p  { font-size: 0.95rem; opacity: 0.7; margin: 0; }

.step-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 16px;
    border: 1px solid #e8ecf0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.step-num {
    display: inline-block;
    background: #0f3460;
    color: white;
    border-radius: 50%;
    width: 28px;
    height: 28px;
    line-height: 28px;
    text-align: center;
    font-size: 0.85rem;
    font-weight: 700;
    margin-right: 10px;
}
.step-title { font-weight: 600; font-size: 1rem; color: #1a1a2e; }

.success-box {
    background: linear-gradient(135deg, #e6fffa, #f0fff4);
    border: 1px solid #9ae6b4;
    border-radius: 12px;
    padding: 20px 24px;
    margin-top: 16px;
}
.warning-box {
    background: #fffbeb;
    border: 1px solid #fcd34d;
    border-radius: 10px;
    padding: 14px 18px;
    font-size: 0.9rem;
    color: #92400e;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5e0 !important;
    border-radius: 12px !important;
    background: #f7f8fa !important;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# ヘッダー
# ============================================================
st.markdown("""
<div class="hero">
    <h1>📄 最終見積書 自動生成</h1>
    <p>見積算出表（.xlsb）をアップロードするだけで最終見積書を自動生成します</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# generate_estimate の関数群（インライン）
# ============================================================
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_S   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'
SLOTS  = [24, 26, 28, 30, 32, 34, 36, 38, 40]

def parse_betto(text):
    items = []
    if not text: return items
    for line in str(text).strip().split('\n'):
        line = line.strip()
        if not line: continue
        m = re.match(r'^(.+?)\s+(\d+)([^\d\s]+)\s+(\d+)円?$', line)
        if m:
            items.append({'name': m.group(1).strip(), 'qty': int(m.group(2)),
                          'unit': m.group(3), 'price': int(m.group(4))})
        else:
            parts = line.split('\u3000')
            if len(parts) == 3:
                name = parts[0].strip()
                qu = re.match(r'(\d+)(.+)', parts[1].strip())
                price = re.sub(r'[^\d]', '', parts[2])
                if qu and price:
                    items.append({'name': name, 'qty': int(qu.group(1)),
                                  'unit': qu.group(2), 'price': int(price)})
                else:
                    items.append({'name': line, 'qty': 1, 'unit': '式', 'price': 0})
            else:
                items.append({'name': line, 'qty': 1, 'unit': '式', 'price': 0})
    return items

def read_estimate(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb['見積算出表']
    anken = str(ws['D7'].value or '')
    anken_stripped = re.sub(r'^\d{3}', '', anken).strip()
    nouki = ws['D10'].value
    if isinstance(nouki, datetime): nouki = nouki.strftime('%Y/%m/%d')
    honnohin = ws['D43'].value or 0
    yubi     = ws['K43'].value or 0
    tanka = None
    for r in range(100, 160):
        for c in range(14, 20):
            if ws.cell(row=r, column=c).value == '最終単価':
                tanka = ws.cell(row=r, column=17).value; break
        if tanka: break
    betto_text = None
    for r in range(100, 160):
        for c in range(18, 24):
            if ws.cell(row=r, column=c).value == '別途請求内容':
                betto_text = ws.cell(row=r+1, column=c).value; break
        if betto_text is not None: break
    sku = ws['D15'].value or 1
    insatsu_men = ws['D16'].value or ''
    insatsu_ura = ws['D18'].value or ''
    item_info   = str(ws['D8'].value or '')
    size_match  = re.search(r'(\d+[×x]\d+mm)', item_info)
    if size_match:
        size_label = '枠サイズ'; size_str = size_match.group(1)
    else:
        size_label = 'デザイン本体サイズ'; size_str = item_info
    insatsu_display = f"{insatsu_men}面）{insatsu_ura}" if insatsu_men and insatsu_ura else str(insatsu_ura)
    return dict(anken=anken, anken_stripped=anken_stripped, nouki=str(nouki),
                honnohin=honnohin, yubi=yubi, tanka=tanka,
                betto_items=parse_betto(betto_text), sku=sku,
                size_label=size_label, size_str=size_str, insatsu=insatsu_display)

def get_images_from_sheet(xlsx_path, sheet_name='種別'):
    images = []
    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml  = etree.fromstring(z.read('xl/workbook.xml'))
        wb_rels = etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid_to_target = {r.get('Id'): r.get('Target') for r in wb_rels}
        sheet_file = None
        for sh in wb_xml.findall(f'{{{NS_S}}}sheets/{{{NS_S}}}sheet'):
            if sh.get('name') == sheet_name:
                rid = sh.get(f'{{{NS_R}}}id')
                sheet_file = rid_to_target.get(rid, '').split('/')[-1]
                break
        if not sheet_file: return images
        rels_path = f'xl/worksheets/_rels/{sheet_file}.rels'
        if rels_path not in z.namelist(): return images
        sheet_rels = etree.fromstring(z.read(rels_path))
        drawing_target = None
        for rel in sheet_rels:
            if 'drawing' in rel.get('Type','') and 'vml' not in rel.get('Target',''):
                drawing_target = rel.get('Target'); break
        if not drawing_target: return images
        drawing_path      = 'xl/' + drawing_target.lstrip('../')
        drawing_rels_path = drawing_path.replace('drawings/','drawings/_rels/') + '.rels'
        if drawing_path not in z.namelist(): return images
        drawing = etree.fromstring(z.read(drawing_path))
        d_rels  = etree.fromstring(z.read(drawing_rels_path)) if drawing_rels_path in z.namelist() else None
        rid_map = {r.get('Id'): r.get('Target') for r in d_rels} if d_rels is not None else {}
        for anchor in drawing:
            fe  = anchor.find(f'{{{NS_XDR}}}from')
            pic = anchor.find(f'{{{NS_XDR}}}pic')
            if fe is None or pic is None: continue
            fc = int(fe.find(f'{{{NS_XDR}}}col').text)
            fr = int(fe.find(f'{{{NS_XDR}}}row').text)
            blip = pic.find(f'.//{{{NS_A}}}blip')
            if blip is None: continue
            rid = blip.get(f'{{{NS_R}}}embed','')
            img_rel  = rid_map.get(rid,'')
            img_path = 'xl/media/' + img_rel.split('/')[-1]
            if img_path in z.namelist():
                images.append((z.read(img_path), fc, fr))
    images.sort(key=lambda x: (x[2], x[1]))
    return [(d,) for d,_,_ in images]

def make_pic_anchor(rid, img_id, img_name, col_from, row_from, col_to, row_to):
    anchor = etree.Element(f'{{{NS_XDR}}}twoCellAnchor'); anchor.set('editAs','oneCell')
    fe = etree.SubElement(anchor, f'{{{NS_XDR}}}from')
    etree.SubElement(fe, f'{{{NS_XDR}}}col').text    = str(col_from)
    etree.SubElement(fe, f'{{{NS_XDR}}}colOff').text = '25400'
    etree.SubElement(fe, f'{{{NS_XDR}}}row').text    = str(row_from)
    etree.SubElement(fe, f'{{{NS_XDR}}}rowOff').text = '88339'
    te = etree.SubElement(anchor, f'{{{NS_XDR}}}to')
    etree.SubElement(te, f'{{{NS_XDR}}}col').text    = str(col_to)
    etree.SubElement(te, f'{{{NS_XDR}}}colOff').text = '127000'
    etree.SubElement(te, f'{{{NS_XDR}}}row').text    = str(row_to)
    etree.SubElement(te, f'{{{NS_XDR}}}rowOff').text = '83525'
    pic = etree.SubElement(anchor, f'{{{NS_XDR}}}pic')
    nv  = etree.SubElement(pic, f'{{{NS_XDR}}}nvPicPr')
    cNvPr = etree.SubElement(nv, f'{{{NS_XDR}}}cNvPr')
    cNvPr.set('id', str(img_id)); cNvPr.set('name', img_name)
    cNvPicPr = etree.SubElement(nv, f'{{{NS_XDR}}}cNvPicPr')
    locks = etree.SubElement(cNvPicPr, f'{{{NS_A}}}picLocks'); locks.set('noChangeAspect','1')
    bf = etree.SubElement(pic, f'{{{NS_XDR}}}blipFill'); bf.set('rotWithShape','1')
    blip = etree.SubElement(bf, f'{{{NS_A}}}blip'); blip.set(f'{{{NS_R}}}embed', rid)
    etree.SubElement(bf, f'{{{NS_A}}}stretch')
    spPr = etree.SubElement(pic, f'{{{NS_XDR}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{NS_A}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{NS_A}}}off'); off.set('x','0'); off.set('y','0')
    ext = etree.SubElement(xfrm, f'{{{NS_A}}}ext'); ext.set('cx','1900000'); ext.set('cy','1400000')
    pg = etree.SubElement(spPr, f'{{{NS_A}}}prstGeom'); pg.set('prst','rect')
    etree.SubElement(pg, f'{{{NS_A}}}avLst')
    etree.SubElement(anchor, f'{{{NS_XDR}}}clientData')
    return anchor

def build_drawing(tmpl_drawing_bytes, tmpl_rels_bytes, images_list, all_files, media_prefix):
    drawing = etree.fromstring(tmpl_drawing_bytes)
    rels    = etree.fromstring(tmpl_rels_bytes)
    for anchor in list(drawing):
        fe  = anchor.find(f'{{{NS_XDR}}}from')
        pic = anchor.find(f'{{{NS_XDR}}}pic')
        if fe is None or pic is None: continue
        if int(fe.find(f'{{{NS_XDR}}}row').text) == 45:
            ne = pic.find(f'.//{{{NS_XDR}}}cNvPr')
            if ne is not None and ne.get('name','') == '図 8':
                drawing.remove(anchor)
    n = len(images_list)
    if n == 0:
        return (etree.tostring(drawing, xml_declaration=True, encoding='UTF-8', standalone=True),
                etree.tostring(rels,    xml_declaration=True, encoding='UTF-8', standalone=True))
    col_width = 37 // n
    for idx, (img_data,) in enumerate(images_list):
        rid = f'rId{20+idx}'; img_id = 20+idx
        fname = f'{media_prefix}_{idx}.png'
        col_from = idx * col_width; col_to = col_from + col_width - 1
        rel = etree.SubElement(rels, f'{{{NS_PKG}}}Relationship')
        rel.set('Id', rid)
        rel.set('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        rel.set('Target', f'../media/{fname}')
        drawing.append(make_pic_anchor(rid, img_id, f'design_{idx}', col_from, 46, col_to, 57))
        all_files[f'xl/media/{fname}'] = img_data
    return (etree.tostring(drawing, xml_declaration=True, encoding='UTF-8', standalone=True),
            etree.tostring(rels,    xml_declaration=True, encoding='UTF-8', standalone=True))

def fill_sheet(ws, d):
    tomorrow = datetime.now() + timedelta(days=1)
    ws['AD4'] = tomorrow.year; ws['AH4'] = tomorrow.month; ws['AK4'] = tomorrow.day
    ws['F11'] = d['nouki']
    ws['A22'].value = d['anken_stripped']
    all_items = [
        {'name':'本納品','qty':d['honnohin'],'unit':'個','price':d['tanka']},
        {'name':'予備',  'qty':d['yubi'],    'unit':'個','price':f"=Y{SLOTS[0]}"},
    ]
    for bi in d['betto_items']:
        all_items.append({'name':bi['name'],'qty':bi['qty'],'unit':bi['unit'],'price':bi['price']})
    for i, slot_row in enumerate(SLOTS):
        if i < len(all_items):
            item = all_items[i]
            ws.cell(row=slot_row, column=1).value  = item['name']
            ws.cell(row=slot_row, column=16).value = item['qty']
            ws.cell(row=slot_row, column=22).value = item['unit']
            ws.cell(row=slot_row, column=25).value = f"=Y{SLOTS[0]}" if i==1 else item['price']
        else:
            ws.cell(row=slot_row, column=1).value  = None
            ws.cell(row=slot_row, column=16).value = None
            ws.cell(row=slot_row, column=22).value = None
            ws.cell(row=slot_row, column=25).value = None
    specs = [
        (46, f'■種別：{d["sku"]}種 '),
        (48, f'■{d["size_label"]}：{d["size_str"]}'),
        (50, '■材質：アクリル'),
        (52, f'■印刷：{d["insatsu"]}　 '),
        (54, '■個装：OPP'),
        (56, '■備考：以上の内容で進行いたします。'),
    ]
    for row, text in specs:
        cell = ws.cell(row=row, column=2)
        cell.value = text
        cell.font  = Font(name='游ゴシック', size=11)
        cell.alignment = Alignment(vertical='center')

def generate(xlsx_paths, template_path, output_path):
    tomorrow = datetime.now() + timedelta(days=1)
    shutil.copy(template_path, output_path)
    wb_out = load_workbook(output_path)
    with zipfile.ZipFile(template_path) as zt:
        tmpl_drawing = zt.read('xl/drawings/drawing1.xml')
        tmpl_rels    = zt.read('xl/drawings/_rels/drawing1.xml.rels')
    datasets = []
    for xlsx_path in xlsx_paths:
        d = read_estimate(xlsx_path)
        raw = get_images_from_sheet(xlsx_path, '種別')
        datasets.append({'d': d, 'images': raw, 'path': xlsx_path})
    for idx, ds in enumerate(datasets):
        sheet_name = ds['d']['anken'][:31]
        if idx == 0:
            ws = wb_out['例']; ws.title = sheet_name
        else:
            ws = wb_out.copy_worksheet(wb_out.worksheets[0]); ws.title = sheet_name
        fill_sheet(ws, ds['d'])
    wb_out.save(output_path)
    with zipfile.ZipFile(output_path,'r') as zin:
        all_files = {name: zin.read(name) for name in zin.namelist()}
    with zipfile.ZipFile(template_path) as zt:
        for f in zt.namelist():
            if f.startswith('xl/media/'): all_files[f] = zt.read(f)
    for idx, ds in enumerate(datasets):
        drw_xml, drw_rels = build_drawing(tmpl_drawing, tmpl_rels,
                                          ds['images'], all_files, f'design_est{idx}')
        all_files[f'xl/drawings/drawing{idx+1}.xml'] = drw_xml
        all_files[f'xl/drawings/_rels/drawing{idx+1}.xml.rels'] = drw_rels
        s = all_files[f'xl/worksheets/sheet{idx+1}.xml'].decode('utf-8')
        if '<drawing' not in s:
            tag = '<drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            s = s.replace('</worksheet>', f'{tag}</worksheet>')
            all_files[f'xl/worksheets/sheet{idx+1}.xml'] = s.encode('utf-8')
        if idx > 0:
            all_files[f'xl/worksheets/_rels/sheet{idx+1}.xml.rels'] = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/drawing{idx+1}.xml" Id="rId1"/>
</Relationships>'''.encode('utf-8')
            ct = etree.fromstring(all_files['[Content_Types].xml'])
            if not any(f'drawing{idx+1}' in (el.get('PartName') or '') for el in ct):
                ov = etree.SubElement(ct, f'{{{NS_CT}}}Override')
                ov.set('PartName', f'/xl/drawings/drawing{idx+1}.xml')
                ov.set('ContentType','application/vnd.openxmlformats-officedocument.drawing+xml')
            all_files['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)
            wb_xml = etree.fromstring(all_files['xl/workbook.xml'])
            dns = wb_xml.find(f'{{{NS_S}}}definedNames')
            if dns is None: dns = etree.SubElement(wb_xml, f'{{{NS_S}}}definedNames')
            if not any(dn.get('localSheetId') == str(idx) for dn in dns):
                dn = etree.SubElement(dns, f'{{{NS_S}}}definedName')
                dn.set('name','_xlnm.Print_Area'); dn.set('localSheetId', str(idx))
                dn.text = f"'{ds['d']['anken'][:31]}'!$A$1:$AM$60"
            all_files['xl/workbook.xml'] = etree.tostring(wb_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(output_path,'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    return output_path

# ============================================================
# UI
# ============================================================
st.markdown("""
<div class="step-card">
  <span class="step-num">1</span><span class="step-title">見積算出表をアップロード（複数OK）</span>
</div>
""", unsafe_allow_html=True)

uploaded_xlsb = st.file_uploader(
    "見積算出表 .xlsb をドラッグ＆ドロップ",
    type=['xlsb'],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

st.markdown("""
<div class="step-card">
  <span class="step-num">2</span><span class="step-title">テンプレートをアップロード</span>
</div>
""", unsafe_allow_html=True)

uploaded_template = st.file_uploader(
    "最終見積書テンプレート .xlsx",
    type=['xlsx'],
    label_visibility="collapsed",
    key="template"
)

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_xlsb and uploaded_template:
    names = [f.name for f in uploaded_xlsb]
    st.markdown(f"""
    <div class="warning-box">
    📋 <b>{len(uploaded_xlsb)}件</b> の見積算出表を受け取りました：{"　".join(names)}
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("📄 最終見積書を生成する", type="primary", use_container_width=True):
        with st.spinner("生成中... しばらくお待ちください"):
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    # xlsb → xlsx 変換（pyxlsb使用）
                    xlsx_paths = []
                    progress = st.progress(0)
                    for i, uf in enumerate(uploaded_xlsb):
                        import pyxlsb, io as _io
                        from openpyxl import Workbook as _WB
                        xlsb_bytes = uf.read()
                        xlsx_path = os.path.join(tmpdir, uf.name.replace('.xlsb', '.xlsx'))
                        with pyxlsb.open_workbook(_io.BytesIO(xlsb_bytes)) as wb_xlsb:
                            wb_new = _WB()
                            wb_new.remove(wb_new.active)
                             for si, sname in enumerate(wb_xlsb.sheets):
                                  title = sname.strip() if sname and sname.strip() else f'Sheet{si+1}'
                                  ws_new = wb_new.create_sheet(title=title)
                             with wb_xlsb.get_sheet(sname) as sheet:
                                    for row in sheet.rows():
                                        for cell in row:
                                            if cell.v is not None and cell.r >= 1 and cell.c >= 1:
                                                ws_new.cell(row=cell.r, column=cell.c, value=cell.v)
                            wb_new.save(xlsx_path)
                        xlsx_paths.append(xlsx_path)
                        progress.progress((i+1)/len(uploaded_xlsb))

                    # テンプレート保存
                    template_path = os.path.join(tmpdir, '_template.xlsx')
                    with open(template_path, 'wb') as out:
                        out.write(uploaded_template.read())

                    # 生成
                    today = datetime.now().strftime('%Y%m%d')
                    output_path = os.path.join(tmpdir, f'_最終見積書_カードラボ様_{today}.xlsx')
                    generate(xlsx_paths, template_path, output_path)

                    with open(output_path, 'rb') as f:
                        output_bytes = f.read()

                    st.markdown(f"""
                    <div class="success-box">
                        ✅ <b>生成完了！</b><br>
                        <small style="color:#276749">{len(xlsx_paths)}件の案件を1ファイルにまとめました</small>
                    </div>
                    """, unsafe_allow_html=True)

                    fname = f'_最終見積書_カードラボ様_{today}.xlsx'
                    st.download_button(
                        label=f"⬇️ {fname} をダウンロード",
                        data=output_bytes,
                        file_name=fname,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"エラーが発生しました：{str(e)}")

else:
    st.markdown("""
    <div style="text-align:center; color:#a0aec0; padding: 32px 0; font-size:0.9rem;">
        ↑ 見積算出表とテンプレートをアップロードしてください
    </div>
    """, unsafe_allow_html=True)

# フッター
st.markdown("---")
st.markdown("<p style='text-align:center;color:#a0aec0;font-size:0.8rem;'>株式会社アイナック 事業推進部</p>",
            unsafe_allow_html=True)
