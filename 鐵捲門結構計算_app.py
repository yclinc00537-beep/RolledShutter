import streamlit as st
import math
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ==========================================
# 1. DATABASE & SPECIFICATIONS (ENGINEERING DATA)
# ==========================================
SHUTTER_CATEGORIES = ["防颱型", "一般型", "遮煙防火等級 (60A+S)"]
SLAT_TYPES = {
    "100防颱型": {"t_min": 1.5, "t_max": 2.0, "t_step": 0.1, "I_base": 21.5, "Z_base": 11.0, "pitch": 100, "desc": "防颱主流。深帽高凹凸幾何結構，高剛性。"},
    "100重防颱型": {"t_min": 2.0, "t_max": 3.0, "t_step": 0.2, "I_base": 30.5, "Z_base": 16.2, "pitch": 100, "desc": "超大跨距、沿海重負壓區或高樓層專用。"},
    "60型": {"t_min": 1.0, "t_max": 1.2, "t_step": 0.1, "I_base": 2.0, "Z_base": 1.5, "pitch": 60, "desc": "輕型門片，不建議用於大跨距防颱。"},
    "75型": {"t_min": 1.0, "t_max": 1.5, "t_step": 0.1, "I_base": 6.5, "Z_base": 4.5, "pitch": 75, "desc": "商業與一般車庫標準規格。"}
}

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# ==========================================
# 2. STREAMLIT UI & INPUTS
# ==========================================
st.set_page_config(page_title="大跨距捲門結構應力檢核 App", layout="wide")
st.title("🗜️ 大跨距捲門結構應力、防颱中柱與動力優化計算系統")
st.markdown("針對最大 **10m x 10m** 超大型工業捲門進行精密耐風變形量與可拆卸中柱鋼材選型檢核。")

col_in, col_out = st.columns([1, 1.2])

with col_in:
    st.header("📋 設計參數配置")
    
    category = st.selectbox("1. 選擇捲門級別類型", SHUTTER_CATEGORIES)
    slat_type = st.selectbox("2. 選擇門片幾何規格", list(SLAT_TYPES.keys()))
    s_info = SLAT_TYPES[slat_type]
    st.caption(f"💡 **規格備註：** {s_info['desc']}")
    
    thick = st.slider("3. 門片設計厚度 t (mm)", min_value=s_info["t_min"], max_value=s_info["t_max"], step=s_info["t_step"], value=s_info["t_max"])
    
    st.divider()
    st.subheader("📐 幾何維度與風壓 (支援 10m 大跨距)")
    width = st.slider("4. 捲門總淨寬度 L (m)", 1.5, 12.0, 12.0, 0.5)
    height = st.slider("5. 捲門總淨高度 H (m)", 2.0, 10.0, 10.0, 0.5)
    wind_load = st.slider("6. 設計風壓負載 W (kg/m²)", 100, 450, 280, 10)
    
    st.divider()
    st.subheader("🛡️ 防颱防護配置優化")
    
    # 智慧中柱檢討開關
    has_middle_pillar = st.checkbox("💡 啟用可拆卸式防颱中柱 (颱風警報時裝上)", value=True)
    if has_middle_pillar:
        st.success("🟢 已啟用防颱中柱：計算模型中，門片的抗彎計算跨距將直接減半 (L_eff = L/2)！")
        calc_width = width / 2
    else:
        st.warning("⚠️ 警報：10m 跨距若不設中柱，門片極易產生塑性破壞，請注意右側檢核報告。")
        calc_width = width

    # 防颱勾由高度自動計算
    hook_density = st.radio("防颱勾配置密度", ["每片皆裝 (高強防颱型)", "每跳一片裝設 (標準防颱型)"])
    pitch_m = s_info["pitch"] / 1000  # 門片高度(米)
    total_slats = math.ceil(height / pitch_m)
    if hook_density == "每片皆裝 (高強防颱型)":
        hook_count_side = total_slats
    else:
        hook_count_side = math.ceil(total_slats / 2)
        
    st.info(f"⚓ **智慧防颱勾計算法：** 依據門高 {height}m 與門片節距，自動算出單側所需防颱勾數量為 **{hook_count_side}** 個。")

    st.divider()
    st.subheader("⚙️ 傳動心軸與基本配置")
    shaft_dia = st.selectbox("捲軸無縫鋼管外徑 (吋)", [4, 5, 6, 8, 10, 12], index=4)
    shaft_thick = st.slider("捲軸鋼管壁厚 (mm)", 3.0, 12.0, 8.0, 0.5)

    # 基礎力學參數設定
    t_ratio = thick / s_info["t_min"]
    I_eff = s_info["I_base"] * (t_ratio ** 1.1)
    Z_eff = s_info["Z_base"] * (t_ratio ** 1.05)
    E_steel = 2.04e6  # kg/cm²
    Fy_steel = 2400   # kg/cm²
    Fa_bend = 0.6 * Fy_steel

# ==========================================
# 3. CORE CALCULATIONS
# ==========================================
total_area = width * height
total_wind_force = wind_load * total_area

# 門片彎曲（套用有效計算跨距）
w_line = wind_load * 1.0  # kg/m
w_line_cm = w_line / 100  # kg/cm
L_calc_cm = calc_width * 100

M_max = (w_line_cm * (L_calc_cm ** 2)) / 8
slat_stress = M_max / Z_eff
slat_deflection = (5 * w_line_cm * (L_calc_cm ** 4)) / (384 * E_steel * I_eff)
deflection_limit = L_calc_cm / 100

# 防颱勾拉力分析
total_hooks = hook_count_side * (4 if has_middle_pillar else 2) # 有中柱時多了中柱兩側導軌的勾數
force_per_hook = total_wind_force / total_hooks if total_hooks > 0 else 0

# 中柱本體受力與鋼材選型計算 (若有啟用)
if has_middle_pillar:
    # 中柱承受左右各半範圍的風壓，即均布載荷 w_pillar = W * (W_total / 2)
    w_pillar_m = wind_load * (width / 2)  # kg/m
    w_pillar_cm = w_pillar_m / 100
    H_pillar_cm = height * 100
    
    M_pillar = (w_pillar_cm * (H_pillar_cm ** 2)) / 8 # 中柱最大彎矩
    # 檢核是否需要重型 H 型鋼 (以常見 H250x250 或 H300x300 作基準評估)
    # 假設容許應力 1440 kg/cm²，求最小需要之 Z 值
    Z_required_pillar = M_pillar / (0.6 * Fy_steel)
    
    if Z_required_pillar <= 400:
        pillar_spec = "H 175 x 175 x 7.5 x 11 (輕型工字鋼)"
    elif Z_required_pillar <= 700:
        pillar_spec = "H 250 x 250 x 9 x 14 (標準結構型)"
    elif Z_required_pillar <= 1300:
        pillar_spec = "H 300 x 300 x 10 x 15 (重型防颱鋼柱)"
    else:
        pillar_spec = "H 400 x 400 級或格構式加勁雙拼型雙柱"

# 捲軸承重與馬達計算
slat_weight_map = {"60型": 12, "75型": 16, "100防颱型": 25, "100重防颱型": 35}
total_slat_weight = total_area * slat_weight_map[slat_type]
w_shaft_cm = (total_slat_weight / width) / 100

r_out = (shaft_dia * 2.54) / 2
r_in = r_out - (shaft_thick / 10)
I_shaft = (math.pi / 4) * (r_out**4 - r_in**4)
shaft_deflection = (5 * w_shaft_cm * ((width * 100) ** 4)) / (384 * E_steel * I_shaft)
shaft_limit = (width * 100) / 200

# 馬達噸位修正
if total_slat_weight <= 400:
    motor_hp, motor_kg = "1/3 HP", "400 kg 級"
elif total_slat_weight <= 600:
    motor_hp, motor_kg = "1/2 HP", "600 kg 級"
elif total_slat_weight <= 1200:
    motor_hp, motor_kg = "1 HP (三相)", "1200 kg 級"
elif total_slat_weight <= 2200:
    motor_hp, motor_kg = "2 HP (三相重載)", "2000 ~ 2500 kg 級"
else:
    motor_hp, motor_kg = "3 HP ~ 5 HP 超重級開閉機", "3500 kg 級超高規格"

# ==========================================
# 4. STREAMLIT OUTPUTS VIEW
# ==========================================
with col_out:
    st.header("📊 巨型結構檢核報告與優化建議")
    
    c1, c2 = st.columns(2)
    c1.metric("🌪️ 總風壓推力面外負載", f"{total_wind_force:,.0f} kg", f"總受風面高達 {total_area:.1f} ㎡")
    c2.metric("⚖️ 門扇懸吊總重量", f"{total_slat_weight:,.0f} kg", f"門片基礎自重約 {slat_weight_map[slat_type]} kg/㎡")
    
    st.divider()
    
    # 門片狀態
    st.subheader("1. 門片企口抗彎與撓度安全檢核")
    stress_pass = slat_stress <= Fa_bend
    defl_pass = slat_deflection <= deflection_limit
    
    cc1, cc2 = st.columns(2)
    cc1.metric("門片最大應力狀態", f"{slat_stress:.1f} / {Fa_bend:.0f} kg/cm²", 
              delta=f"安全" if stress_pass else f"超限破壞", delta_color="normal" if stress_pass else "inverse")
    cc2.metric("門片中央變形撓度", f"{slat_deflection:.2f} / {deflection_limit:.2f} cm", 
              delta=f"合格" if defl_pass else f"嚴重脫軌", delta_color="normal" if defl_pass else "inverse")
    
    if stress_pass and defl_pass:
        st.success("✅ 經由可拆卸式中柱有效削減跨距，門片結構撓度與應力通過安全檢核！")
    else:
        st.error("❌ 警告：此幾何斷面在當前跨距下無法承受風壓！門片會折斷或從導軌被完全拉出。")

    # 中柱專屬檢核結果
    if has_middle_pillar:
        st.divider()
        st.subheader("🛡️ 2. 可拆卸式防颱中柱 (Middle Pillar) 結構檢討")
        st.info(f"📐 中柱所需之彈性截面抗彎模數 $Z_{{req}}$ 需大於：**{Z_required_pillar:.1f} cm³**")
        st.warning(f"🏗️ **結構技師建議鋼材選型：** 現場防颱中柱鋼軌必須採用 **{pillar_spec}** 規格之熱軋工字鋼或 H 型鋼，並配備上下固定承座螺栓組，方能抵擋強風。")

    st.divider()
    st.subheader("⚙️ 3. 心軸剛性與動力系統配置")
    
    if shaft_deflection <= shaft_limit:
        st.success(f"捲軸無縫鋼管自重撓度：{shaft_deflection:.3f} cm ≦ 規範 {shaft_limit:.2f} cm (L/200) 🟢 剛性足夠")
    else:
        st.error(f"捲軸無縫鋼管自重撓度：{shaft_deflection:.3f} cm ＞ 規範 {shaft_limit:.2f} cm (L/200) 🔴 管徑過細或壁厚不足，捲轴下垂將導致馬達卡死破壞！")
        
    st.info(f"⚓ **防颱勾點狀拉拔負載：** 單個專利防颱勾與中柱導軌接合處需承受 **{force_per_hook:.1f} kg** 的剪切抽拉力。")
    st.warning(f"⚡ **建議動力配置：** 選用 **{motor_hp}** （提升荷重等級：**{motor_kg}**），並應設置機械式防過捲極限二次安全保護裝置。")

# ==========================================
# 5. WORD REPORT EXPORT LOGIC (動態優化版)
# ==========================================
def generate_word_report():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = '新細明體'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"{width}m × {height}m 大跨距捲門結構強度與防颱中柱計算書")
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)
    
    doc.add_paragraph(f"報告生成時間：2026年10月02日  |  結構檢核標準：CNS 4212 / 建築物耐風設計規範")
    
    h1 = doc.add_heading(level=1)
    h1.add_run("一、 設計基本參數與邊界條件").font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)
    p = doc.add_paragraph()
    p.add_run(f"• 捲門用途等級分級：{category}\n")
    p.add_run(f"• 選用門片型式：{slat_type} (厚度 {thick} mm)\n")
    p.add_run(f"• 總幾何維度大小：寬度 L = {width} m, 高度 H = {height} m (總受風面積 {total_area:.1f} ㎡)\n")
    p.add_run(f"• 可拆卸式防颱中柱配置：{'🟢 已配置 (颱風警報時裝上，有效計算跨距折半為 ' + str(calc_width) + ' m)' if has_middle_pillar else '❌ 未配置 (採全跨距計算)'}\n")
    p.add_run(f"• 智慧防颱勾計算法：依據高度與節距算出單側配置數量為 {hook_count_side} 個\n")
    p.add_run(f"• 設計風壓負載強度 (W)：{wind_load} kg/m²\n")
    p.add_run(f"• 捲軸無縫鋼管：外徑 {shaft_dia} 吋 , 壁厚 {shaft_thick} mm")

    h2 = doc.add_heading(level=1)
    h2.add_run("二、 結構應力與物理變形形變量檢核結果").font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)
    
    table = doc.add_table(rows=6, cols=5)
    table.style = 'Table Grid'
    headers = ["檢核項目", "設計計算值", "規範上限值", "安全判定", "結構優化與設計指引建議"]
    
    hdr_cells = table.rows[0].cells
    for i, title_text in enumerate(headers):
        hdr_cells[i].text = title_text
        set_cell_background(hdr_cells[i], "0F2C59")
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        set_cell_margins(hdr_cells[i])
        
    # 動態調整表格內顯示的字樣，明確標示是否有經過中柱修正
    slat_stress_label = f"{slat_stress:.1f} kg/cm²\n(已依中柱減半跨距 {calc_width}m 修正)" if has_middle_pillar else f"{slat_stress:.1f} kg/cm²\n(全跨距 {width}m 計算)"
    slat_defl_label = f"{slat_deflection:.2f} cm\n(已依中柱形變大幅修正)" if has_middle_pillar else f"{slat_deflection:.2f} cm\n(全跨距未修正變形)"
    defl_limit_label = f"{deflection_limit:.2f} cm\n(L_eff/100)" if has_middle_pillar else f"{deflection_limit:.2f} cm\n(L/100)"

    row_data = [
        ["門片彎曲應力", slat_stress_label, f"{Fa_bend:.0f} kg/cm²", "🟢 合格" if stress_pass else "🔴 超限破壞", "材料未降伏，中柱分擔效果顯著。" if stress_pass else "應力超限！強風下門片有折斷危險，必須增厚或加中柱。"],
        ["門片中央撓度", slat_defl_label, defl_limit_label, "🟢 合格" if defl_pass else "🔴 嚴重脫軌", f"符合限制。防颱導軌插入淨深需大於 {slat_deflection*1.5:.1f} cm。" if defl_pass else "形變過大！門片將在颱風中被拉出軌道。"],
        ["單點防颱勾拉力", f"{force_per_hook:.1f} kg", "視錨固件強度", "⚠️ 需對標", "請確保專利T型勾與側軌之抗拉拔極限剪力大於此計算值。"],
        ["心軸結構撓度", f"{shaft_deflection:.3f} cm", f"{shaft_limit:.2f} cm", "🟢 合格" if shaft_deflection <= shaft_limit else "🔴 剛性不足", "心軸垂直自重下垂量小於 L/200，運轉順暢。" if shaft_deflection <= shaft_limit else "軸心嚴重下垂，將導致捲筒卡死，請加大管徑。"],
        ["動力馬達配置", f"總重約 {total_slat_weight:.0f} kg", f"建議 {motor_hp}", "⚡ 推薦選型", f"建議採用 {motor_kg} 電動開閉機組。"]
    ]
    
    for row_idx, data in enumerate(row_data):
        row_cells = table.rows[row_idx + 1].cells
        for col_idx, text in enumerate(data):
            row_cells[col_idx].text = text
            set_cell_margins(row_cells[col_idx])
            if col_idx == 3:
                if "🟢" in text:
                    row_cells[col_idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)
                elif "🔴" in text:
                    row_cells[col_idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xC6, 0x28, 0x28)

    doc.add_paragraph()
    h3 = doc.add_heading(level=1)
    h3.add_run("三、 國內外主流規範管理要點對標評比表").font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)
    
    table_compare = doc.add_table(rows=5, cols=4)
    table_compare.style = 'Table Grid'
    comp_headers = ["規範體系", "最大容許撓度限制", "風壓與環境認證要點", "巨型尺寸結構特殊要求"]
    
    hdr_comp_cells = table_compare.rows[0].cells
    for i, title_text in enumerate(comp_headers):
        hdr_comp_cells[i].text = title_text
        set_cell_background(hdr_comp_cells[i], "4A5568")
        hdr_comp_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_margins(hdr_comp_cells[i])
        
    comp_rows = [
        ["臺灣 (CNS 4212 / 建築物耐風規範)", "L/100 至 L/150", "依現地公式計算靜態設計風壓，大跨距需附公正第三方風洞測試報告。", "跨距超過 5.6m 時，實務施工圖說多強制要求加設抗風活動中柱。"],
        ["歐盟 (EN 12424 / EN 13241-1)", "不脫軌與結構安全為唯一判定原則", "劃分為 Class 0 至 Class 5+ 級別。巨型尺寸需執行嚴格的有限元素有效斷面法(FEA)分析。", "強制要求底部需配備阻礙反轉防夾感應安全裝置(Safety Edge)。"],
        ["美國 (ASTM E330 / DASMA 108)", "L/120 (通用標準)", "特定颶風區(HVHZ)要求進行高達數千次之『正負交變氣壓疲勞推拉測試』。", "大型工業門片與防颱勾(Windlocks)的扣件剪力與點狀拉拔應力需全面簽證。"],
        ["日本 (JIS A 4710 / 建築基準法)", "L/100 至 L/150", "極度聚焦於『災後功能回復性』，暴風過後門片殘餘變形量不得導致卡死。", "超大型機棚大門需獨立進行結構力學局部座屈與地錨剪力檢核。"]
    ]
    
    for row_idx, data in enumerate(comp_rows):
        row_cells = table_compare.rows[row_idx + 1].cells
        for col_idx, text in enumerate(data):
            row_cells[col_idx].text = text
            set_cell_margins(row_cells[col_idx])
            
    doc.add_paragraph()
    h4 = doc.add_heading(level=2)
    h4.add_run("四、 特殊大跨距巨型結構工程設計備忘錄").font.color.rgb = RGBColor(0x2B, 0x6C, 0xBF)
    p_note = doc.add_paragraph()
    
    # 修正：全面動態化尺寸與面積標示
    p_note.add_run(f"⚠️ 【{width}m × {height}m 大跨距結構特定工程指引】\n").font.bold = True
    p_note.add_run(f"1. 本案捲門屬於 {total_area:.1f} ㎡ 之大面積受風面。{'當啟用可拆卸式防颱中柱時，中柱本體在颱風時需承受中段高達數噸之集中風剪力。計算得出中柱鋼材斷面模數需求，現場強制建議配置不低於 【' + pillar_spec + '】 之結構熱軋型鋼軌道柱。' if has_middle_pillar else '🚨警告：本案未配置防颱中柱！在 ' + str(width) + 'm 全跨距下，風壓將完全由兩側牆體導軌與門片企口承受，請務必再度確認右側檢核報告之應力與撓度是否超限。'}\n")
    p_note.add_run(f"2. 中柱與地面之『防颱地錨承座（Floor Socket）』及頂端天車結構之固定梢，必須能夠承受高達數千公斤之剪力推力，地錨埋設深度與混凝土強度(fc')需一併進行結構簽證檢核。\n")
    p_note.add_run(f"3. 智慧防颱勾裝設方式：高度方向單側已優化由系統依實際門高自動精算配置 {hook_count_side} 組防颱勾。施工時兩側及中柱導軌之防颱反扣條必須採用一體成型且厚度 4.5mm 以上之鋼板，嚴禁現場以薄板電焊點焊組裝，以免受颱風交變正負壓抽吸時發生局部剪切撕裂而脫軌。")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- DOWNLOAD BUTTON ---
st.divider()
doc_bytes = generate_word_report()
st.download_button(
    label="📥 下載大跨距捲門專屬結構計算書 & 規範評比 (Word 檔)",
    data=doc_bytes,
    file_name=f"{width}mx{height}m_捲門結構中柱與動力計算報告.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)