import streamlit as st
import requests
import math
import io
from itertools import combinations
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Distâncias entre Cidades",
    page_icon="🗺️",
    layout="wide"
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.header-box {
    background: linear-gradient(135deg, #0f1117 60%, #1a2744);
    color: #f5f0e8;
    padding: 28px 36px;
    border-radius: 8px;
    margin-bottom: 24px;
}
.header-tag {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #c84b2f;
    margin-bottom: 8px;
}
.header-title {
    font-size: 2rem;
    font-weight: 300;
    line-height: 1.2;
    margin: 0;
}
.header-sub {
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem;
    color: #8a8070;
    margin-top: 6px;
    letter-spacing: 0.5px;
}

/* Chip de cidade */
.city-chip {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #eff6ff;
    border: 1.5px solid #bfdbfe;
    border-radius: 20px;
    padding: 4px 12px;
    font-size: 0.82rem;
    color: #1e40af;
    font-weight: 600;
    margin: 3px;
}
.chip-coords {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    color: #64748b;
}

/* Tabela */
.dist-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 8px;
}
.dist-table thead th {
    background: #0f1117;
    color: #f5f0e8;
    padding: 10px 12px;
    text-align: center;
    font-weight: 600;
    white-space: nowrap;
    font-size: 0.75rem;
}
.dist-table thead th:first-child {
    background: #1a1d26;
    text-align: left;
    min-width: 140px;
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 1px;
    text-transform: uppercase;
}
.dist-table tbody tr:nth-child(even) { background: #f8fafc; }
.dist-table tbody tr:hover td { background: #f0f9ff !important; }
.dist-table tbody td {
    padding: 9px 12px;
    text-align: center;
    border-bottom: 1px solid #e2e8f0;
    border-right: 1px solid #f1f5f9;
    vertical-align: middle;
}
.dist-table tbody td:first-child {
    font-weight: 600;
    text-align: left;
    color: #0f1117;
    white-space: nowrap;
    background: #fafaf8;
    border-right: 2px solid #cbd5e1;
}
.cell-self { background: #f1f5f9 !important; color: #94a3b8; font-size: 1.1rem; }
.val-line { color: #1e40af; font-weight: 700; font-size: 0.84rem; }
.val-road { color: #166534; font-weight: 700; font-size: 0.84rem; }
.val-label { font-size: 0.62rem; color: #94a3b8; font-family: 'DM Mono', monospace; }
.val-nd { color: #fca5a5; font-size: 0.72rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <div class="header-tag">// análise geoespacial · Brasil</div>
    <h1 class="header-title">🗺️ Distâncias entre <em>Cidades</em></h1>
    <div class="header-sub">linha reta · por estrada · coordenadas das sedes municipais (prefeituras)</div>
</div>
""", unsafe_allow_html=True)

# ── Dados das capitais ─────────────────────────────────────────────────────────
CAPITALS = [
    {"name": "Rio Branco",     "state": "AC", "lat": -9.97499,  "lon": -67.82471, "address": "R. Benjamim Constant, 945 - Centro"},
    {"name": "Maceió",         "state": "AL", "lat": -9.66583,  "lon": -35.73528, "address": "Praça Thomaz Espíndola, s/n - Centro"},
    {"name": "Macapá",         "state": "AP", "lat":  0.03444,  "lon": -51.06639, "address": "Av. Iracema Carvão Do Nascimento, 600"},
    {"name": "Manaus",         "state": "AM", "lat": -3.10194,  "lon": -60.02500, "address": "Av. Brasil, 2971 - Compensa"},
    {"name": "Salvador",       "state": "BA", "lat": -12.97111, "lon": -38.51083, "address": "Praça Municipal, s/n - Centro Histórico"},
    {"name": "Fortaleza",      "state": "CE", "lat": -3.72389,  "lon": -38.54306, "address": "Av. Desembargador Moreira, 310 - Meireles"},
    {"name": "Brasília",       "state": "DF", "lat": -15.78361, "lon": -47.89833, "address": "Palácio do Buriti - Setor de Áreas Isoladas Sul"},
    {"name": "Vitória",        "state": "ES", "lat": -20.31944, "lon": -40.33778, "address": "Av. Marechal Mascarenhas de Moraes, 1927"},
    {"name": "Goiânia",        "state": "GO", "lat": -16.67861, "lon": -49.25389, "address": "Av. do Cerrado, 999 - Park Lozandes"},
    {"name": "São Luís",       "state": "MA", "lat": -2.52972,  "lon": -44.30278, "address": "Rua Afonso Pena, s/n - Centro"},
    {"name": "Cuiabá",         "state": "MT", "lat": -15.59611, "lon": -56.09667, "address": "Av. General Mello, s/n - Porto"},
    {"name": "Campo Grande",   "state": "MS", "lat": -20.44278, "lon": -54.64611, "address": "Av. Afonso Pena, 3297 - Centro"},
    {"name": "Belo Horizonte", "state": "MG", "lat": -19.91722, "lon": -43.93444, "address": "Av. Afonso Pena, 1212 - Centro"},
    {"name": "Belém",          "state": "PA", "lat": -1.45583,  "lon": -48.50444, "address": "Praça Felipe Patroni, s/n - Cidade Velha"},
    {"name": "João Pessoa",    "state": "PB", "lat": -7.11528,  "lon": -34.86278, "address": "Praça João Pessoa, s/n - Centro"},
    {"name": "Curitiba",       "state": "PR", "lat": -25.42944, "lon": -49.27167, "address": "Av. Cândido de Abreu, 817 - Centro Cívico"},
    {"name": "Recife",         "state": "PE", "lat": -8.05361,  "lon": -34.88111, "address": "Av. Cais do Apolo, 925 - Recife Antigo"},
    {"name": "Teresina",       "state": "PI", "lat": -5.08917,  "lon": -42.80194, "address": "R. Areolino de Abreu, 900 - Centro"},
    {"name": "Rio de Janeiro", "state": "RJ", "lat": -22.90278, "lon": -43.17444, "address": "R. Afonso Cavalcanti, 455 - Cidade Nova"},
    {"name": "Natal",          "state": "RN", "lat": -5.79500,  "lon": -35.21139, "address": "Av. Deodoro da Fonseca, 384 - Cidade Alta"},
    {"name": "Porto Velho",    "state": "RO", "lat": -8.76194,  "lon": -63.90389, "address": "Av. 7 de Setembro, 237 - Centro"},
    {"name": "Boa Vista",      "state": "RR", "lat":  2.81972,  "lon": -60.67333, "address": "Rua Coronel Pinto, 241 - Centro"},
    {"name": "Porto Alegre",   "state": "RS", "lat": -30.03444, "lon": -51.21750, "address": "Av. Loureiro da Silva, 255 - Centro Histórico"},
    {"name": "Florianópolis",  "state": "SC", "lat": -27.59500, "lon": -48.54861, "address": "R. Timóteo Pereira da Costa, 10 - Centro"},
    {"name": "Aracaju",        "state": "SE", "lat": -10.91111, "lon": -37.07167, "address": "Av. Dr. Carlos Firpo, s/n - Capucho"},
    {"name": "São Paulo",      "state": "SP", "lat": -23.55028, "lon": -46.63361, "address": "Viaduto do Chá, 15 - Centro"},
    {"name": "Palmas",         "state": "TO", "lat": -10.18611, "lon": -48.33361, "address": "Quadra 502 Sul, Av. NS-02, s/n"},
]

# ── Funções ────────────────────────────────────────────────────────────────────
def haversine(la1, lo1, la2, lo2):
    R = 6371
    dL = math.radians(la2 - la1)
    dO = math.radians(lo2 - lo1)
    a = math.sin(dL/2)**2 + math.cos(math.radians(la1)) * math.cos(math.radians(la2)) * math.sin(dO/2)**2
    return round(R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a)), 1)

def get_road_distance(c1, c2, ors_key):
    try:
        resp = requests.post(
            "https://api.openrouteservice.org/v2/directions/driving-car",
            headers={"Authorization": ors_key, "Content-Type": "application/json"},
            json={"coordinates": [[c1["lon"], c1["lat"]], [c2["lon"], c2["lat"]]]},
            timeout=15
        )
        if resp.status_code == 200:
            data = resp.json()
            m = (data.get("routes", [{}])[0].get("summary", {}).get("distance") or
                 data["features"][0]["properties"]["segments"][0]["distance"])
            return round(m / 1000, 1)
        else:
            return None
    except Exception:
        return None

def build_excel(cities, matrix):
    wb = Workbook()

    # Estilos
    hdr_font   = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    hdr_fill_d = PatternFill("solid", fgColor="0F1117")
    hdr_fill_l = PatternFill("solid", fgColor="1A1D26")
    cell_font  = Font(name="Arial", size=9)
    bold_font  = Font(name="Arial", bold=True, size=9)
    self_fill  = PatternFill("solid", fgColor="F1F5F9")
    line_font  = Font(name="Arial", bold=True, color="1E40AF", size=9)
    road_font  = Font(name="Arial", bold=True, color="166534", size=9)
    alt_fill   = PatternFill("solid", fgColor="F8FAFC")
    thin = Side(style="thin", color="E2E8F0")
    border = Border(bottom=thin, right=thin)

    def make_matrix_sheet(wb, title, value_fn):
        ws = wb.create_sheet(title)
        # Cabeçalho
        ws.cell(1, 1, "Origem / Destino").font = Font(name="Arial", bold=True, color="FFFFFF", size=8)
        ws.cell(1, 1).fill = hdr_fill_l
        ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
        ws.column_dimensions["A"].width = 22
        for j, c in enumerate(cities, 2):
            cell = ws.cell(1, j, f"{c['name']}, {c['state']}")
            cell.font = hdr_font
            cell.fill = hdr_fill_d
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(j)].width = 16
        ws.row_dimensions[1].height = 28
        # Dados
        for i, c1 in enumerate(cities):
            row = i + 2
            name_cell = ws.cell(row, 1, f"{c1['name']}, {c1['state']}")
            name_cell.font = bold_font
            name_cell.alignment = Alignment(vertical="center")
            if i % 2 == 1:
                name_cell.fill = alt_fill
            for j, c2 in enumerate(cities, 2):
                cell = ws.cell(row, j)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border
                if c1["name"] == c2["name"]:
                    cell.value = "—"
                    cell.fill = self_fill
                    cell.font = Font(name="Arial", color="94A3B8", size=10)
                else:
                    val = value_fn(c1, c2)
                    cell.value = val
                    cell.font = Font(name="Arial", size=9)
                    if i % 2 == 1:
                        cell.fill = alt_fill
                ws.row_dimensions[row].height = 18
        return ws

    # Aba 1 — Linha reta
    wb.remove(wb.active)
    ws1 = make_matrix_sheet(wb, "Linha Reta (km)",
        lambda c1, c2: haversine(c1["lat"], c1["lon"], c2["lat"], c2["lon"]))
    for row in ws1.iter_rows(min_row=2):
        for cell in row[1:]:
            if isinstance(cell.value, float):
                cell.font = line_font
                cell.number_format = '#,##0.0'

    # Aba 2 — Por estrada
    ws2 = make_matrix_sheet(wb, "Por Estrada (km)",
        lambda c1, c2: matrix.get(f"{c1['name']}-{c2['name']}", {}).get("road"))
    for row in ws2.iter_rows(min_row=2):
        for cell in row[1:]:
            if isinstance(cell.value, float):
                cell.font = road_font
                cell.number_format = '#,##0.0'
            elif cell.value is None and cell.row > 1:
                cell.value = "N/D"
                cell.font = Font(name="Arial", color="FCA5A5", size=8)

    # Aba 3 — Comparativo completo
    ws3 = wb.create_sheet("Comparativo Completo")
    headers = ["Cidade A", "UF", "Cidade B", "UF", "Linha Reta (km)", "Por Estrada (km)", "Fator Tortuosidade"]
    widths =  [20, 5, 20, 5, 16, 16, 18]
    for j, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws3.cell(1, j, h)
        cell.font = hdr_font
        cell.fill = hdr_fill_d
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws3.column_dimensions[get_column_letter(j)].width = w
    ws3.row_dimensions[1].height = 22
    row = 2
    for c1, c2 in combinations(cities, 2):
        line = haversine(c1["lat"], c1["lon"], c2["lat"], c2["lon"])
        road = matrix.get(f"{c1['name']}-{c2['name']}", {}).get("road")
        fator = round(road / line, 2) if road and line else None
        vals = [c1["name"], c1["state"], c2["name"], c2["state"], line, road or "N/D", fator or ""]
        fills = [None, None, None, None, "1E40AF", "166534", "6B7280"]
        for j, (v, fc) in enumerate(zip(vals, fills), 1):
            cell = ws3.cell(row, j, v)
            cell.font = Font(name="Arial", bold=(j in [1,3]), color=fc or "000000", size=9)
            cell.alignment = Alignment(horizontal="center" if j > 4 else "left", vertical="center")
            cell.border = border
            if isinstance(v, float):
                cell.number_format = '#,##0.0' if j < 7 else '0.00'
            if row % 2 == 1:
                cell.fill = alt_fill
        ws3.row_dimensions[row].height = 16
        row += 1

    # Aba 4 — Coordenadas
    ws4 = wb.create_sheet("Coordenadas das Sedes")
    h4 = ["Cidade", "UF", "Latitude (Sede)", "Longitude (Sede)", "Endereço da Sede Municipal"]
    w4 = [20, 5, 16, 16, 46]
    for j, (h, w) in enumerate(zip(h4, w4), 1):
        cell = ws4.cell(1, j, h)
        cell.font = hdr_font; cell.fill = hdr_fill_d
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws4.column_dimensions[get_column_letter(j)].width = w
    ws4.row_dimensions[1].height = 22
    for i, c in enumerate(cities, 2):
        vals = [c["name"], c["state"], c["lat"], c["lon"], c["address"]]
        for j, v in enumerate(vals, 1):
            cell = ws4.cell(i, j, v)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center")
            cell.border = border
            if isinstance(v, float):
                cell.number_format = '0.00000'
            if i % 2 == 0:
                cell.fill = alt_fill
        ws4.row_dimensions[i].height = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Session state ──────────────────────────────────────────────────────────────
if "selected" not in st.session_state:
    st.session_state.selected = [c["name"] for c in CAPITALS]
if "matrix" not in st.session_state:
    st.session_state.matrix = {}
if "calculated" not in st.session_state:
    st.session_state.calculated = False
if "calc_cities" not in st.session_state:
    st.session_state.calc_cities = []

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuração")
    ors_key = st.text_input(
        "Chave OpenRouteService",
        type="password",
        placeholder="5b3ce3597851...",
        help="Gratuita em openrouteservice.org — necessária para calcular rotas por estrada"
    )
    if ors_key:
        st.success("✓ Chave configurada")
    else:
        st.info("Sem chave: apenas linha reta ✈\n\n[Obter chave gratuita →](https://openrouteservice.org/dev/#/signup)")

    st.markdown("---")
    st.markdown("### 📍 Selecionar Cidades")

    col1, col2 = st.columns(2)
    if col1.button("✓ Todas", use_container_width=True):
        st.session_state.selected = [c["name"] for c in CAPITALS]
        st.session_state.calculated = False
        st.rerun()
    if col2.button("✗ Nenhuma", use_container_width=True):
        st.session_state.selected = []
        st.session_state.calculated = False
        st.rerun()

    new_selected = []
    for c in CAPITALS:
        checked = st.checkbox(
            f"{c['name']} ({c['state']})",
            value=c["name"] in st.session_state.selected,
            key=f"cb_{c['name']}"
        )
        if checked:
            new_selected.append(c["name"])

    if set(new_selected) != set(st.session_state.selected):
        st.session_state.selected = new_selected
        st.session_state.calculated = False

    st.markdown(f"**{len(st.session_state.selected)}** cidade(s) selecionada(s)")

# ── Main content ───────────────────────────────────────────────────────────────
selected_cities = [c for c in CAPITALS if c["name"] in st.session_state.selected]

# Chips das cidades selecionadas
if selected_cities:
    chips_html = "<div style='margin-bottom:16px;line-height:2'>"
    for c in selected_cities:
        chips_html += f"""<span class="city-chip">
            {c['name']}, {c['state']}
            <span class="chip-coords">{c['lat']:.4f}, {c['lon']:.4f}</span>
        </span>"""
    chips_html += "</div>"
    st.markdown(chips_html, unsafe_allow_html=True)
else:
    st.warning("Selecione pelo menos 2 cidades no menu lateral.")

# Botão calcular
col_btn1, col_btn2, col_btn3 = st.columns([2, 2, 4])
can_calc = len(selected_cities) >= 2

with col_btn1:
    calc_clicked = st.button(
        "⚡ Calcular Distâncias",
        disabled=not can_calc,
        use_container_width=True,
        type="primary"
    )

# ── Cálculo ────────────────────────────────────────────────────────────────────
if calc_clicked and can_calc:
    pairs = list(combinations(selected_cities, 2))
    matrix = {}
    has_ors = bool(ors_key and len(ors_key) > 10)

    if has_ors:
        # Testa a chave primeiro
        test = get_road_distance(selected_cities[0], selected_cities[1], ors_key)
        if test is None:
            st.error("❌ Chave ORS inválida ou sem créditos. Calculando apenas linha reta.")
            has_ors = False

    progress = st.progress(0, text="Iniciando cálculo…")
    status_text = st.empty()

    for idx, (c1, c2) in enumerate(pairs):
        pct = int((idx / len(pairs)) * 100)
        line = haversine(c1["lat"], c1["lon"], c2["lat"], c2["lon"])
        road = None

        if has_ors:
            status_text.markdown(f"🚗 **{c1['name']} → {c2['name']}** ({idx+1}/{len(pairs)})")
            road = get_road_distance(c1, c2, ors_key)

        key = f"{c1['name']}-{c2['name']}"
        matrix[key] = {"line": line, "road": road}
        matrix[f"{c2['name']}-{c1['name']}"] = {"line": line, "road": road}
        progress.progress(pct, text=f"Calculando… {idx+1}/{len(pairs)}")

    progress.progress(100, text="✅ Concluído!")
    status_text.empty()

    st.session_state.matrix = matrix
    st.session_state.calculated = True
    st.session_state.calc_cities = selected_cities
    st.session_state.has_ors = has_ors

# ── Resultado ──────────────────────────────────────────────────────────────────
if st.session_state.calculated and st.session_state.calc_cities:
    cities = st.session_state.calc_cities
    matrix = st.session_state.matrix
    has_ors = st.session_state.get("has_ors", False)

    if not has_ors:
        st.info("ℹ️ Distâncias por estrada não disponíveis. Configure a chave ORS para calculá-las.")

    st.markdown("---")

    # Legenda
    st.markdown("""
    <div style='display:flex;gap:20px;margin-bottom:12px;font-size:0.8rem;color:#64748b;font-family:monospace'>
        <span><span style='color:#1e40af;font-weight:700'>■</span> ✈ Linha reta (km)</span>
        <span><span style='color:#166534;font-weight:700'>■</span> 🚗 Por estrada (km)</span>
    </div>
    """, unsafe_allow_html=True)

    # Monta HTML da tabela
    tbl = "<div style='overflow-x:auto'><table class='dist-table'><thead><tr>"
    tbl += "<th>Origem / Destino</th>"
    for c in cities:
        tbl += f"<th>{c['name']}<br><span style='font-family:monospace;font-size:.6rem;opacity:.6'>{c['state']}</span></th>"
    tbl += "</tr></thead><tbody>"

    for c1 in cities:
        tbl += f"<tr><td>{c1['name']}, {c1['state']}</td>"
        for c2 in cities:
            if c1["name"] == c2["name"]:
                tbl += "<td class='cell-self'>—</td>"
            else:
                d = matrix.get(f"{c1['name']}-{c2['name']}", {})
                line = d.get("line")
                road = d.get("road")
                road_html = (f"<div><span class='val-road'>{road} km</span> <span class='val-label'>🚗</span></div>"
                             if road else "<div><span class='val-nd'>N/D</span> <span class='val-label'>🚗</span></div>")
                tbl += f"""<td>
                    <div><span class='val-line'>{line} km</span> <span class='val-label'>✈</span></div>
                    {road_html}
                </td>"""
        tbl += "</tr>"
    tbl += "</tbody></table></div>"

    st.markdown(tbl, unsafe_allow_html=True)

    # Export Excel
    st.markdown("---")
    excel_buf = build_excel(cities, matrix)
    with col_btn2:
        st.download_button(
            label="📊 Baixar Excel",
            data=excel_buf,
            file_name="distancias-cidades.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
