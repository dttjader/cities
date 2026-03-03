import streamlit as st
import requests
import math
import io
from itertools import combinations
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Distâncias entre Cidades", page_icon="🗺️", layout="wide")

# ── Índice de Serpentividade ───────────────────────────────────────────────────
# < 1.20  → Direta     (verde)
# 1.20–1.50 → Moderada (amarelo)
# > 1.50  → Sinuosa    (vermelho)
def serp_class(idx):
    if idx is None:   return None, None, None
    if idx < 1.20:    return "Direta",   "#15803d", "#dcfce7"
    if idx < 1.50:    return "Moderada", "#92400e", "#fef3c7"
    return             "Sinuosa",  "#991b1b", "#fee2e2"

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.header-box {
    background: linear-gradient(135deg, #0f1117 60%, #1a2744);
    color: #f5f0e8; padding: 28px 36px; border-radius: 8px; margin-bottom: 24px;
}
.header-tag { font-family:'DM Mono',monospace; font-size:.7rem; letter-spacing:3px;
    text-transform:uppercase; color:#c84b2f; margin-bottom:8px; }
.header-title { font-size:2rem; font-weight:300; line-height:1.2; margin:0; }
.header-sub { font-family:'DM Mono',monospace; font-size:.75rem; color:#8a8070;
    margin-top:6px; letter-spacing:.5px; }
.city-chip { display:inline-flex; align-items:center; gap:6px; background:#eff6ff;
    border:1.5px solid #bfdbfe; border-radius:20px; padding:4px 12px; font-size:.82rem;
    color:#1e40af; font-weight:600; margin:3px; }
.chip-coords { font-family:'DM Mono',monospace; font-size:.65rem; color:#64748b; }
/* Tabela principal */
.dist-table { width:100%; border-collapse:collapse; font-size:.8rem; margin-top:8px; }
.dist-table thead th { background:#0f1117; color:#f5f0e8; padding:9px 10px;
    text-align:center; font-weight:600; white-space:nowrap; font-size:.72rem; }
.dist-table thead th:first-child { background:#1a1d26; text-align:left; min-width:140px;
    font-family:'DM Mono',monospace; font-size:.63rem; letter-spacing:1px; text-transform:uppercase; }
.dist-table tbody tr:nth-child(even) { background:#f8fafc; }
.dist-table tbody tr:hover td { background:#f0f9ff !important; }
.dist-table tbody td { padding:8px 10px; text-align:center; border-bottom:1px solid #e2e8f0;
    border-right:1px solid #f1f5f9; vertical-align:middle; }
.dist-table tbody td:first-child { font-weight:600; text-align:left; color:#0f1117;
    white-space:nowrap; background:#fafaf8; border-right:2px solid #cbd5e1; }
.cell-self { background:#f1f5f9 !important; color:#94a3b8; }
.val-line  { color:#1e40af; font-weight:700; font-size:.82rem; }
.val-road  { color:#166534; font-weight:700; font-size:.82rem; }
.val-label { font-size:.6rem; color:#94a3b8; font-family:'DM Mono',monospace; }
.val-nd    { color:#fca5a5; font-size:.7rem; }
.serp-badge { display:inline-block; padding:1px 7px; border-radius:10px;
    font-family:'DM Mono',monospace; font-size:.68rem; font-weight:700; }
/* Tabela serpentividade */
.serp-table { width:100%; border-collapse:collapse; font-size:.82rem; margin-top:8px; }
.serp-table thead th { background:#0f1117; color:#f5f0e8; padding:9px 12px;
    text-align:center; font-weight:600; font-size:.73rem; }
.serp-table thead th:first-child,.serp-table thead th:nth-child(2) { text-align:left; background:#1a1d26; }
.serp-table tbody tr:nth-child(even) { background:#f8fafc; }
.serp-table tbody tr:hover td { background:#f0f9ff !important; }
.serp-table tbody td { padding:8px 12px; border-bottom:1px solid #e2e8f0; vertical-align:middle; }
.serp-table tbody td:nth-child(3),.serp-table tbody td:nth-child(4),.serp-table tbody td:nth-child(5) { text-align:center; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
    <div class="header-tag">// análise geoespacial · Brasil</div>
    <h1 class="header-title">🗺️ Distâncias entre <em>Cidades</em></h1>
    <div class="header-sub">linha reta · por estrada · índice de serpentividade · sedes municipais</div>
</div>
""", unsafe_allow_html=True)

# ── Dados ──────────────────────────────────────────────────────────────────────
CAPITALS_DEFAULT = [
    {"name":"Rio Branco",    "state":"AC","lat":-9.97499, "lon":-67.82471,"address":"R. Benjamim Constant, 945"},
    {"name":"Maceió",        "state":"AL","lat":-9.66583, "lon":-35.73528,"address":"Praça Thomaz Espíndola, s/n"},
    {"name":"Macapá",        "state":"AP","lat": 0.03444, "lon":-51.06639,"address":"Av. Iracema Carvão do Nascimento, 600"},
    {"name":"Manaus",        "state":"AM","lat":-3.10194, "lon":-60.02500,"address":"Av. Brasil, 2971 - Compensa"},
    {"name":"Salvador",      "state":"BA","lat":-12.97111,"lon":-38.51083,"address":"Praça Municipal, s/n"},
    {"name":"Fortaleza",     "state":"CE","lat":-3.72389, "lon":-38.54306,"address":"Av. Desembargador Moreira, 310"},
    {"name":"Brasília",      "state":"DF","lat":-15.78361,"lon":-47.89833,"address":"Palácio do Buriti"},
    {"name":"Vitória",       "state":"ES","lat":-20.31944,"lon":-40.33778,"address":"Av. Marechal Mascarenhas de Moraes, 1927"},
    {"name":"Goiânia",       "state":"GO","lat":-16.67861,"lon":-49.25389,"address":"Av. do Cerrado, 999"},
    {"name":"São Luís",      "state":"MA","lat":-2.52972, "lon":-44.30278,"address":"Rua Afonso Pena, s/n"},
    {"name":"Cuiabá",        "state":"MT","lat":-15.59611,"lon":-56.09667,"address":"Av. General Mello, s/n"},
    {"name":"Campo Grande",  "state":"MS","lat":-20.44278,"lon":-54.64611,"address":"Av. Afonso Pena, 3297"},
    {"name":"Belo Horizonte","state":"MG","lat":-19.91722,"lon":-43.93444,"address":"Av. Afonso Pena, 1212"},
    {"name":"Belém",         "state":"PA","lat":-1.45583, "lon":-48.50444,"address":"Praça Felipe Patroni, s/n"},
    {"name":"João Pessoa",   "state":"PB","lat":-7.11528, "lon":-34.86278,"address":"Praça João Pessoa, s/n"},
    {"name":"Curitiba",      "state":"PR","lat":-25.42944,"lon":-49.27167,"address":"Av. Cândido de Abreu, 817"},
    {"name":"Recife",        "state":"PE","lat":-8.05361, "lon":-34.88111,"address":"Av. Cais do Apolo, 925"},
    {"name":"Teresina",      "state":"PI","lat":-5.08917, "lon":-42.80194,"address":"R. Areolino de Abreu, 900"},
    {"name":"Rio de Janeiro","state":"RJ","lat":-22.90278,"lon":-43.17444,"address":"R. Afonso Cavalcanti, 455"},
    {"name":"Natal",         "state":"RN","lat":-5.79500, "lon":-35.21139,"address":"Av. Deodoro da Fonseca, 384"},
    {"name":"Porto Velho",   "state":"RO","lat":-8.76194, "lon":-63.90389,"address":"Av. 7 de Setembro, 237"},
    {"name":"Boa Vista",     "state":"RR","lat": 2.81972, "lon":-60.67333,"address":"Rua Coronel Pinto, 241"},
    {"name":"Porto Alegre",  "state":"RS","lat":-30.03444,"lon":-51.21750,"address":"Av. Loureiro da Silva, 255"},
    {"name":"Florianópolis", "state":"SC","lat":-27.59500,"lon":-48.54861,"address":"R. Timóteo Pereira da Costa, 10"},
    {"name":"Aracaju",       "state":"SE","lat":-10.91111,"lon":-37.07167,"address":"Av. Dr. Carlos Firpo, s/n"},
    {"name":"São Paulo",     "state":"SP","lat":-23.55028,"lon":-46.63361,"address":"Viaduto do Chá, 15"},
    {"name":"Palmas",        "state":"TO","lat":-10.18611,"lon":-48.33361,"address":"Quadra 502 Sul, Av. NS-02"},
]

# ── Funções ────────────────────────────────────────────────────────────────────
def haversine(la1, lo1, la2, lo2):
    R = 6371
    dL, dO = math.radians(la2-la1), math.radians(lo2-lo1)
    a = math.sin(dL/2)**2 + math.cos(math.radians(la1))*math.cos(math.radians(la2))*math.sin(dO/2)**2
    return round(R*2*math.atan2(math.sqrt(a), math.sqrt(1-a)), 1)

def get_road_distance(c1, c2, ors_key):
    try:
        r = requests.post(
            "https://api.openrouteservice.org/v2/directions/driving-car",
            headers={"Authorization": ors_key, "Content-Type": "application/json"},
            json={"coordinates": [[c1["lon"],c1["lat"]], [c2["lon"],c2["lat"]]]},
            timeout=15
        )
        if r.status_code == 200:
            d = r.json()
            m = (d.get("routes",[{}])[0].get("summary",{}).get("distance") or
                 d["features"][0]["properties"]["segments"][0]["distance"])
            return round(m/1000, 1)
        return None
    except Exception:
        return None

def serp_idx(line, road):
    if line and road and line > 0:
        return round(road/line, 3)
    return None

# ── Excel ──────────────────────────────────────────────────────────────────────
def build_excel(cities, matrix):
    wb = Workbook()
    hf  = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    hfd = PatternFill("solid", fgColor="0F1117")
    hfl = PatternFill("solid", fgColor="1A1D26")
    bf  = Font(name="Arial", bold=True, size=9)
    af  = PatternFill("solid", fgColor="F8FAFC")
    sf  = PatternFill("solid", fgColor="F1F5F9")
    thin = Side(style="thin", color="E2E8F0")
    bdr  = Border(bottom=thin, right=thin)

    def make_matrix(title, val_fn, font_fn):
        ws = wb.create_sheet(title)
        c0 = ws.cell(1,1,"Origem / Destino")
        c0.font = Font(name="Arial",bold=True,color="FFFFFF",size=8)
        c0.fill = hfl; c0.alignment = Alignment(horizontal="left",vertical="center")
        ws.column_dimensions["A"].width = 22
        for j,c in enumerate(cities,2):
            cell = ws.cell(1,j,f"{c['name']}, {c['state']}")
            cell.font = hf; cell.fill = hfd
            cell.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
            ws.column_dimensions[get_column_letter(j)].width = 14
        ws.row_dimensions[1].height = 28
        for i,c1 in enumerate(cities):
            r = i+2
            nc = ws.cell(r,1,f"{c1['name']}, {c1['state']}")
            nc.font = bf; nc.alignment = Alignment(vertical="center")
            if i%2==1: nc.fill = af
            for j,c2 in enumerate(cities,2):
                cell = ws.cell(r,j)
                cell.alignment = Alignment(horizontal="center",vertical="center")
                cell.border = bdr
                if c1["name"]==c2["name"]:
                    cell.value="—"; cell.fill=sf
                    cell.font=Font(name="Arial",color="94A3B8",size=10)
                else:
                    v = val_fn(c1,c2)
                    cell.value = v
                    cell.font = font_fn(v)
                    if isinstance(v,float): cell.number_format='#,##0.0'
                    if i%2==1 and v is not None: cell.fill=af
                ws.row_dimensions[r].height = 18
        return ws

    wb.remove(wb.active)

    # Aba 1 — Linha Reta
    make_matrix("Linha Reta (km)",
        lambda c1,c2: haversine(c1["lat"],c1["lon"],c2["lat"],c2["lon"]),
        lambda v: Font(name="Arial",bold=True,color="1E40AF",size=9) if isinstance(v,float)
                  else Font(name="Arial",size=9))

    # Aba 2 — Por Estrada
    def road_val(c1,c2):
        v = matrix.get(f"{c1['name']}-{c2['name']}",{}).get("road")
        return v
    def road_font(v):
        if isinstance(v,float): return Font(name="Arial",bold=True,color="166534",size=9)
        return Font(name="Arial",color="FCA5A5",size=8)
    ws2 = make_matrix("Por Estrada (km)", road_val, road_font)
    for row in ws2.iter_rows(min_row=2):
        for cell in row[1:]:
            if cell.value is None and cell.row>1 and cell.column>1:
                cell.value="N/D"

    # Aba 3 — Serpentividade
    serp_colors = {"Direta":("15803D","DCFCE7"), "Moderada":("92400E","FEF3C7"), "Sinuosa":("991B1B","FEE2E2")}
    def serp_val(c1,c2):
        d = matrix.get(f"{c1['name']}-{c2['name']}",{})
        return serp_idx(d.get("line"), d.get("road"))
    def serp_font_fn(v):
        if not isinstance(v,float): return Font(name="Arial",size=9)
        label,_,_ = serp_class(v)
        fc = serp_colors.get(label,("000000","FFFFFF"))[0]
        return Font(name="Arial",bold=True,color=fc,size=9)
    ws3 = make_matrix("Índice de Serpentividade", serp_val, serp_font_fn)
    # Colorir fundo das células
    for i,c1 in enumerate(cities):
        for j,c2 in enumerate(cities,2):
            if c1["name"]!=c2["name"]:
                d = matrix.get(f"{c1['name']}-{c2['name']}",{})
                v = serp_idx(d.get("line"),d.get("road"))
                if v:
                    label,_,_ = serp_class(v)
                    bg = serp_colors.get(label,("000000","FFFFFF"))[1]
                    ws3.cell(i+2,j).fill = PatternFill("solid",fgColor=bg)
                    ws3.cell(i+2,j).number_format='0.000'
                else:
                    ws3.cell(i+2,j).value="N/D"
                    ws3.cell(i+2,j).font=Font(name="Arial",color="FCA5A5",size=8)

    # Aba 4 — Ranking Serpentividade
    ws4 = wb.create_sheet("Ranking Serpentividade")
    headers4 = ["#","Cidade A","UF","Cidade B","UF","Reta (km)","Estrada (km)","Índice","Classificação"]
    widths4  = [4,20,5,20,5,12,14,10,14]
    for j,(h,w) in enumerate(zip(headers4,widths4),1):
        c=ws4.cell(1,j,h); c.font=hf; c.fill=hfd
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws4.column_dimensions[get_column_letter(j)].width=w
    ws4.row_dimensions[1].height=22
    rows4=[]
    for c1,c2 in combinations(cities,2):
        d = matrix.get(f"{c1['name']}-{c2['name']}",{})
        line=d.get("line"); road=d.get("road")
        idx=serp_idx(line,road)
        if idx: rows4.append((c1,c2,line,road,idx))
    rows4.sort(key=lambda x: x[4], reverse=True)
    for rank,(c1,c2,line,road,idx) in enumerate(rows4,1):
        label,fc,bg=serp_class(idx)
        vals=[rank,c1["name"],c1["state"],c2["name"],c2["state"],line,road,idx,label]
        for j,v in enumerate(vals,1):
            cell=ws4.cell(rank+1,j,v)
            cell.alignment=Alignment(horizontal="center" if j not in [2,4] else "left",vertical="center")
            cell.border=bdr
            if rank%2==0: cell.fill=af
            if j==8:
                cell.font=Font(name="Arial",bold=True,color=fc[1:] if fc.startswith("#") else fc,size=9)
                cell.fill=PatternFill("solid",fgColor=bg[1:] if bg.startswith("#") else bg)
                cell.number_format='0.000'
            elif j==9:
                cell.font=Font(name="Arial",bold=True,color=fc[1:] if fc.startswith("#") else fc,size=9)
                cell.fill=PatternFill("solid",fgColor=bg[1:] if bg.startswith("#") else bg)
            elif j in [6,7]:
                cell.font=Font(name="Arial",size=9); cell.number_format='#,##0.0'
            else:
                cell.font=Font(name="Arial",size=9)
        ws4.row_dimensions[rank+1].height=16

    # Aba 5 — Coordenadas
    ws5=wb.create_sheet("Coordenadas das Sedes")
    h5=["Cidade","UF","Latitude","Longitude","Endereço da Sede"]
    w5=[20,5,14,14,44]
    for j,(h,w) in enumerate(zip(h5,w5),1):
        c=ws5.cell(1,j,h); c.font=hf; c.fill=hfd
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws5.column_dimensions[get_column_letter(j)].width=w
    ws5.row_dimensions[1].height=22
    for i,c in enumerate(cities,2):
        for j,v in enumerate([c["name"],c["state"],c["lat"],c["lon"],c["address"]],1):
            cell=ws5.cell(i,j,v)
            cell.font=Font(name="Arial",size=9)
            cell.alignment=Alignment(horizontal="center" if j>2 else "left",vertical="center")
            cell.border=bdr
            if isinstance(v,float): cell.number_format='0.00000'
            if i%2==0: cell.fill=af
        ws5.row_dimensions[i].height=16

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ── Importar Excel existente ───────────────────────────────────────────────────
def import_from_xlsx(uploaded_file):
    """Lê o Excel exportado pelo app e retorna matrix e lista de cidades."""
    import openpyxl
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)

    # Lê coordenadas da aba Coordenadas
    cities_map = {}
    if "Coordenadas das Sedes" in wb.sheetnames:
        ws = wb["Coordenadas das Sedes"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                name, state, lat, lon = row[0], row[1], row[2], row[3]
                key = f"{name}, {state}"
                cities_map[key] = {"name": name, "state": state,
                                   "lat": float(lat), "lon": float(lon),
                                   "address": row[4] if len(row) > 4 else ""}

    # Lê distâncias da aba Por Estrada
    matrix = {}
    road_data = {}
    if "Por Estrada (km)" in wb.sheetnames:
        ws = wb["Por Estrada (km)"]
        headers = [ws.cell(1, j).value for j in range(2, ws.max_column + 1)]
        for i, row in enumerate(ws.iter_rows(min_row=2, min_col=1, values_only=True)):
            c1_name = row[0]
            if not c1_name: continue
            for j, val in enumerate(row[1:]):
                c2_name = headers[j] if j < len(headers) else None
                if c2_name and c1_name != c2_name and isinstance(val, (int, float)):
                    road_data[f"{c1_name}|{c2_name}"] = float(val)

    # Lê distâncias da aba Linha Reta
    line_data = {}
    if "Linha Reta (km)" in wb.sheetnames:
        ws = wb["Linha Reta (km)"]
        headers = [ws.cell(1, j).value for j in range(2, ws.max_column + 1)]
        for i, row in enumerate(ws.iter_rows(min_row=2, min_col=1, values_only=True)):
            c1_name = row[0]
            if not c1_name: continue
            for j, val in enumerate(row[1:]):
                c2_name = headers[j] if j < len(headers) else None
                if c2_name and c1_name != c2_name and isinstance(val, (int, float)):
                    line_data[f"{c1_name}|{c2_name}"] = float(val)

    # Monta matrix no formato do app (chave = "NomeCidade1-NomeCidade2" sem UF)
    all_pairs = set(list(road_data.keys()) + list(line_data.keys()))
    for pair_key in all_pairs:
        parts = pair_key.split("|")
        if len(parts) != 2: continue
        c1_full, c2_full = parts
        # Extrai só o nome (sem ", UF")
        c1_short = c1_full.split(",")[0].strip()
        c2_short = c2_full.split(",")[0].strip()
        road = road_data.get(pair_key)
        line = line_data.get(pair_key) or line_data.get(f"{c2_full}|{c1_full}")
        mk = f"{c1_short}-{c2_short}"
        mk2 = f"{c2_short}-{c1_short}"
        entry = {"line": line, "road": road}
        matrix[mk] = entry
        matrix[mk2] = entry

    # Descobre cidades presentes
    imported_cities = []
    seen = set()
    for c in CAPITALS:
        full = f"{c['name']}, {c['state']}"
        if full in cities_map and c["name"] not in seen:
            imported_cities.append(c)
            seen.add(c["name"])

    pairs_with_road = sum(1 for v in matrix.values() if v.get("road") is not None)
    pairs_total = len(matrix) // 2

    return matrix, imported_cities, pairs_total, pairs_with_road

# ── Session state ──────────────────────────────────────────────────────────────
if "capitals" not in st.session_state:
    st.session_state.capitals = [dict(c) for c in CAPITALS_DEFAULT]
CAPITALS = st.session_state.capitals  # sempre aponta para a lista persistida

for k,v in [("selected",[c["name"] for c in CAPITALS]),("matrix",{}),
            ("calculated",False),("calc_cities",[]),("has_ors",False)]:
    if k not in st.session_state: st.session_state[k]=v

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuração")
    ors_key = st.text_input("Chave OpenRouteService", type="password",
        placeholder="5b3ce3597851...",
        help="Gratuita em openrouteservice.org — necessária para distâncias por estrada e índice de serpentividade")
    if ors_key:
        st.success("✓ Chave configurada")
    else:
        st.info("Sem chave: apenas linha reta ✈\n\n[Obter chave gratuita →](https://openrouteservice.org/dev/#/signup)")

    st.markdown("---")
    st.markdown("### 📍 Selecionar Cidades")
    c1b,c2b=st.columns(2)
    if c1b.button("✓ Todas",use_container_width=True):
        st.session_state.selected=[c["name"] for c in CAPITALS]
        for c in CAPITALS:
            st.session_state[f"cb_{c['name']}"]=True
        st.session_state.calculated=False; st.rerun()
    if c2b.button("✗ Nenhuma",use_container_width=True):
        st.session_state.selected=[]
        for c in CAPITALS:
            st.session_state[f"cb_{c['name']}"]=False
        st.session_state.calculated=False; st.rerun()
    new_sel=[]
    for c in CAPITALS:
        val = st.session_state.get(f"cb_{c['name']}", c["name"] in st.session_state.selected)
        if st.checkbox(f"{c['name']} ({c['state']})", value=val, key=f"cb_{c['name']}"):
            new_sel.append(c["name"])
    if set(new_sel)!=set(st.session_state.selected):
        st.session_state.selected=new_sel; st.session_state.calculated=False
    st.markdown(f"**{len(st.session_state.selected)}** cidade(s) selecionada(s)")
    st.markdown("---")
    st.markdown("### 📥 Importar Excel anterior")
    uploaded = st.file_uploader(
        "Carregue um .xlsx exportado pelo app",
        type=["xlsx"],
        help="Importa distâncias já calculadas — o app completará apenas os pares faltantes"
    )
    if uploaded:
        if st.button("⬆ Importar dados", use_container_width=True):
            try:
                matrix, imp_cities, total, with_road = import_from_xlsx(uploaded)
                st.session_state.matrix = matrix
                st.session_state.calc_cities = imp_cities
                st.session_state.calculated = True
                st.session_state.has_ors = with_road > 0
                st.session_state.selected = [c["name"] for c in imp_cities]
                # Calcula pares faltantes para retomar
                from itertools import combinations as comb2
                all_pairs = list(comb2(imp_cities, 2))
                pending = [(c1,c2) for c1,c2 in all_pairs
                           if matrix.get(f"{c1['name']}-{c2['name']}", {}).get("road") is None]
                st.session_state.pending_pairs = pending
                st.session_state.done_count = total - len(pending)
                st.session_state.total_pairs_count = total
                st.success(f"✅ Importado! {with_road} pares com estrada, {len(pending)} faltando.")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao importar: {e}")

# ── Main ───────────────────────────────────────────────────────────────────────
selected_cities=[c for c in CAPITALS if c["name"] in st.session_state.selected]

# Aviso se nenhuma cidade selecionada
if not selected_cities:
    st.warning("Selecione pelo menos 2 cidades no menu lateral.")

btn_col1, btn_col2, _ = st.columns([2,2,4])
with btn_col1:
    calc_clicked = st.button("⚡ Calcular Distâncias", disabled=len(selected_cities)<2,
                              use_container_width=True, type="primary")

# ── Cálculo ────────────────────────────────────────────────────────────────────
# ── Botão retomar ────────────────────────────────────────────────────────────
can_resume = (st.session_state.get("pending_pairs") and
              st.session_state.get("calc_cities") == selected_cities)

with btn_col1:
    pass  # já renderizado acima

resume_clicked = False
if can_resume:
    pending = st.session_state.pending_pairs
    done_count = st.session_state.get("done_count", 0)
    total_count = st.session_state.get("total_pairs_count", 0)
    st.info(f"⚠️ Cálculo incompleto: **{done_count}/{total_count}** pares concluídos. "
            f"Restam **{len(pending)}** pares.")
    resume_clicked = st.button("▶ Retomar Cálculo", use_container_width=False, type="secondary")

def run_calculation(pairs_to_calc, ors_key, has_ors, existing_matrix):
    matrix = dict(existing_matrix)
    all_pairs_count = st.session_state.get("total_pairs_count", len(pairs_to_calc))
    done_so_far = st.session_state.get("done_count", 0)

    prog = st.progress(int(done_so_far/all_pairs_count*100), text="Iniciando…")
    status = st.empty()
    remaining = list(pairs_to_calc)

    for idx, (c1, c2) in enumerate(remaining):
        line = haversine(c1["lat"], c1["lon"], c2["lat"], c2["lon"])
        road = None
        if has_ors:
            status.markdown(f"🚗 **{c1['name']} → {c2['name']}** "
                            f"({done_so_far+idx+1}/{all_pairs_count})")
            road = get_road_distance(c1, c2, ors_key)

        key = f"{c1['name']}-{c2['name']}"
        matrix[key] = {"line": line, "road": road}
        matrix[f"{c2['name']}-{c1['name']}"] = {"line": line, "road": road}

        # Salva progresso parcial a cada par
        st.session_state.matrix = matrix
        st.session_state.done_count = done_so_far + idx + 1
        st.session_state.pending_pairs = remaining[idx+1:]

        pct = int((done_so_far+idx+1)/all_pairs_count*100)
        prog.progress(pct, text=f"Calculando… {done_so_far+idx+1}/{all_pairs_count}")

    prog.progress(100, text="✅ Concluído!")
    status.empty()
    st.session_state.pending_pairs = []
    return matrix

if calc_clicked and len(selected_cities)>=2:
    pairs=list(combinations(selected_cities,2))
    has_ors=bool(ors_key and len(ors_key)>10)

    if has_ors:
        test=get_road_distance(selected_cities[0],selected_cities[1],ors_key)
        if test is None:
            st.error("❌ Chave ORS inválida ou sem créditos. Calculando apenas linha reta.")
            has_ors=False

    # Inicializa estado
    st.session_state.pending_pairs = pairs
    st.session_state.done_count = 0
    st.session_state.total_pairs_count = len(pairs)
    st.session_state.calc_cities = selected_cities
    st.session_state.has_ors = has_ors
    st.session_state.calculated = False

    matrix = run_calculation(pairs, ors_key, has_ors, {})
    st.session_state.matrix = matrix
    st.session_state.calculated = True
    st.session_state.has_ors = has_ors

elif resume_clicked:
    has_ors = st.session_state.get("has_ors", False)
    pending = st.session_state.pending_pairs
    existing = st.session_state.get("matrix", {})

    if has_ors and not ors_key:
        st.error("Cole a chave ORS novamente para retomar.")
    else:
        matrix = run_calculation(pending, ors_key, has_ors, existing)
        st.session_state.matrix = matrix
        st.session_state.calculated = True

# ── Resultados ─────────────────────────────────────────────────────────────────

# Aba de Cidades sempre visível (mesmo antes de calcular)
st.markdown("---")

# ── TABS PRINCIPAIS ────────────────────────────────────────────────────────────
tab_calc, tab_serp, tab_cities = st.tabs([
    "📊 Matriz de Distâncias",
    "🐍 Índice de Serpentividade",
    "📍 Cidades & Geolocalização"
])

# ── TAB CIDADES ────────────────────────────────────────────────────────────────
with tab_cities:
    st.markdown("#### Base de Cidades")
    st.markdown("Cidades disponíveis para cálculo. Selecione as desejadas no menu lateral.")

    # Tabela de cidades com indicação de selecionada
    tbl_c = ("<div style='overflow-x:auto'><table class='dist-table'>"
             "<thead><tr>"
             "<th style='text-align:left'>Cidade</th>"
             "<th>UF</th>"
             "<th>Latitude (Sede)</th>"
             "<th>Longitude (Sede)</th>"
             "<th>Endereço da Sede</th>"
             "<th>Selecionada</th>"
             "</tr></thead><tbody>")
    for i, c in enumerate(CAPITALS):
        sel = c["name"] in st.session_state.selected
        badge = ("<span style='background:#dcfce7;color:#15803d;padding:2px 9px;"
                 "border-radius:10px;font-size:.72rem;font-weight:700'>✓ Sim</span>" if sel else
                 "<span style='background:#f1f5f9;color:#94a3b8;padding:2px 9px;"
                 "border-radius:10px;font-size:.72rem'>— Não</span>")
        bg = "background:#f0fdf4;" if sel else ""
        tbl_c += (f"<tr style='{bg}'>"
                  f"<td style='font-weight:600;text-align:left'>{c['name']}</td>"
                  f"<td style='text-align:center;font-family:monospace;font-size:.8rem'>{c['state']}</td>"
                  f"<td style='text-align:center;font-family:monospace;font-size:.78rem'>{c['lat']:.5f}</td>"
                  f"<td style='text-align:center;font-family:monospace;font-size:.78rem'>{c['lon']:.5f}</td>"
                  f"<td style='text-align:left;font-size:.78rem;color:#64748b'>{c.get('address','')}</td>"
                  f"<td style='text-align:center'>{badge}</td>"
                  f"</tr>")
    tbl_c += "</tbody></table></div>"
    st.markdown(tbl_c, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### ➕ Adicionar Nova Cidade")
    st.caption("Preencha os dados da sede municipal (prefeitura) da cidade que deseja adicionar.")

    with st.form("form_add_city", clear_on_submit=True):
        fc1, fc2 = st.columns([3,1])
        fc3, fc4 = st.columns(2)
        fc5, fc6 = st.columns([3,1])
        new_name    = fc1.text_input("Nome da Cidade", placeholder="Ex: Campinas")
        new_state   = fc2.text_input("UF", placeholder="SP", max_chars=2)
        new_lat     = fc3.number_input("Latitude (Sede)", value=-23.0, format="%.5f", step=0.00001)
        new_lon     = fc4.number_input("Longitude (Sede)", value=-47.0, format="%.5f", step=0.00001)
        new_address = fc5.text_input("Endereço da Sede", placeholder="Rua XV de Novembro, 1000 - Centro")
        submitted   = fc6.form_submit_button("➕ Adicionar", use_container_width=True, type="primary")

        if submitted:
            if not new_name or not new_state:
                st.error("Nome e UF são obrigatórios.")
            elif any(c["name"].lower() == new_name.strip().lower() for c in st.session_state.capitals):
                st.warning(f"'{new_name}' já está na base.")
            else:
                new_city = {
                    "name": new_name.strip(),
                    "state": new_state.strip().upper(),
                    "lat": float(new_lat),
                    "lon": float(new_lon),
                    "address": new_address.strip()
                }
                st.session_state.capitals.append(new_city)
                st.session_state.selected.append(new_name.strip())
                st.session_state[f"cb_{new_name.strip()}"] = True
                st.success(f"✅ '{new_name.strip()}, {new_state.strip().upper()}' adicionada!")
                st.rerun()

if st.session_state.calculated and st.session_state.calc_cities:
    cities=st.session_state.calc_cities
    matrix=st.session_state.matrix
    has_ors=st.session_state.has_ors

    if not has_ors:
        st.info("ℹ️ Distâncias por estrada e índice de serpentividade indisponíveis sem a chave ORS.")

    # ── Tabs de resultados ─────────────────────────────────────────────────────
    tab1, tab2 = tab_calc, tab_serp  # alias para os resultados

    # ── TAB 1: Matriz ──────────────────────────────────────────────────────────
    with tab1:
        st.markdown("""
        <div style='display:flex;gap:20px;margin-bottom:12px;font-size:.78rem;color:#64748b;font-family:monospace'>
            <span><b style='color:#1e40af'>■</b> ✈ Linha reta (km)</span>
            <span><b style='color:#166534'>■</b> 🚗 Por estrada (km)</span>
            <span><b style='color:#15803d'>■</b> 🐍 Índice serpentividade</span>
        </div>""", unsafe_allow_html=True)

        tbl="<div style='overflow-x:auto'><table class='dist-table'><thead><tr><th>Origem / Destino</th>"
        for c in cities:
            tbl+=f"<th>{c['name']}<br><span style='font-family:monospace;font-size:.58rem;opacity:.6'>{c['state']}</span></th>"
        tbl+="</tr></thead><tbody>"
        for c1 in cities:
            tbl+=f"<tr><td>{c1['name']}, {c1['state']}</td>"
            for c2 in cities:
                if c1["name"]==c2["name"]:
                    tbl+="<td class='cell-self'>—</td>"
                else:
                    d=matrix.get(f"{c1['name']}-{c2['name']}",{})
                    line=d.get("line"); road=d.get("road")
                    idx=serp_idx(line,road)
                    road_html=(f"<div><span class='val-road'>{road} km</span> <span class='val-label'>🚗</span></div>"
                               if road else "<div><span class='val-nd'>N/D</span> <span class='val-label'>🚗</span></div>")
                    if idx:
                        label,fc,bg=serp_class(idx)
                        serp_html=(f"<div><span class='serp-badge' style='color:{fc};background:{bg}'>"
                                   f"{idx:.3f} · {label}</span></div>")
                    else:
                        serp_html=""
                    tbl+=(f"<td><div><span class='val-line'>{line} km</span> <span class='val-label'>✈</span></div>"
                          f"{road_html}{serp_html}</td>")
            tbl+="</tr>"
        tbl+="</tbody></table></div>"
        st.markdown(tbl, unsafe_allow_html=True)

    # ── TAB 2: Ranking Serpentividade ──────────────────────────────────────────
    with tab2:
        st.markdown("""
        <div style='margin-bottom:14px;font-size:.82rem;color:#475569;line-height:1.6'>
            O <b>Índice de Serpentividade</b> = Distância por Estrada ÷ Distância em Linha Reta.<br>
            Quanto mais próximo de <b>1.0</b>, mais direta é a estrada. Valores altos indicam rotas sinuosas.
        </div>
        <div style='display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap'>
            <span style='background:#dcfce7;color:#15803d;padding:4px 12px;border-radius:10px;font-size:.78rem;font-family:monospace;font-weight:700'>🟢 &lt; 1.20 — Direta</span>
            <span style='background:#fef3c7;color:#92400e;padding:4px 12px;border-radius:10px;font-size:.78rem;font-family:monospace;font-weight:700'>🟡 1.20–1.50 — Moderada</span>
            <span style='background:#fee2e2;color:#991b1b;padding:4px 12px;border-radius:10px;font-size:.78rem;font-family:monospace;font-weight:700'>🔴 &gt; 1.50 — Sinuosa</span>
        </div>
        """, unsafe_allow_html=True)

        if not has_ors:
            st.warning("Configure a chave ORS para calcular o índice de serpentividade.")
        else:
            # Ranking ordenado
            rows=[]
            for c1,c2 in combinations(cities,2):
                d=matrix.get(f"{c1['name']}-{c2['name']}",{})
                line=d.get("line"); road=d.get("road")
                idx=serp_idx(line,road)
                if idx: rows.append({"c1":c1,"c2":c2,"line":line,"road":road,"idx":idx})
            rows.sort(key=lambda x:x["idx"],reverse=True)

            tbl2=("<div style='overflow-x:auto'><table class='serp-table'>"
                  "<thead><tr><th>#</th><th>Cidade A</th><th>Cidade B</th>"
                  "<th>Reta (km)</th><th>Estrada (km)</th><th>Índice</th><th>Classificação</th></tr></thead><tbody>")
            for rank,r in enumerate(rows,1):
                label,fc,bg=serp_class(r["idx"])
                badge=(f"<span class='serp-badge' style='color:{fc};background:{bg}'>{label}</span>")
                tbl2+=(f"<tr>"
                       f"<td style='text-align:center;color:#94a3b8;font-size:.75rem'>{rank}</td>"
                       f"<td><b>{r['c1']['name']}</b>, {r['c1']['state']}</td>"
                       f"<td><b>{r['c2']['name']}</b>, {r['c2']['state']}</td>"
                       f"<td style='text-align:center;color:#1e40af;font-weight:700'>{r['line']}</td>"
                       f"<td style='text-align:center;color:#166534;font-weight:700'>{r['road']}</td>"
                       f"<td style='text-align:center'><span style='font-family:monospace;font-weight:700;color:{fc}'>{r['idx']:.3f}</span></td>"
                       f"<td style='text-align:center'>{badge}</td>"
                       f"</tr>")
            tbl2+="</tbody></table></div>"
            st.markdown(tbl2, unsafe_allow_html=True)

            # Métricas resumo
            st.markdown("---")
            idxs=[r["idx"] for r in rows]
            m1,m2,m3,m4=st.columns(4)
            m1.metric("Média geral", f"{sum(idxs)/len(idxs):.3f}")
            m2.metric("Mais direta 🟢", f"{min(idxs):.3f}",
                      f"{rows[-1]['c1']['name']} ↔ {rows[-1]['c2']['name']}")
            m3.metric("Mais sinuosa 🔴", f"{max(idxs):.3f}",
                      f"{rows[0]['c1']['name']} ↔ {rows[0]['c2']['name']}")
            diretas=sum(1 for i in idxs if i<1.2)
            m4.metric("Rotas diretas (< 1.20)", f"{diretas}/{len(idxs)}")

    # ── Download ───────────────────────────────────────────────────────────────
    st.markdown("---")
    with btn_col2:
        excel=build_excel(cities,matrix)
        st.download_button("📊 Baixar Excel", data=excel,
            file_name="distancias-cidades.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
