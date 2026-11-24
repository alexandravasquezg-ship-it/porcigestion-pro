import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
from streamlit_gsheets import GSheetsConnection
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

st.set_page_config(
    page_title="PorciGestión Pro | Control Integral Porcino",
    page_icon="🐷",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos visuales
st.markdown("""
<style>
    .main-header {font-size: 2rem; font-weight: 800; color: #1E8449; margin-bottom: 0.2rem;}
    .sub-header {font-size: 1rem; color: #566573; margin-bottom: 1.5rem;}
    .metric-card {background-color: #F8F9F9; padding: 15px; border-radius: 10px; border-left: 5px solid #1E8449;}
    .alert-box {padding: 12px; border-radius: 8px; margin-bottom: 10px;}
</style>
""", unsafe_allow_html=True)

# Conexión a Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

ESQUEMA_HOJAS = {
    "Reproductores": [
        "ID_Animal", "Tipo", "Linea_Genetica", "Banda", "Padre_ID", "Madre_ID",
        "Partos", "Abortos", "Repeticiones", "Estado", "Racion_Diaria_Kg",
        "Costo_Acumulado_Ciclo", "Fecha_Ultimo_Cambio_Estado", "Observaciones"
    ],
    "Reproduccion": [
        "ID_Evento", "ID_Cerda", "Raza_Cerda", "Banda", "Fecha_Servicio", "Tipo_Servicio",
        "Macho_O_Semen", "Uso_Recelador", "Costo_Servicio", "Fecha_Eco_28d",
        "Estado_Prenez", "Fecha_Prob_Parto", "Fecha_Real_Parto", "Nacidos_Vivos",
        "Mortinatos", "Momias", "Nombre_Lote_Asignado", "Fecha_Prob_Destete",
        "Destetados_Reales", "Peso_Destete_Total_Kg", "Costo_Por_Lechon"
    ],
    "Lotes_Engorde": [
        "ID_Lote", "Madre_Origen", "Raza_Cruce", "Corral", "Etapa", "Fecha_Inicio",
        "Edad_Inicial_Dias", "Cant_Inicial", "Cant_Actual", "Peso_Prom_Inicial_Kg",
        "Peso_Prom_Actual_Kg", "Kg_Alimento_Consumido", "Costo_Inicial_Lechones",
        "Costo_Alimento", "Costo_Sanidad", "Estado"
    ],
    "Bodega": [
        "ID_Compra", "Fecha", "Categoria_Insumo", "Nombre_Insumo", "Presentacion",
        "Unidades_Compradas", "Contenido_Por_Unidad", "Unidad_Medida", "Precio_Unitario",
        "Costo_Por_Medida", "Saldo_Disponible"
    ],
    "Consumo_Alimento": [
        "Fecha", "Destino_Tipo", "ID_Destino", "Tipo_Concentrado", "Kilos_Servidos", "Costo_Total"
    ],
    "Sanidad_Eventos": [
        "Fecha", "ID_Destino", "Categoria", "Detalle_Sintoma", "Medicamento",
        "Cantidad_Aplicada", "Unidad_Medida", "Costo", "Dias_Retiro",
        "Fecha_Fin_Retiro", "Estado_Caso"
    ],
    "Ventas": [
        "Fecha", "ID_Origen", "Modalidad", "Cantidad_Animales", "Peso_Pie_Kg",
        "Peso_Carne_Canal_Kg", "Costo_Sacrificio", "Ingreso_Bruto",
        "Costo_Produccion_Asignado", "Utilidad_Neta"
    ],
    "Inversiones_Gastos": [
        "ID_Movimiento", "Fecha", "Tipo_Movimiento", "Categoria", "Descripcion",
        "Monto", "Responsable", "Observaciones"
    ],
    "Configuracion": [
        "Parametro", "Valor", "Descripcion"
    ]
}

CONFIG_DEFAULT = pd.DataFrame([
    {"Parametro": "Peso_Meta_Kg", "Valor": "105.0", "Descripcion": "Peso objetivo del cerdo finalizado en pie (kg)"},
    {"Parametro": "Edad_Meta_Dias", "Valor": "150", "Descripcion": "Edad objetivo desde nacimiento para alcanzar el peso meta (días)"},
    {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": "1.4", "Descripcion": "Peso promedio al nacer (kg)"},
    {"Parametro": "Intervalo_Bandas_Dias", "Valor": "15", "Descripcion": "Días entre inseminación de una banda y la siguiente (14, 15 o 21)"},
    {"Parametro": "Numero_Bandas", "Valor": "10", "Descripcion": "Total de bandas en el ciclo completo (10 para quincenal, 7 para 21 días)"},
    {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": "2", "Descripcion": "Máximo de cerdas a inseminar por banda/quincena"},
    {"Parametro": "Alerta_Dias_Abiertos", "Valor": "7", "Descripcion": "Días máximos que una cerda vacía puede estar sin servicio"}
])

def asegurar_columnas(df, cols):
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]

@st.cache_data(ttl=120)
def cargar_datos():
    datos = {}
    for hoja, cols in ESQUEMA_HOJAS.items():
        try:
            df = conn.read(worksheet=hoja, ttl=0)
            df = df.dropna(how="all")
            datos[hoja] = asegurar_columnas(df, cols)
        except Exception:
            if hoja == "Configuracion":
                datos[hoja] = CONFIG_DEFAULT.copy()
            else:
                datos[hoja] = pd.DataFrame(columns=cols)
    if datos["Configuracion"].empty:
        datos["Configuracion"] = CONFIG_DEFAULT.copy()
    return datos

def guardar_hoja(hoja, df):
    df_clean = asegurar_columnas(df.copy(), ESQUEMA_HOJAS[hoja])
    conn.update(worksheet=hoja, data=df_clean)
    st.cache_data.clear()

def to_num(series, default=0.0):
    return pd.to_numeric(series, errors="coerce").fillna(default)

def parse_fecha(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        return pd.to_datetime(str(val).strip()[:10]).date()
    except Exception:
        return None

def get_config(df_cfg, param, default_val):
    if df_cfg.empty:
        return default_val
    fila = df_cfg[df_cfg["Parametro"] == param]
    if fila.empty:
        return default_val
    try:
        return float(fila.iloc[0]["Valor"])
    except Exception:
        return default_val

# Cargar datos iniciales
datos = cargar_datos()
df_rep = datos["Reproductores"]
df_repro = datos["Reproduccion"]
df_lotes = datos["Lotes_Engorde"]
df_bodega = datos["Bodega"]
df_consumo = datos["Consumo_Alimento"]
df_sanidad = datos["Sanidad_Eventos"]
df_ventas = datos["Ventas"]
df_inv = datos["Inversiones_Gastos"]
df_cfg = datos["Configuracion"]

hoy = date.today()

# Parámetros globales
PESO_META = get_config(df_cfg, "Peso_Meta_Kg", 105.0)
EDAD_META = int(get_config(df_cfg, "Edad_Meta_Dias", 150))
PESO_NAC = get_config(df_cfg, "Peso_Nacimiento_Prom_Kg", 1.4)
INTERVALO_BANDAS = int(get_config(df_cfg, "Intervalo_Bandas_Dias", 15))
NUM_BANDAS = int(get_config(df_cfg, "Numero_Bandas", 10))
CUPO_BANDA = int(get_config(df_cfg, "Cupo_Max_Cerdas_Banda", 2))
MAX_DIAS_ABIERTOS = int(get_config(df_cfg, "Alerta_Dias_Abiertos", 7))
GDP_META_G = ((PESO_META - PESO_NAC) / max(EDAD_META, 1)) * 1000.0

# Control de Acceso en Pantalla Principal (Operario sin clave / Administrador con clave obligatoria)
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "rol_usuario" not in st.session_state:
    st.session_state["rol_usuario"] = None

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
    with col_l2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.image("https://img.icons8.com/color/96/pig.png", width=85)
        st.markdown('<div class="main-header">PorciGestión Pro</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Sistema Inteligente de Administración y Analítica Porcina</div>', unsafe_allow_html=True)

        perfil_sel = st.selectbox(
            "Selecciona tu perfil de ingreso:",
            ["Operario de Granja (Trabajador)", "Administrador (Gerencia)"]
        )

        if perfil_sel == "Operario de Granja (Trabajador)":
            st.info("👷 Ingreso directo para tareas operativas diarias (alimentación, pesajes, partos y sanidad). No requiere contraseña.")
            if st.button("Ingresar como Operario", type="primary", use_container_width=True):
                st.session_state["autenticado"] = True
                st.session_state["rol_usuario"] = "operario"
                st.rerun()
        else:
            clave_ing = st.text_input("Contraseña de Administrador:", type="password", placeholder="Ingresa la clave de administrador...")
            if st.button("Ingresar como Administrador", type="primary", use_container_width=True):
                if clave_ing == "admin123":
                    st.session_state["autenticado"] = True
                    st.session_state["rol_usuario"] = "admin"
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta. El acceso de Administrador está restringido.")
    st.stop()

es_admin = (st.session_state["rol_usuario"] == "admin")

# Barra lateral según el perfil autenticado
with st.sidebar:
    st.image("https://img.icons8.com/color/96/pig.png", width=70)
    st.title("PorciGestión Pro")
    if es_admin:
        st.success("👑 Perfil: **Administrador**")
    else:
        st.info("👷 Perfil: **Operario de Granja**")

    if st.button("🚪 Cerrar Sesión / Cambiar Perfil", use_container_width=True):
        st.session_state["autenticado"] = False
        st.session_state["rol_usuario"] = None
        st.rerun()

    if st.button("🔄 Sincronizar con Google Sheets", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    if es_admin:
        opciones_menu = [
            "🚨 Tablero de Alertas",
            "📅 Escalonamiento y Bandas",
            "🧬 Reproducción (IA y Partos)",
            "🐖 Lotes, Madres y Unión",
            "🌾 Bodega y Botiquín",
            "💉 Sanidad y Mortalidad",
            "💰 Ventas (Pie / Canal)",
            "🏗️ Inversiones y Gastos Fijos",
            "📊 Analítica, Metas y Reportes",
            "⚙️ Gestionar Cerdas y Metas"
        ]
    else:
        opciones_menu = [
            "🚨 Tablero de Alertas",
            "🧬 Reproducción (IA y Partos)",
            "🐖 Lotes, Madres y Unión",
            "🌾 Bodega y Botiquín",
            "💉 Sanidad y Mortalidad"
        ]

    menu = st.radio("Módulos del Sistema", opciones_menu)

# ==============================================================================
# CÁLCULO CENTRAL DE ALERTAS E INDICADORES
# ==============================================================================
alertas_criticas = []
alertas_preventivas = []

if not df_repro.empty:
    for _, r in df_repro.iterrows():
        est = str(r["Estado_Prenez"]).strip()
        cerda = str(r["ID_Cerda"])
        f_parto = parse_fecha(r["Fecha_Prob_Parto"])
        f_eco = parse_fecha(r["Fecha_Eco_28d"])
        f_dest = parse_fecha(r["Fecha_Prob_Destete"])

        if est == "Preñada" and f_parto:
            dias_p = (f_parto - hoy).days
            if dias_p <= 5:
                alertas_criticas.append(f"🔴 **Parto Próximo ({cerda} - {r.get('Banda','')}):** Programado para `{f_parto}` (en {dias_p} días). Preparar sala de maternidad.")
        if est == "Pendiente Eco" and f_eco:
            if hoy >= f_eco:
                alertas_preventivas.append(f"🔵 **Confirmar Preñez ({cerda}):** Cumplió 28 días desde inseminación (`{f_eco}`). Realizar ecografía o chequeo de retorno a celo.")
        if est == "En Lactancia" and f_dest:
            dias_d = (hoy - f_dest).days
            if dias_d >= -2:
                alertas_preventivas.append(f"🍼 **Destete Listo ({cerda} / Lote {r.get('Nombre_Lote_Asignado','')}):** Fecha objetivo `{f_dest}`. Destetar para reingresar madre a servicio.")

if not df_rep.empty:
    hembras_vacias = df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (df_rep["Estado"] == "Vacía")]
    for _, h in hembras_vacias.iterrows():
        f_cambio = parse_fecha(h.get("Fecha_Ultimo_Cambio_Estado", ""))
        if f_cambio:
            dias_ab = (hoy - f_cambio).days
            if dias_ab > MAX_DIAS_ABIERTOS:
                alertas_criticas.append(f"⏳ **Cerda con Días Abiertos ({h['ID_Animal']} - {h.get('Banda','')}):** Lleva **{dias_ab} días vacía** (límite: {MAX_DIAS_ABIERTOS} días) consumiendo concentrado sin gestar.")

lotes_activos = df_lotes[df_lotes["Estado"] == "Activo"].copy() if not df_lotes.empty else pd.DataFrame()
if not lotes_activos.empty:
    for _, l in lotes_activos.iterrows():
        f_ini = parse_fecha(l["Fecha_Inicio"])
        edad_ini = int(to_num(pd.Series([l.get("Edad_Inicial_Dias", 0)])).iloc[0])
        dias_lote = (hoy - f_ini).days if f_ini else 0
        edad_actual = edad_ini + max(dias_lote, 1)
        peso_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])

        peso_esperado_hoy = PESO_NAC + ((PESO_META - PESO_NAC) / max(EDAD_META, 1)) * edad_actual
        dif_peso = peso_act - peso_esperado_hoy
        if edad_actual >= 35 and dif_peso < -4.0:
            alertas_preventivas.append(
                f"📉 **Bajo Ritmo de Crecimiento ({l['ID_Lote']}):** Pesa **{peso_act:.1f} kg** a los {edad_actual} días, cuando la meta a 5 meses exige **{peso_esperado_hoy:.1f} kg** ({dif_peso:+.1f} kg). Revisar ración o sanidad."
            )

if not df_sanidad.empty:
    for _, s in df_sanidad.iterrows():
        f_ret = parse_fecha(s.get("Fecha_Fin_Retiro", ""))
        if f_ret and f_ret >= hoy:
            alertas_criticas.append(f"🚫 **Tiempo de Retiro Vigente ({s['ID_Destino']}):** Medicamento `{s['Medicamento']}`. Prohibida venta/sacrificio hasta `{f_ret}`.")

if not df_bodega.empty:
    df_bodega["Saldo_Num"] = to_num(df_bodega["Saldo_Disponible"])
    resumen_stock = df_bodega.groupby(["Categoria_Insumo", "Nombre_Insumo", "Unidad_Medida"], as_index=False)["Saldo_Num"].sum()
    for _, st_row in resumen_stock.iterrows():
        if st_row["Categoria_Insumo"] == "Concentrado" and st_row["Saldo_Num"] < 120:
            alertas_preventivas.append(f"🌾 **Stock Bajo Concentrado ({st_row['Nombre_Insumo']}):** Quedan **{st_row['Saldo_Num']:.1f} kg** ({st_row['Saldo_Num']/40:.1f} bultos de 40 kg).")
        elif st_row["Categoria_Insumo"] != "Concentrado" and st_row["Saldo_Num"] < 20:
            alertas_preventivas.append(f"💊 **Botiquín Bajo ({st_row['Nombre_Insumo']}):** Quedan **{st_row['Saldo_Num']:.1f} {st_row['Unidad_Medida']}**.")

# ==============================================================================
# MÓDULO 1: TABLERO DE ALERTAS
# ==============================================================================
if menu == "🚨 Tablero de Alertas":
    st.markdown('<div class="main-header">🚨 Centro de Control y Alertas Operativas</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Fecha del sistema: <b>{hoy.strftime("%d/%m/%Y")}</b> | Meta productiva: <b>{PESO_META:.0f} kg a los {EDAD_META} días ({GDP_META_G:.0f} g/día)</b></div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    total_madres = len(df_rep[df_rep["Tipo"].str.contains("Hembra", na=False)]) if not df_rep.empty else 0
    total_cerdos_engorde = int(to_num(lotes_activos["Cant_Actual"]).sum()) if not lotes_activos.empty else 0
    utilidad_lotes = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0

    c1.metric("🐷 Cerdas Madres Activas", total_madres)
    c2.metric("🐖 Cerdos en Crecimiento", total_cerdos_engorde)
    if es_admin:
        c3.metric("💵 Utilidad Operativa Lotes", f"${utilidad_lotes:,.0f}")
    else:
        c3.metric("🌾 Lotes Activos en Galpón", len(lotes_activos))
    c4.metric("🔔 Alertas Activas", len(alertas_criticas) + len(alertas_preventivas))

    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🔴 Alertas Críticas (Acción Inmediata)")
        if alertas_criticas:
            for a in alertas_criticas:
                st.error(a)
        else:
            st.success("✅ No hay alertas críticas en este momento.")

    with col_b:
        st.subheader("🟡 Alertas Preventivas y Seguimiento")
        if alertas_preventivas:
            for a in alertas_preventivas:
                st.warning(a)
        else:
            st.info("✅ Todos los indicadores preventivos están dentro de meta.")

# ==============================================================================
# MÓDULO 2: ESCALONAMIENTO Y BANDAS DE PARTO
# ==============================================================================
elif menu == "📅 Escalonamiento y Bandas":
    st.markdown('<div class="main-header">📅 Planificador de Escalonamiento por Bandas</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="sub-header">Ciclo biológico: <b>114 días gestación + 28 días lactancia + 5 días destete-celo = 147 días</b>. '
        f'Configuración actual: <b>{NUM_BANDAS} bandas cada {INTERVALO_BANDAS} días</b> (máx. {CUPO_BANDA} cerdas/banda).</div>',
        unsafe_allow_html=True
    )

    hembras = df_rep[df_rep["Tipo"].str.contains("Hembra", na=False)].copy() if not df_rep.empty else pd.DataFrame()
    lista_bandas = [f"Banda {i:02d}" for i in range(1, NUM_BANDAS + 1)]

    tab1, tab2, tab3 = st.tabs(["📊 Distribución Actual de Bandas", "🗓️ Cronograma Maestro de Rotación", "➕ Incorporar / Reorganizar Cerdas"])

    with tab1:
        if hembras.empty:
            st.info("No hay hembras registradas aún.")
        else:
            hembras["Banda_Limpia"] = hembras["Banda"].replace("", "Sin Asignar")
            resumen_bandas = []
            for b in lista_bandas:
                sub = hembras[hembras["Banda_Limpia"] == b]
                cupo_disp = CUPO_BANDA - len(sub)
                estado_cupo = "🟢 Disponible" if cupo_disp > 0 else ("✅ Completa" if cupo_disp == 0 else "🔴 Sobrecupo")
                resumen_bandas.append({
                    "Banda": b,
                    "Cerdas Asignadas": len(sub),
                    "Cupo Ideal": CUPO_BANDA,
                    "Espacios Libres": max(cupo_disp, 0),
                    "Estado Banda": estado_cupo,
                    "IDs Cerdas": ", ".join(sub["ID_Animal"].tolist()) if not sub.empty else "—"
                })
            df_res_bandas = pd.DataFrame(resumen_bandas)
            st.dataframe(df_res_bandas, use_container_width=True, hide_index=True)

            sin_asignar = hembras[hembras["Banda_Limpia"] == "Sin Asignar"]
            if not sin_asignar.empty:
                st.warning(f"⚠️ Tienes **{len(sin_asignar)} cerdas sin banda asignada**: {', '.join(sin_asignar['ID_Animal'].tolist())}. Asígnalas en la pestaña 'Incorporar / Reorganizar Cerdas'.")

    with tab2:
        st.subheader("Proyección de Vuelta Completa (Inseminación → Parto → Destete → Venta)")
        fecha_inicio_ciclo = st.date_input("Fecha de inicio de la Banda 01 (referencia)", value=hoy)
        filas_cron = []
        for idx, b in enumerate(lista_bandas):
            f_ins = fecha_inicio_ciclo + timedelta(days=idx * INTERVALO_BANDAS)
            f_par = f_ins + timedelta(days=114)
            f_des = f_par + timedelta(days=28)
            f_reins = f_des + timedelta(days=5)
            f_venta = f_par + timedelta(days=EDAD_META)
            cerdas_b = hembras[hembras["Banda"] == b]["ID_Animal"].tolist() if not hembras.empty else []
            filas_cron.append({
                "Banda": b,
                "Cerdas": ", ".join(cerdas_b) if cerdas_b else "(Cupo libre)",
                "Inseminación": f_ins.strftime("%Y-%m-%d"),
                "Parto (114d)": f_par.strftime("%Y-%m-%d"),
                "Destete (28d)": f_des.strftime("%Y-%m-%d"),
                "Nueva Inseminación (+5d)": f_reins.strftime("%Y-%m-%d"),
                f"Salida Venta ({EDAD_META}d)": f_venta.strftime("%Y-%m-%d")
            })
        df_cron = pd.DataFrame(filas_cron)
        st.dataframe(df_cron, use_container_width=True, hide_index=True)
        st.info("💡 **Cómo funciona el ciclo:** Cuando terminas de inseminar la última banda, la **Banda 01** ya cumplió sus 114 días de gestación + 28 de lactancia + 5 de descanso y entra exactamente de nuevo al turno.")

    with tab3:
        if not es_admin:
            st.warning("Inicia sesión como Administrador para reorganizar bandas.")
        elif hembras.empty:
            st.info("Primero agrega cerdas en 'Gestionar Cerdas y Metas'.")
        else:
            col_izq, col_der = st.columns(2)
            with col_izq:
                st.subheader("Asignación Manual por Cerda")
                cerda_sel = st.selectbox("Seleccionar Cerda", hembras["ID_Animal"].tolist())
                banda_nueva = st.selectbox("Asignar a Banda", lista_bandas)
                if st.button("💾 Guardar Banda de esta Cerda"):
                    df_rep.loc[df_rep["ID_Animal"] == cerda_sel, "Banda"] = banda_nueva
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{cerda_sel} asignada a {banda_nueva}.")
                    st.rerun()

            with col_der:
                st.subheader("Reparto Automático Equilibrado")
                st.write("Distribuye automáticamente todas las cerdas activas entre las bandas sin sobrepasar el cupo por quincena.")
                if st.button("🪄 Distribuir Todas las Cerdas Automáticamente"):
                    idx_h = df_rep[df_rep["Tipo"].str.contains("Hembra", na=False)].index.tolist()
                    for i, row_idx in enumerate(idx_h):
                        banda_asignada = lista_bandas[i % NUM_BANDAS]
                        df_rep.at[row_idx, "Banda"] = banda_asignada
                    guardar_hoja("Reproductores", df_rep)
                    st.success("Cerdas distribuidas equilibradamente en las bandas.")
                    st.rerun()

# ==============================================================================
# MÓDULO 3: REPRODUCCIÓN (IA Y PARTOS)
# ==============================================================================
elif menu == "🧬 Reproducción (IA y Partos)":
    st.markdown('<div class="main-header">🧬 Gestión Reproductiva (Inseminación Artificial)</div>', unsafe_allow_html=True)

    hembras = df_rep[df_rep["Tipo"].str.contains("Hembra", na=False)] if not df_rep.empty else pd.DataFrame()
    t1, t2, t3, t4 = st.tabs(["1️⃣ Registrar Inseminación", "2️⃣ Confirmar Preñez / Falla", "3️⃣ Registrar Parto", "4️⃣ Registrar Destete"])

    with t1:
        if hembras.empty:
            st.info("Agrega cerdas madres primero.")
        else:
            c1, c2 = st.columns(2)
            cerda_ia = c1.selectbox("Cerda a Inseminar", hembras["ID_Animal"].tolist())
            info_c = hembras[hembras["ID_Animal"] == cerda_ia].iloc[0]
            banda_c = str(info_c.get("Banda", "Banda 01")) or "Banda 01"
            raza_c = str(info_c["Linea_Genetica"])

            f_serv = c2.date_input("Fecha de Inseminación", value=hoy)
            semen_raza = c1.selectbox("Raza del Semen / Pajilla", ["Duroc", "Pietrain", "Landrace", "Yorkshire", "Large White", "Hampshire"])
            recelador = c2.selectbox("Estímulo con Verraco Recelador", ["Sí (V-01)", "No"])
            costo_ia = c1.number_input("Costo Total Servicio / Pajillas ($)", min_value=0.0, value=45000.0, step=5000.0)

            f_parto_est = f_serv + timedelta(days=114)
            partos_cercanos = 0
            if not df_repro.empty:
                for _, rx in df_repro.iterrows():
                    if str(rx["Estado_Prenez"]) in ["Preñada", "Pendiente Eco"]:
                        fp = parse_fecha(rx["Fecha_Prob_Parto"])
                        if fp and abs((fp - f_parto_est).days) <= 7:
                            partos_cercanos += 1

            if partos_cercanos >= CUPO_BANDA:
                st.warning(f"⚠️ **Alerta de Escalonamiento:** Ya hay **{partos_cercanos} cerdas** programadas para parir alrededor de `{f_parto_est}` (cupo máximo por banda: {CUPO_BANDA}).")
            else:
                st.success(f"✅ **Cupo Óptimo:** Esta inseminación corresponde a **{banda_c}** (parto probable: `{f_parto_est}`). Hay {partos_cercanos}/{CUPO_BANDA} cupos ocupados en esa ventana.")

            if st.button("💾 Guardar Inseminación"):
                nuevo_id = f"INS-{len(df_repro)+1:03d}"
                nueva_fila = pd.DataFrame([{
                    "ID_Evento": nuevo_id,
                    "ID_Cerda": cerda_ia,
                    "Raza_Cerda": raza_c,
                    "Banda": banda_c,
                    "Fecha_Servicio": f_serv.strftime("%Y-%m-%d"),
                    "Tipo_Servicio": "Inseminación Artificial",
                    "Macho_O_Semen": semen_raza,
                    "Uso_Recelador": recelador,
                    "Costo_Servicio": costo_ia,
                    "Fecha_Eco_28d": (f_serv + timedelta(days=28)).strftime("%Y-%m-%d"),
                    "Estado_Prenez": "Pendiente Eco",
                    "Fecha_Prob_Parto": f_parto_est.strftime("%Y-%m-%d"),
                    "Fecha_Real_Parto": "",
                    "Nacidos_Vivos": 0,
                    "Mortinatos": 0,
                    "Momias": 0,
                    "Nombre_Lote_Asignado": "",
                    "Fecha_Prob_Destete": (f_parto_est + timedelta(days=28)).strftime("%Y-%m-%d"),
                    "Destetados_Reales": 0,
                    "Peso_Destete_Total_Kg": 0,
                    "Costo_Por_Lechon": 0
                }])
                df_repro = pd.concat([df_repro, nueva_fila], ignore_index=True)
                df_rep.loc[df_rep["ID_Animal"] == cerda_ia, "Estado"] = "Inseminada"
                df_rep.loc[df_rep["ID_Animal"] == cerda_ia, "Fecha_Ultimo_Cambio_Estado"] = f_serv.strftime("%Y-%m-%d")
                df_rep.loc[df_rep["ID_Animal"] == cerda_ia, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_ia, "Costo_Acumulado_Ciclo"]) + costo_ia
                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Servicio {nuevo_id} registrado para {cerda_ia}.")
                st.rerun()

    with t2:
        pendientes = df_repro[df_repro["Estado_Prenez"].isin(["Pendiente Eco", "Preñada"])] if not df_repro.empty else pd.DataFrame()
        if pendientes.empty:
            st.info("No hay servicios pendientes de confirmación.")
        else:
            ev_sel = st.selectbox("Seleccionar Evento", pendientes["ID_Evento"] + " | " + pendientes["ID_Cerda"] + " (" + pendientes["Macho_O_Semen"] + ")")
            id_ev = ev_sel.split(" | ")[0]
            res_eco = st.selectbox("Resultado Diagnóstico", ["Preñada", "Repitió Celo", "Aborto"])
            if st.button("💾 Actualizar Estado Reproductivo"):
                cerda_id = df_repro.loc[df_repro["ID_Evento"] == id_ev, "ID_Cerda"].iloc[0]
                df_repro.loc[df_repro["ID_Evento"] == id_ev, "Estado_Prenez"] = res_eco
                if res_eco == "Preñada":
                    df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Estado"] = "Preñada"
                else:
                    df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Estado"] = "Vacía"
                    df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")
                    if res_eco == "Repitió Celo":
                        df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Repeticiones"] = to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Repeticiones"]) + 1
                    elif res_eco == "Aborto":
                        df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Abortos"] = to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Abortos"]) + 1
                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Reproductores", df_rep)
                st.success("Estado actualizado.")
                st.rerun()

    with t3:
        prenadas = df_repro[df_repro["Estado_Prenez"] == "Preñada"] if not df_repro.empty else pd.DataFrame()
        if prenadas.empty:
            st.info("No hay cerdas preñadas pendientes de parto.")
        else:
            ev_p = st.selectbox("Cerda que Parió", prenadas["ID_Evento"] + " | " + prenadas["ID_Cerda"])
            id_ev_p = ev_p.split(" | ")[0]
            fila_p = prenadas[prenadas["ID_Evento"] == id_ev_p].iloc[0]
            c1, c2, c3 = st.columns(3)
            f_parto_r = c1.date_input("Fecha Real de Parto", value=hoy)
            vivos = c2.number_input("Nacidos Vivos", min_value=1, max_value=25, value=12)
            mortinatos = c3.number_input("Mortinatos / Momias", min_value=0, max_value=10, value=0)
            nombre_lote = st.text_input("Nombre del Nuevo Lote", value=f"LOTE-{len(df_lotes)+1:02d}-{fila_p['ID_Cerda']}")
            corral_parto = st.text_input("Corral / Paridera", value=f"Paridera-{fila_p['ID_Cerda']}")

            if st.button("💾 Registrar Parto y Crear Lote Lactante"):
                cerda_id = fila_p["ID_Cerda"]
                costo_ciclo = float(to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Costo_Acumulado_Ciclo"]).iloc[0])
                cruce_txt = f"{fila_p['Raza_Cerda']} × {fila_p['Macho_O_Semen']}"

                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Estado_Prenez"] = "En Lactancia"
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Fecha_Real_Parto"] = f_parto_r.strftime("%Y-%m-%d")
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Nacidos_Vivos"] = vivos
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Mortinatos"] = mortinatos
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Nombre_Lote_Asignado"] = nombre_lote
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Fecha_Prob_Destete"] = (f_parto_r + timedelta(days=28)).strftime("%Y-%m-%d")

                nuevo_lote = pd.DataFrame([{
                    "ID_Lote": nombre_lote,
                    "Madre_Origen": cerda_id,
                    "Raza_Cruce": cruce_txt,
                    "Corral": corral_parto,
                    "Etapa": "Lactante",
                    "Fecha_Inicio": f_parto_r.strftime("%Y-%m-%d"),
                    "Edad_Inicial_Dias": 0,
                    "Cant_Inicial": vivos,
                    "Cant_Actual": vivos,
                    "Peso_Prom_Inicial_Kg": PESO_NAC,
                    "Peso_Prom_Actual_Kg": PESO_NAC,
                    "Kg_Alimento_Consumido": 0.0,
                    "Costo_Inicial_Lechones": costo_ciclo,
                    "Costo_Alimento": 0.0,
                    "Costo_Sanidad": 0.0,
                    "Estado": "Activo"
                }])
                df_lotes = pd.concat([df_lotes, nuevo_lote], ignore_index=True)

                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Estado"] = "Lactante"
                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Partos"] = to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Partos"]) + 1
                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Costo_Acumulado_Ciclo"] = 0.0
                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = f_parto_r.strftime("%Y-%m-%d")

                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Lotes_Engorde", df_lotes)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Parto registrado. Lote {nombre_lote} creado.")
                st.rerun()

    with t4:
        lactantes = df_repro[df_repro["Estado_Prenez"] == "En Lactancia"] if not df_repro.empty else pd.DataFrame()
        if lactantes.empty:
            st.info("No hay camadas en lactancia.")
        else:
            ev_d = st.selectbox("Camada a Destetar", lactantes["ID_Evento"] + " | " + lactantes["ID_Cerda"] + " -> " + lactantes["Nombre_Lote_Asignado"])
            id_ev_d = ev_d.split(" | ")[0]
            fila_d = lactantes[lactantes["ID_Evento"] == id_ev_d].iloc[0]
            c1, c2 = st.columns(2)
            dest_reales = c1.number_input("Lechones Destetados", min_value=1, max_value=25, value=int(to_num(pd.Series([fila_d["Nacidos_Vivos"]])).iloc[0]))
            peso_tot_dest = c2.number_input("Peso Total Camada Destetada (kg)", min_value=1.0, value=float(dest_reales * 7.2))
            corral_dest = st.text_input("Corral de Precebo / Levante", value="Corral-Levante-1")

            if st.button("💾 Confirmar Destete y Pasar a Levante"):
                cerda_id = fila_d["ID_Cerda"]
                lote_id = fila_d["Nombre_Lote_Asignado"]
                costo_lact = float(to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Costo_Acumulado_Ciclo"]).iloc[0])

                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Estado_Prenez"] = "Destetada"
                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Destetados_Reales"] = dest_reales
                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Peso_Destete_Total_Kg"] = peso_tot_dest

                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Etapa"] = "Levante"
                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Corral"] = corral_dest
                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Cant_Actual"] = dest_reales
                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Peso_Prom_Actual_Kg"] = round(peso_tot_dest / dest_reales, 2)
                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Costo_Inicial_Lechones"] = to_num(df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Costo_Inicial_Lechones"]) + costo_lact

                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Estado"] = "Vacía"
                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Costo_Acumulado_Ciclo"] = 0.0
                df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")

                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Lotes_Engorde", df_lotes)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Destete completado. {cerda_id} pasa a Vacía (lista para servicio en ~5 días).")
                st.rerun()

    st.divider()
    st.subheader("Historial Reproductivo")
    st.dataframe(df_repro, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 4: LOTES, MADRES Y UNIÓN
# ==============================================================================
elif menu == "🐖 Lotes, Madres y Unión":
    st.markdown('<div class="main-header">🐖 Control Productivo y Crecimiento por Lote</div>', unsafe_allow_html=True)

    if lotes_activos.empty:
        st.info("No hay lotes activos.")
    else:
        filas_ind = []
        for _, l in lotes_activos.iterrows():
            f_ini = parse_fecha(l["Fecha_Inicio"])
            edad_ini = int(to_num(pd.Series([l.get("Edad_Inicial_Dias", 0)])).iloc[0])
            dias_trans = max((hoy - f_ini).days, 1) if f_ini else 1
            edad_dias = edad_ini + dias_trans

            cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
            p_ini = float(to_num(pd.Series([l["Peso_Prom_Inicial_Kg"]])).iloc[0])
            p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
            kg_alim = float(to_num(pd.Series([l["Kg_Alimento_Consumido"]])).iloc[0])

            c_ini = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0])
            c_ali = float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0])
            c_san = float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
            c_tot = c_ini + c_ali + c_san

            gdp_real_g = ((p_act - p_ini) / dias_trans) * 1000.0
            peso_esp_hoy = PESO_NAC + ((PESO_META - PESO_NAC) / max(EDAD_META, 1)) * edad_dias
            cumpl_pct = (p_act / max(peso_esp_hoy, 0.1)) * 100.0

            kg_ganados_tot = max((p_act - p_ini) * cant, 0.1)
            ca = kg_alim / kg_ganados_tot if kg_alim > 0 else 0.0
            peso_tot_pie = p_act * cant
            punto_eq = c_tot / max(peso_tot_pie, 1.0)

            fila_data = {
                "Lote": l["ID_Lote"],
                "Cruce": l["Raza_Cruce"],
                "Etapa": l["Etapa"],
                "Cerdos": cant,
                "Edad (días)": edad_dias,
                "Peso Real (kg)": round(p_act, 1),
                "Peso Meta Hoy (kg)": round(peso_esp_hoy, 1),
                "Cumpl. Curva (%)": round(cumpl_pct, 1),
                "GDP Real (g/día)": round(gdp_real_g, 0),
                "Conv. Alim.": round(ca, 2)
            }
            if es_admin:
                fila_data["Costo Total ($)"] = round(c_tot, 0)
                fila_data["Costo x Cerdo ($)"] = round(c_tot / max(cant, 1), 0)
                fila_data["Punto Equilibrio ($/kg)"] = round(punto_eq, 0)

            filas_ind.append(fila_data)

        df_ind = pd.DataFrame(filas_ind)
        st.subheader("📊 Indicadores y Curva de Crecimiento en Tiempo Real")
        st.dataframe(df_ind, use_container_width=True, hide_index=True)

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("⚖️ Actualizar Pesaje / Etapa de Lote")
            lote_edit = st.selectbox("Seleccionar Lote", lotes_activos["ID_Lote"].tolist())
            fila_le = lotes_activos[lotes_activos["ID_Lote"] == lote_edit].iloc[0]
            nuevo_peso = st.number_input("Nuevo Peso Promedio Actual (kg)", min_value=1.0, value=float(to_num(pd.Series([fila_le["Peso_Prom_Actual_Kg"]])).iloc[0]), step=0.5)
            nueva_etapa = st.selectbox("Etapa", ["Lactante", "Precebo", "Levante", "Engorde"], index=["Lactante", "Precebo", "Levante", "Engorde"].index(fila_le["Etapa"]) if fila_le["Etapa"] in ["Lactante", "Precebo", "Levante", "Engorde"] else 2)
            nuevo_corral = st.text_input("Corral", value=str(fila_le["Corral"]))
            if st.button("💾 Guardar Pesaje"):
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Peso_Prom_Actual_Kg"] = nuevo_peso
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Etapa"] = nueva_etapa
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Corral"] = nuevo_corral
                guardar_hoja("Lotes_Engorde", df_lotes)
                st.success("Pesaje actualizado.")
                st.rerun()

        with c2:
            st.subheader("🔗 Unir / Fusionar Dos Lotes")
            if len(lotes_activos) >= 2:
                l1 = st.selectbox("Lote A (Principal)", lotes_activos["ID_Lote"].tolist(), key="u1")
                l2 = st.selectbox("Lote B (Se integrará al A)", [x for x in lotes_activos["ID_Lote"].tolist() if x != l1], key="u2")
                nuevo_nom_union = st.text_input("Nombre Lote Unificado", value=f"{l1}+{l2}")
                if st.button("🔗 Fusionar Lotes"):
                    r1 = lotes_activos[lotes_activos["ID_Lote"] == l1].iloc[0]
                    r2 = lotes_activos[lotes_activos["ID_Lote"] == l2].iloc[0]
                    c1_n = int(to_num(pd.Series([r1["Cant_Actual"]])).iloc[0])
                    c2_n = int(to_num(pd.Series([r2["Cant_Actual"]])).iloc[0])
                    tot_c = c1_n + c2_n
                    p_prom = (c1_n * float(to_num(pd.Series([r1["Peso_Prom_Actual_Kg"]])).iloc[0]) + c2_n * float(to_num(pd.Series([r2["Peso_Prom_Actual_Kg"]])).iloc[0])) / max(tot_c, 1)

                    df_lotes.loc[df_lotes["ID_Lote"] == l1, "ID_Lote"] = nuevo_nom_union
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Madre_Origen"] = f"{r1['Madre_Origen']} + {r2['Madre_Origen']}"
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Raza_Cruce"] = f"Mixto ({r1['Raza_Cruce']} / {r2['Raza_Cruce']})"
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Cant_Inicial"] = int(to_num(pd.Series([r1["Cant_Inicial"]])).iloc[0]) + int(to_num(pd.Series([r2["Cant_Inicial"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Cant_Actual"] = tot_c
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Peso_Prom_Actual_Kg"] = round(p_prom, 2)
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Kg_Alimento_Consumido"] = float(to_num(pd.Series([r1["Kg_Alimento_Consumido"]])).iloc[0]) + float(to_num(pd.Series([r2["Kg_Alimento_Consumido"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Costo_Inicial_Lechones"] = float(to_num(pd.Series([r1["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Inicial_Lechones"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Costo_Alimento"] = float(to_num(pd.Series([r1["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Alimento"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"] == nuevo_nom_union, "Costo_Sanidad"] = float(to_num(pd.Series([r1["Costo_Sanidad"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Sanidad"]])).iloc[0])

                    df_lotes = df_lotes[df_lotes["ID_Lote"] != l2]
                    guardar_hoja("Lotes_Engorde", df_lotes)
                    st.success(f"Lotes unidos en {nuevo_nom_union} ({tot_c} cerdos).")
                    st.rerun()

# ==============================================================================
# MÓDULO 5: BODEGA Y BOTIQUÍN
# ==============================================================================
elif menu == "🌾 Bodega y Botiquín":
    st.markdown('<div class="main-header">🌾 Inventario de Concentrados y Botiquín Veterinario</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📦 Registrar Compra (Alimento o Medicamento)")
        cat_ins = st.selectbox("Categoría de Insumo", ["Concentrado", "Antibiótico / Inyectable", "Vacuna / Biológico", "Vitamina / Desparasitante"])
        if cat_ins == "Concentrado":
            nom_ins = st.selectbox("Tipo de Concentrado", ["Gestación", "Lactancia", "Preiniciador", "Iniciación", "Levante", "Engorde"])
            pres = "Bulto"
            und_med = "kg"
            cont_und = st.number_input("Kilos por Bulto", min_value=1.0, value=40.0)
        else:
            nom_ins = st.text_input("Nombre del Medicamento / Vacuna", placeholder="Ej: Florfenicol Inyectable")
            pres = st.selectbox("Presentación", ["Frasco", "Caja / Dosis", "Sobre / Bolsa"])
            und_med = st.selectbox("Unidad de Descuento", ["ml", "dosis", "g"])
            cont_und = st.number_input(f"Contenido por {pres} ({und_med})", min_value=1.0, value=100.0)

        cant_comp = st.number_input(f"Cantidad de {pres}s comprados", min_value=1, value=1)
        precio_u = st.number_input(f"Precio por {pres} ($)", min_value=0.0, value=85000.0, step=1000.0)
        tot_medida = cant_comp * cont_und
        costo_med = precio_u / max(cont_und, 1.0)
        st.caption(f"Total ingresado: **{tot_medida:.1f} {und_med}** | Costo unitario: **${costo_med:,.2f} / {und_med}**")

        if st.button("💾 Ingresar a Bodega"):
            if not nom_ins:
                st.error("Escribe el nombre del insumo.")
            else:
                nuevo_id = f"INS-{len(df_bodega)+1:03d}"
                nueva_c = pd.DataFrame([{
                    "ID_Compra": nuevo_id,
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Categoria_Insumo": cat_ins,
                    "Nombre_Insumo": nom_ins,
                    "Presentacion": pres,
                    "Unidades_Compradas": cant_comp,
                    "Contenido_Por_Unidad": cont_und,
                    "Unidad_Medida": und_med,
                    "Precio_Unitario": precio_u,
                    "Costo_Por_Medida": round(costo_med, 2),
                    "Saldo_Disponible": tot_medida
                }])
                df_bodega = pd.concat([df_bodega, nueva_c], ignore_index=True)
                guardar_hoja("Bodega", df_bodega)
                st.success(f"Ingresado {nom_ins} ({tot_medida} {und_med}) a bodega.")
                st.rerun()

    with col2:
        st.subheader("🍽️ Registrar Consumo Diario de Concentrado")
        conc_disp = df_bodega[(df_bodega["Categoria_Insumo"] == "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()
        if conc_disp.empty:
            st.warning("No hay concentrado con saldo disponible en bodega.")
        else:
            tipos_conc = conc_disp["Nombre_Insumo"].unique().tolist()
            t_conc_sel = st.selectbox("Concentrado a Servir", tipos_conc)
            saldo_tot_c = to_num(conc_disp[conc_disp["Nombre_Insumo"] == t_conc_sel]["Saldo_Disponible"]).sum()
            st.info(f"Saldo disponible de **{t_conc_sel}**: **{saldo_tot_c:.1f} kg**")

            dest_tipo = st.selectbox("Destino", ["Lote de Engorde", "Cerda Madre / Verraco"])
            if dest_tipo == "Lote de Engorde":
                lista_dest = lotes_activos["ID_Lote"].tolist() if not lotes_activos.empty else []
            else:
                lista_dest = df_rep["ID_Animal"].tolist() if not df_rep.empty else []

            id_dest = st.selectbox("Seleccionar Lote o Animal", lista_dest)
            kg_serv = st.number_input("Kilos Servidos (kg)", min_value=0.5, max_value=float(saldo_tot_c), value=min(20.0, float(saldo_tot_c)), step=0.5)

            if st.button("💾 Descontar Alimento y Cargar Costo"):
                restante = kg_serv
                costo_acum = 0.0
                filas_idx = df_bodega[(df_bodega["Categoria_Insumo"] == "Concentrado") & (df_bodega["Nombre_Insumo"] == t_conc_sel) & (to_num(df_bodega["Saldo_Disponible"]) > 0)].index
                for idx in filas_idx:
                    if restante <= 0:
                        break
                    disp = float(to_num(pd.Series([df_bodega.at[idx, "Saldo_Disponible"]])).iloc[0])
                    c_kg = float(to_num(pd.Series([df_bodega.at[idx, "Costo_Por_Medida"]])).iloc[0])
                    tomar = min(disp, restante)
                    df_bodega.at[idx, "Saldo_Disponible"] = round(disp - tomar, 2)
                    costo_acum += tomar * c_kg
                    restante -= tomar

                nuevo_cons = pd.DataFrame([{
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Destino_Tipo": dest_tipo,
                    "ID_Destino": id_dest,
                    "Tipo_Concentrado": t_conc_sel,
                    "Kilos_Servidos": kg_serv,
                    "Costo_Total": round(costo_acum, 0)
                }])
                df_consumo = pd.concat([df_consumo, nuevo_cons], ignore_index=True)

                if dest_tipo == "Lote de Engorde":
                    df_lotes.loc[df_lotes["ID_Lote"] == id_dest, "Kg_Alimento_Consumido"] = to_num(df_lotes.loc[df_lotes["ID_Lote"] == id_dest, "Kg_Alimento_Consumido"]) + kg_serv
                    df_lotes.loc[df_lotes["ID_Lote"] == id_dest, "Costo_Alimento"] = to_num(df_lotes.loc[df_lotes["ID_Lote"] == id_dest, "Costo_Alimento"]) + costo_acum
                    guardar_hoja("Lotes_Engorde", df_lotes)
                else:
                    df_rep.loc[df_rep["ID_Animal"] == id_dest, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"] == id_dest, "Costo_Acumulado_Ciclo"]) + costo_acum
                    guardar_hoja("Reproductores", df_rep)

                guardar_hoja("Bodega", df_bodega)
                guardar_hoja("Consumo_Alimento", df_consumo)
                st.success(f"Descontados {kg_serv} kg y cargados a {id_dest}.")
                st.rerun()

    st.divider()
    st.subheader("📊 Saldos Actuales en Bodega y Botiquín")
    if not df_bodega.empty:
        df_bodega["Saldo_Disponible"] = to_num(df_bodega["Saldo_Disponible"])
        res_bod = df_bodega.groupby(["Categoria_Insumo", "Nombre_Insumo", "Unidad_Medida"], as_index=False).agg(
            Saldo_Total=("Saldo_Disponible", "sum"),
            Costo_Promedio_Unidad=("Costo_Por_Medida", "mean")
        )
        if es_admin:
            res_bod["Valor_Inventario ($)"] = (res_bod["Saldo_Total"] * res_bod["Costo_Promedio_Unidad"]).round(0)
        else:
            res_bod = res_bod[["Categoria_Insumo", "Nombre_Insumo", "Unidad_Medida", "Saldo_Total"]]
        st.dataframe(res_bod, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 6: SANIDAD Y MORTALIDAD
# ==============================================================================
elif menu == "💉 Sanidad y Mortalidad":
    st.markdown('<div class="main-header">💉 Sanidad, Botiquín y Control de Mortalidad</div>', unsafe_allow_html=True)

    destinos_posibles = (lotes_activos["ID_Lote"].tolist() if not lotes_activos.empty else []) + (df_rep["ID_Animal"].tolist() if not df_rep.empty else [])
    meds_disp = df_bodega[(df_bodega["Categoria_Insumo"] != "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()

    c1, c2 = st.columns(2)
    dest_san = c1.selectbox("Lote o Animal Afectado", destinos_posibles)
    cat_san = c2.selectbox("Tipo de Evento", ["Tratamiento / Enfermedad", "Vacunación Preventiva", "Mortalidad (Baja)"])
    sintoma = c1.text_input("Detalle / Síntoma / Causa", placeholder="Ej: Tos respiratoria / Vacuna Circovirus")

    usar_botiquin = False
    med_nombre = "Ninguno"
    cant_aplicada = 0.0
    unidad_ap = "ml"
    costo_san = 0.0

    if cat_san != "Mortalidad (Baja)":
        if not meds_disp.empty:
            usar_botiquin = st.checkbox("💊 Descontar medicamento del Botiquín en Bodega", value=True)
        if usar_botiquin and not meds_disp.empty:
            lista_meds = meds_disp["Nombre_Insumo"].unique().tolist()
            med_nombre = c2.selectbox("Seleccionar Medicamento del Botiquín", lista_meds)
            sub_m = meds_disp[meds_disp["Nombre_Insumo"] == med_nombre]
            saldo_m = to_num(sub_m["Saldo_Disponible"]).sum()
            unidad_ap = str(sub_m.iloc[0]["Unidad_Medida"])
            costo_u_m = float(to_num(sub_m["Costo_Por_Medida"]).mean())
            cant_aplicada = c1.number_input(f"Cantidad Aplicada ({unidad_ap}) [Disponible: {saldo_m:.1f}]", min_value=0.1, max_value=float(saldo_m), value=min(5.0, float(saldo_m)), step=0.5)
            costo_san = round(cant_aplicada * costo_u_m, 0)
            if es_admin:
                st.info(f"Costo calculado automáticamente según botiquín: **${costo_san:,.0f}**")
        else:
            med_nombre = c2.text_input("Medicamento Aplicado (Externo)", value="Ninguno")
            costo_san = c1.number_input("Costo del Tratamiento ($)", min_value=0.0, value=0.0, step=1000.0)

    dias_ret = c2.number_input("Días de Retiro (Prohibición de venta)", min_value=0, max_value=60, value=0)
    f_fin_ret = (hoy + timedelta(days=int(dias_ret))).strftime("%Y-%m-%d") if dias_ret > 0 else ""

    if st.button("💾 Registrar Evento Sanitario"):
        if usar_botiquin and cant_aplicada > 0:
            rest = cant_aplicada
            idxs = df_bodega[(df_bodega["Categoria_Insumo"] != "Concentrado") & (df_bodega["Nombre_Insumo"] == med_nombre) & (to_num(df_bodega["Saldo_Disponible"]) > 0)].index
            for idx in idxs:
                if rest <= 0:
                    break
                disp = float(to_num(pd.Series([df_bodega.at[idx, "Saldo_Disponible"]])).iloc[0])
                tomar = min(disp, rest)
                df_bodega.at[idx, "Saldo_Disponible"] = round(disp - tomar, 2)
                rest -= tomar
            guardar_hoja("Bodega", df_bodega)

        nuevo_san = pd.DataFrame([{
            "Fecha": hoy.strftime("%Y-%m-%d"),
            "ID_Destino": dest_san,
            "Categoria": cat_san,
            "Detalle_Sintoma": sintoma,
            "Medicamento": med_nombre,
            "Cantidad_Aplicada": cant_aplicada,
            "Unidad_Medida": unidad_ap if usar_botiquin else "",
            "Costo": costo_san,
            "Dias_Retiro": dias_ret,
            "Fecha_Fin_Retiro": f_fin_ret,
            "Estado_Caso": "En Tratamiento" if dias_ret > 0 else "Cerrado"
        }])
        df_sanidad = pd.concat([df_sanidad, nuevo_san], ignore_index=True)

        if not lotes_activos.empty and dest_san in lotes_activos["ID_Lote"].values:
            if cat_san == "Mortalidad (Baja)":
                df_lotes.loc[df_lotes["ID_Lote"] == dest_san, "Cant_Actual"] = (to_num(df_lotes.loc[df_lotes["ID_Lote"] == dest_san, "Cant_Actual"]) - 1).clip(lower=0)
            df_lotes.loc[df_lotes["ID_Lote"] == dest_san, "Costo_Sanidad"] = to_num(df_lotes.loc[df_lotes["ID_Lote"] == dest_san, "Costo_Sanidad"]) + costo_san
            guardar_hoja("Lotes_Engorde", df_lotes)
        elif not df_rep.empty and dest_san in df_rep["ID_Animal"].values:
            df_rep.loc[df_rep["ID_Animal"] == dest_san, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"] == dest_san, "Costo_Acumulado_Ciclo"]) + costo_san
            guardar_hoja("Reproductores", df_rep)

        guardar_hoja("Sanidad_Eventos", df_sanidad)
        st.success("Evento sanitario registrado e inventarios actualizados.")
        st.rerun()

    st.divider()
    st.subheader("Historial Sanitario y Mortalidad")
    st.dataframe(df_sanidad, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 7: VENTAS (PIE / CANAL) - SOLO ADMINISTRADOR
# ==============================================================================
elif menu == "💰 Ventas (Pie / Canal)":
    st.markdown('<div class="main-header">💰 Registro de Ventas y Liquidación de Lotes</div>', unsafe_allow_html=True)

    if lotes_activos.empty:
        st.info("No hay lotes activos disponibles para vender.")
    else:
        lote_v = st.selectbox("Seleccionar Lote a Vender", lotes_activos["ID_Lote"].tolist())
        fila_v = lotes_activos[lotes_activos["ID_Lote"] == lote_v].iloc[0]

        bloqueado = False
        if not df_sanidad.empty:
            ret_lote = df_sanidad[df_sanidad["ID_Destino"] == lote_v]
            for _, rs in ret_lote.iterrows():
                fr = parse_fecha(rs.get("Fecha_Fin_Retiro", ""))
                if fr and fr >= hoy:
                    st.error(f"🚫 BLOQUEO SANITARIO: Este lote tiene tiempo de retiro vigente hasta `{fr}` por `{rs['Medicamento']}`.")
                    bloqueado = True

        cant_disp = int(to_num(pd.Series([fila_v["Cant_Actual"]])).iloc[0])
        c_tot_lote = float(to_num(pd.Series([fila_v["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Sanidad"]])).iloc[0])

        c1, c2 = st.columns(2)
        modalidad = c1.selectbox("Modalidad de Venta", ["En Pie (Cerdo entero en báscula)", "Beneficiado (Carne en canal por kilos)"])
        cant_v = c2.number_input("Cantidad de Cerdos Vendidos", min_value=1, max_value=max(cant_disp, 1), value=cant_disp)
        peso_pie_v = c1.number_input("Peso Total en Pie (kg)", min_value=1.0, value=float(cant_v * to_num(pd.Series([fila_v["Peso_Prom_Actual_Kg"]])).iloc[0]))

        if "En Pie" in modalidad:
            precio_kg = c2.number_input("Precio Venta por Kilo en Pie ($)", min_value=0.0, value=8500.0, step=100.0)
            peso_canal_v = 0.0
            costo_sac = 0.0
            ingreso_bruto = peso_pie_v * precio_kg
        else:
            rend_pct = c2.number_input("Rendimiento en Canal (%)", min_value=50.0, max_value=90.0, value=78.0)
            peso_canal_v = round(peso_pie_v * (rend_pct / 100.0), 2)
            precio_canal = c1.number_input("Precio Venta Kilo en Canal ($)", min_value=0.0, value=14500.0, step=100.0)
            costo_sac = c2.number_input("Costo Sacrificio / Desposte / Transporte ($)", min_value=0.0, value=float(cant_v * 65000), step=5000.0)
            ingreso_bruto = peso_canal_v * precio_canal

        costo_prop = (c_tot_lote / max(cant_disp, 1)) * cant_v
        utilidad_neta = ingreso_bruto - costo_prop - costo_sac

        m1, m2, m3 = st.columns(3)
        m1.metric("Ingreso Bruto", f"${ingreso_bruto:,.0f}")
        m2.metric("Costo Producción + Sacrificio", f"${(costo_prop + costo_sac):,.0f}")
        m3.metric("Utilidad Neta del Lote", f"${utilidad_neta:,.0f}")

        if st.button("💾 Confirmar Venta", disabled=bloqueado):
            nueva_v = pd.DataFrame([{
                "Fecha": hoy.strftime("%Y-%m-%d"),
                "ID_Origen": lote_v,
                "Modalidad": modalidad,
                "Cantidad_Animales": cant_v,
                "Peso_Pie_Kg": peso_pie_v,
                "Peso_Carne_Canal_Kg": peso_canal_v,
                "Costo_Sacrificio": costo_sac,
                "Ingreso_Bruto": round(ingreso_bruto, 0),
                "Costo_Produccion_Asignado": round(costo_prop, 0),
                "Utilidad_Neta": round(utilidad_neta, 0)
            }])
            df_ventas = pd.concat([df_ventas, nueva_v], ignore_index=True)

            nuevo_saldo_cerdos = cant_disp - cant_v
            if nuevo_saldo_cerdos <= 0:
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Cant_Actual"] = 0
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Estado"] = "Vendido"
            else:
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Cant_Actual"] = nuevo_saldo_cerdos
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Costo_Inicial_Lechones"] = float(to_num(pd.Series([fila_v["Costo_Inicial_Lechones"]])).iloc[0]) * (nuevo_saldo_cerdos / cant_disp)
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Costo_Alimento"] = float(to_num(pd.Series([fila_v["Costo_Alimento"]])).iloc[0]) * (nuevo_saldo_cerdos / cant_disp)
                df_lotes.loc[df_lotes["ID_Lote"] == lote_v, "Costo_Sanidad"] = float(to_num(pd.Series([fila_v["Costo_Sanidad"]])).iloc[0]) * (nuevo_saldo_cerdos / cant_disp)

            guardar_hoja("Ventas", df_ventas)
            guardar_hoja("Lotes_Engorde", df_lotes)
            st.success("Venta registrada y rentabilidad guardada.")
            st.rerun()

    st.divider()
    st.subheader("Historial de Ventas")
    st.dataframe(df_ventas, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 8: INVERSIONES, INFRAESTRUCTURA Y GASTOS FIJOS - SOLO ADMINISTRADOR
# ==============================================================================
elif menu == "🏗️ Inversiones y Gastos Fijos":
    st.markdown('<div class="main-header">🏗️ Inversiones de Capital, Corrales y Gastos Fijos</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Separa la rentabilidad operativa de los cerdos frente a la inversión en infraestructura y costos fijos de la granja.</div>', unsafe_allow_html=True)

    capex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Inversión Infraestructura / Activos"]["Monto"]).sum() if not df_inv.empty else 0.0
    opex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Gasto Fijo Operativo"]["Monto"]).sum() if not df_inv.empty else 0.0

    utilidad_bruta_cerdos = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0.0
    utilidad_neta_granja = utilidad_bruta_cerdos - opex_tot
    saldo_por_recuperar = max(capex_tot - utilidad_neta_granja, 0.0)
    roi_pct = (utilidad_neta_granja / capex_tot * 100.0) if capex_tot > 0 else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🏗️ Inversión Infraestructura", f"${capex_tot:,.0f}")
    k2.metric("🧾 Gastos Fijos Acumulados", f"${opex_tot:,.0f}")
    k3.metric("💵 Utilidad Neta Real Granja", f"${utilidad_neta_granja:,.0f}")
    k4.metric("📈 Recuperación Inversión (ROI)", f"{roi_pct:.1f}%", delta=f"Falta: ${saldo_por_recuperar:,.0f}" if saldo_por_recuperar > 0 else "Recuperada")

    st.divider()
    c1, c2 = st.columns(2)
    tipo_mov = c1.selectbox("Tipo de Movimiento", [
        "Inversión Infraestructura / Activos",
        "Gasto Fijo Operativo",
        "Aporte de Capital Extra"
    ])
    if tipo_mov == "Inversión Infraestructura / Activos":
        cat_mov = c2.selectbox("Categoría", ["Construcción / Ampliación Corrales", "Jaulas Parideras / Gestación", "Comederos / Bebederos / Báscula", "Compra Pie de Cría (Reproductores)", "Otra Infraestructura"])
    elif tipo_mov == "Gasto Fijo Operativo":
        cat_mov = c2.selectbox("Categoría", ["Nómina / Mano de Obra", "Energía / Agua / Gas", "Mantenimiento General", "Arriendo / Impuestos", "Otro Gasto Fijo"])
    else:
        cat_mov = c2.selectbox("Categoría", ["Inyección Capital Socios", "Fondo Reserva Alimento", "Otro Aporte"])

    desc_mov = c1.text_input("Descripción Detallada", placeholder="Ej: Construcción 2 nuevos corrales de engorde")
    monto_mov = c2.number_input("Monto ($)", min_value=0.0, value=500000.0, step=50000.0)
    resp_mov = c1.text_input("Responsable / Socio", value="Administración")
    obs_mov = c2.text_input("Observaciones", value="")

    if st.button("💾 Registrar Movimiento Financiero"):
        nuevo_m = pd.DataFrame([{
            "ID_Movimiento": f"FIN-{len(df_inv)+1:03d}",
            "Fecha": hoy.strftime("%Y-%m-%d"),
            "Tipo_Movimiento": tipo_mov,
            "Categoria": cat_mov,
            "Descripcion": desc_mov,
            "Monto": monto_mov,
            "Responsable": resp_mov,
            "Observaciones": obs_mov
        }])
        df_inv = pd.concat([df_inv, nuevo_m], ignore_index=True)
        guardar_hoja("Inversiones_Gastos", df_inv)
        st.success("Movimiento financiero registrado.")
        st.rerun()

    st.subheader("Detalle de Inversiones y Gastos Fijos")
    st.dataframe(df_inv, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 9: ANALÍTICA, METAS Y REPORTES (WORD / PDF) - SOLO ADMINISTRADOR
# ==============================================================================
elif menu == "📊 Analítica, Metas y Reportes":
    st.markdown('<div class="main-header">📊 Inteligencia Productiva, Curva de Peso y Reportes</div>', unsafe_allow_html=True)

    tab_a1, tab_a2, tab_a3 = st.tabs(["📈 Curva de Peso vs Meta 5 Meses", "🧬 Genética y Rentabilidad", "📄 Exportar Reporte Ejecutivo (Word / PDF)"])

    with tab_a1:
        st.subheader(f"Seguimiento de Crecimiento frente a Meta ({PESO_META:.0f} kg a los {EDAD_META} días)")
        if lotes_activos.empty:
            st.info("No hay lotes activos para graficar.")
        else:
            dias_eje = list(range(0, EDAD_META + 1, 15))
            peso_eje = [PESO_NAC + ((PESO_META - PESO_NAC) / max(EDAD_META, 1)) * d for d in dias_eje]

            fig_curva = go.Figure()
            fig_curva.add_trace(go.Scatter(x=dias_eje, y=peso_eje, mode="lines+markers", name="Curva Meta Ideal", line=dict(color="#1E8449", width=3, dash="dash")))

            for _, l in lotes_activos.iterrows():
                f_ini = parse_fecha(l["Fecha_Inicio"])
                edad_ini = int(to_num(pd.Series([l.get("Edad_Inicial_Dias", 0)])).iloc[0])
                dias_t = max((hoy - f_ini).days, 1) if f_ini else 1
                edad_a = edad_ini + dias_t
                p_a = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                fig_curva.add_trace(go.Scatter(
                    x=[edad_a], y=[p_a], mode="markers+text",
                    name=l["ID_Lote"], text=[f"{l['ID_Lote']} ({p_a} kg)"],
                    textposition="top center", marker=dict(size=14)
                ))

            fig_curva.update_layout(xaxis_title="Edad del Cerdo (Días)", yaxis_title="Peso Promedio (kg)", height=420)
            st.plotly_chart(fig_curva, use_container_width=True)

    with tab_a2:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Desempeño por Cruce Genético")
            if not df_repro.empty:
                df_repro["Cruce"] = df_repro["Raza_Cerda"] + " × " + df_repro["Macho_O_Semen"]
                df_repro["Nacidos_Num"] = to_num(df_repro["Nacidos_Vivos"])
                partos_df = df_repro[df_repro["Nacidos_Num"] > 0]
                if not partos_df.empty:
                    fig_gen = px.bar(partos_df.groupby("Cruce", as_index=False)["Nacidos_Num"].mean(), x="Cruce", y="Nacidos_Num", color="Cruce", title="Promedio Nacidos Vivos por Cruce")
                    st.plotly_chart(fig_gen, use_container_width=True)
        with c2:
            st.subheader("Comparativa Ventas: En Pie vs. Beneficiado")
            if not df_ventas.empty:
                df_ventas["Util_Num"] = to_num(df_ventas["Utilidad_Neta"])
                fig_v = px.bar(df_ventas.groupby("Modalidad", as_index=False)["Util_Num"].mean(), x="Modalidad", y="Util_Num", color="Modalidad", title="Utilidad Neta Promedio por Modalidad ($)")
                st.plotly_chart(fig_v, use_container_width=True)

    with tab_a3:
        st.subheader("📄 Generador de Informe Ejecutivo Integral")
        st.write("Descarga un documento completo con el estado financiero, inversiones, bandas de parto, curvas de crecimiento, sanidad e inventarios.")

        capex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Inversión Infraestructura / Activos"]["Monto"]).sum() if not df_inv.empty else 0.0
        opex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Gasto Fijo Operativo"]["Monto"]).sum() if not df_inv.empty else 0.0
        ing_tot = to_num(df_ventas["Ingreso_Bruto"]).sum() if not df_ventas.empty else 0.0
        util_lotes = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0.0
        util_granja = util_lotes - opex_tot

        def generar_word():
            doc = Document()
            titulo = doc.add_heading("INFORME EJECUTIVO DE GESTIÓN PORCINA - PORCIGESTIÓN PRO", 0)
            titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_fecha = doc.add_paragraph(f"Fecha de emisión: {hoy.strftime('%d/%m/%Y')} | Meta productiva: {PESO_META:.0f} kg a los {EDAD_META} días ({GDP_META_G:.0f} g/día)")
            p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_heading("1. Resumen Financiero e Inversiones", level=1)
            resumen_fin_txt = "\n".join([
                f"• Inversión acumulada en infraestructura y activos (CAPEX): ${capex_tot:,.0f}",
                f"• Gastos fijos operativos acumulados (OPEX): ${opex_tot:,.0f}",
                f"• Ingresos brutos por ventas registradas: ${ing_tot:,.0f}",
                f"• Utilidad operativa de lotes vendidos: ${util_lotes:,.0f}",
                f"• Utilidad neta real de la granja (descontando gastos fijos): ${util_granja:,.0f}",
                f"• Saldo de inversión por recuperar: ${max(capex_tot - util_granja, 0):,.0f}"
            ])
            doc.add_paragraph(resumen_fin_txt)

            doc.add_heading("2. Estado del Pie de Cría y Bandas", level=1)
            if not df_rep.empty:
                t_rep = doc.add_table(rows=1, cols=5)
                t_rep.style = "Table Grid"
                hdr = t_rep.rows[0].cells
                hdr[0].text = "ID"
                hdr[1].text = "Tipo / Raza"
                hdr[2].text = "Banda"
                hdr[3].text = "Estado"
                hdr[4].text = "Observaciones"
                for _, r in df_rep.iterrows():
                    row_c = t_rep.add_row().cells
                    row_c[0].text = str(r["ID_Animal"])
                    row_c[1].text = f"{r['Tipo']} ({r['Linea_Genetica']})"
                    row_c[2].text = str(r.get("Banda", ""))
                    row_c[3].text = str(r["Estado"])
                    row_c[4].text = str(r["Observaciones"])

            doc.add_heading("3. Rendimiento de Lotes Activos vs. Meta a 5 Meses", level=1)
            if not lotes_activos.empty:
                t_lot = doc.add_table(rows=1, cols=6)
                t_lot.style = "Table Grid"
                hl = t_lot.rows[0].cells
                hl[0].text = "Lote"
                hl[1].text = "Etapa"
                hl[2].text = "Cerdos"
                hl[3].text = "Peso Real (kg)"
                hl[4].text = "Costo Total ($)"
                hl[5].text = "Punto Eq. ($/kg)"
                for _, l in lotes_activos.iterrows():
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    ctot = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
                    peq = ctot / max(cant * p_act, 1)
                    rc = t_lot.add_row().cells
                    rc[0].text = str(l["ID_Lote"])
                    rc[1].text = str(l["Etapa"])
                    rc[2].text = str(cant)
                    rc[3].text = f"{p_act:.1f} kg"
                    rc[4].text = f"${ctot:,.0f}"
                    rc[5].text = f"${peq:,.0f}/kg"

            doc.add_heading("4. Alertas Operativas Vigentes", level=1)
            todas_al = alertas_criticas + alertas_preventivas
            if todas_al:
                for al in todas_al:
                    doc.add_paragraph(al.replace("**", "").replace("`", ""), style="List Bullet")
            else:
                doc.add_paragraph("Sin alertas activas a la fecha.")

            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)
            return bio

        def generar_pdf():
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "INFORME EJECUTIVO - PORCIGESTION PRO", ln=True, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(0, 8, f"Fecha: {hoy.strftime('%d/%m/%Y')} | Meta: {PESO_META:.0f} kg a los {EDAD_META} dias", ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, "1. RESUMEN FINANCIERO E INVERSIONES", ln=True)
            pdf.set_font("Arial", "", 10)
            lines_fin = [
                f"- Inversion Infraestructura (CAPEX): ${capex_tot:,.0f}",
                f"- Gastos Fijos Operativos (OPEX): ${opex_tot:,.0f}",
                f"- Ingresos Brutos Ventas: ${ing_tot:,.0f}",
                f"- Utilidad Operativa Lotes: ${util_lotes:,.0f}",
                f"- Utilidad Neta Real Granja: ${util_granja:,.0f}",
                f"- Inversion por Recuperar: ${max(capex_tot - util_granja, 0):,.0f}"
            ]
            for lf in lines_fin:
                pdf.cell(0, 6, lf, ln=True)
            pdf.ln(4)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, "2. PIE DE CRIA Y BANDAS", ln=True)
            pdf.set_font("Arial", "", 9)
            if not df_rep.empty:
                for _, r in df_rep.iterrows():
                    txt = f"{r['ID_Animal']} | {r['Linea_Genetica']} | {r.get('Banda','')} | Estado: {r['Estado']}"
                    pdf.cell(0, 6, txt.encode("latin-1", "replace").decode("latin-1"), ln=True)
            pdf.ln(4)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, "3. LOTES ACTIVOS Y COSTOS", ln=True)
            pdf.set_font("Arial", "", 9)
            if not lotes_activos.empty:
                for _, l in lotes_activos.iterrows():
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    ctot = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
                    txt = f"{l['ID_Lote']} ({l['Etapa']}): {cant} cerdos | Peso: {p_act:.1f} kg | Costo Tot: ${ctot:,.0f}"
                    pdf.cell(0, 6, txt.encode("latin-1", "replace").decode("latin-1"), ln=True)

            return bytes(pdf.output())

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="📘 Descargar Informe en Word (.docx)",
                data=generar_word(),
                file_name=f"Reporte_PorciGestion_{hoy.strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_d2:
            st.download_button(
                label="📕 Descargar Informe en PDF (.pdf)",
                data=generar_pdf(),
                file_name=f"Reporte_PorciGestion_{hoy.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# ==============================================================================
# MÓDULO 10: GESTIONAR CERDAS Y METAS - SOLO ADMINISTRADOR
# ==============================================================================
elif menu == "⚙️ Gestionar Cerdas y Metas":
    st.markdown('<div class="main-header">⚙️ Administración de Reproductores y Metas de Granja</div>', unsafe_allow_html=True)

    t_cfg1, t_cfg2 = st.tabs(["🎯 Configurar Metas de Peso y Bandas", "🐷 Agregar / Editar / Eliminar Reproductores"])

    with t_cfg1:
        st.subheader("Parámetros Técnicos de la Granja")
        c1, c2, c3 = st.columns(3)
        n_peso_meta = c1.number_input("Peso Objetivo Finalización (kg)", min_value=70.0, max_value=150.0, value=float(PESO_META), step=1.0)
        n_edad_meta = c2.number_input("Edad Objetivo (días)", min_value=100, max_value=220, value=int(EDAD_META), step=5)
        n_dias_ab = c3.number_input("Alerta Días Abiertos Máx.", min_value=3, max_value=30, value=int(MAX_DIAS_ABIERTOS), step=1)

        c4, c5, c6 = st.columns(3)
        n_int_bandas = c4.selectbox("Ritmo entre Bandas (días)", [14, 15, 21], index=[14, 15, 21].index(INTERVALO_BANDAS) if INTERVALO_BANDAS in [14, 15, 21] else 1)
        n_num_bandas = c5.number_input("Número de Bandas en Ciclo", min_value=5, max_value=12, value=int(NUM_BANDAS))
        n_cupo_banda = c6.number_input("Cupo Máx. Cerdas por Banda", min_value=1, max_value=10, value=int(CUPO_BANDA))

        if st.button("💾 Guardar Configuración de Metas"):
            nueva_cfg = pd.DataFrame([
                {"Parametro": "Peso_Meta_Kg", "Valor": str(n_peso_meta), "Descripcion": "Peso objetivo del cerdo finalizado en pie (kg)"},
                {"Parametro": "Edad_Meta_Dias", "Valor": str(n_edad_meta), "Descripcion": "Edad objetivo desde nacimiento (días)"},
                {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": str(PESO_NAC), "Descripcion": "Peso promedio al nacer (kg)"},
                {"Parametro": "Intervalo_Bandas_Dias", "Valor": str(n_int_bandas), "Descripcion": "Días entre bandas"},
                {"Parametro": "Numero_Bandas", "Valor": str(n_num_bandas), "Descripcion": "Número de bandas en ciclo completo"},
                {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": str(n_cupo_banda), "Descripcion": "Máximo de cerdas por banda"},
                {"Parametro": "Alerta_Dias_Abiertos", "Valor": str(n_dias_ab), "Descripcion": "Días máximos vacía antes de alerta"}
            ])
            guardar_hoja("Configuracion", nueva_cfg)
            st.success("Metas actualizadas correctamente.")
            st.rerun()

    with t_cfg2:
        col_add, col_del = st.columns(2)
        with col_add:
            st.subheader("➕ Agregar Nuevo Reproductor")
            id_nuevo = st.text_input("ID Animal", placeholder="Ej: M-09")
            tipo_n = st.selectbox("Tipo", ["Hembra Madre", "Verraco Recelador (Estímulo Celo)"])
            raza_n = st.selectbox("Línea Genética / Raza", ["F1 (Landrace x Yorkshire)", "Landrace", "Yorkshire", "Large White", "Duroc", "Pietrain"])
            lista_b = [f"Banda {i:02d}" for i in range(1, NUM_BANDAS + 1)]
            banda_sug = "Banda 01"
            if not df_rep.empty and "Hembra" in tipo_n:
                conteos = {b: len(df_rep[df_rep["Banda"] == b]) for b in lista_b}
                banda_sug = min(conteos, key=conteos.get)
            banda_n = st.selectbox(f"Banda Asignada (Sugerida más liviana: {banda_sug})", lista_b, index=lista_b.index(banda_sug))
            obs_n = st.text_input("Observaciones", value="Activa")

            if st.button("➕ Agregar Animal"):
                if id_nuevo and id_nuevo not in df_rep["ID_Animal"].values:
                    nueva_r = pd.DataFrame([{
                        "ID_Animal": id_nuevo.strip(),
                        "Tipo": tipo_n,
                        "Linea_Genetica": raza_n,
                        "Banda": banda_n if "Hembra" in tipo_n else "N/A",
                        "Padre_ID": "Inseminación",
                        "Madre_ID": "Externa",
                        "Partos": 0,
                        "Abortos": 0,
                        "Repeticiones": 0,
                        "Estado": "Vacía" if "Hembra" in tipo_n else "Recelador Activo",
                        "Racion_Diaria_Kg": 2.2,
                        "Costo_Acumulado_Ciclo": 0.0,
                        "Fecha_Ultimo_Cambio_Estado": hoy.strftime("%Y-%m-%d"),
                        "Observaciones": obs_n
                    }])
                    df_rep = pd.concat([df_rep, nueva_r], ignore_index=True)
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{id_nuevo} agregado en {banda_n}.")
                    st.rerun()
                else:
                    st.error("ID vacío o ya existente.")

        with col_del:
            st.subheader("🗑️ Eliminar Reproductor")
            if not df_rep.empty:
                id_del = st.selectbox("Seleccionar Animal a Eliminar", df_rep["ID_Animal"].tolist())
                if st.button("🗑️ Eliminar Definitivamente"):
                    df_rep = df_rep[df_rep["ID_Animal"] != id_del]
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{id_del} eliminado.")
                    st.rerun()

        st.divider()
        st.subheader("Tabla Directa de Reproductores")
        df_rep_edit = st.data_editor(df_rep, use_container_width=True, hide_index=True)
        if st.button("💾 Guardar Cambios de la Tabla"):
            guardar_hoja("Reproductores", df_rep_edit)
            st.success("Cambios guardados en Google Sheets.")
            st.rerun()