import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
from streamlit_gsheets import GSheetsConnection
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

st.set_page_config(
    page_title="PorciGestión Pro | Control Integral Porcino",
    page_icon="🐽",
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

# ==============================================================================
# PROGRAMA NUTRICIONAL Y CURVA DE CRECIMIENTO FIJA (21d DESTETE -> 120kg / 153d)
# ==============================================================================
PROGRAMA_FASES = [
    {"fase": "Fase 0", "edad_ini": 21, "edad_fin": 28, "peso_ini": 6.2, "peso_fin": 7.8, "kg_cerdo": 1.2, "siguiente": "Fase 1"},
    {"fase": "Fase 1", "edad_ini": 28, "edad_fin": 35, "peso_ini": 7.8, "peso_fin": 10.2, "kg_cerdo": 2.0, "siguiente": "Fase 2"},
    {"fase": "Fase 2", "edad_ini": 35, "edad_fin": 45, "peso_ini": 10.2, "peso_fin": 14.5, "kg_cerdo": 5.0, "siguiente": "Iniciación"},
    {"fase": "Iniciación", "edad_ini": 45, "edad_fin": 70, "peso_ini": 14.5, "peso_fin": 30.0, "kg_cerdo": 21.0, "siguiente": "Prelevante"},
    {"fase": "Prelevante", "edad_ini": 70, "edad_fin": 90, "peso_ini": 30.0, "peso_fin": 46.0, "kg_cerdo": 27.0, "siguiente": "Levante"},
    {"fase": "Levante", "edad_ini": 90, "edad_fin": 112, "peso_ini": 46.0, "peso_fin": 68.0, "kg_cerdo": 40.0, "siguiente": "Engorde"},
    {"fase": "Engorde", "edad_ini": 112, "edad_fin": 132, "peso_ini": 68.0, "peso_fin": 92.0, "kg_cerdo": 40.0, "siguiente": "Finalizador"},
    {"fase": "Finalizador", "edad_ini": 132, "edad_fin": 153, "peso_ini": 92.0, "peso_fin": 120.0, "kg_cerdo": 80.0, "siguiente": "Venta / Mercado (120 kg)"}
]

LISTA_CONCENTRADOS = [
    "Fase 0", "Fase 1", "Fase 2", "Iniciación",
    "Prelevante", "Levante", "Engorde", "Finalizador",
    "Gestación", "Lactancia"
]

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
        "Fecha", "Linea_Negocio", "ID_Origen", "Modalidad", "Cantidad_Animales_O_Unidades",
        "Peso_Pie_O_Kilos", "Peso_Carne_Canal_Kg", "Costo_Sacrificio_O_Flete",
        "Ingreso_Bruto", "Costo_Produccion_Asignado", "Utilidad_Neta"
    ],
    "Inversiones_Gastos": [
        "ID_Movimiento", "Fecha", "Tipo_Movimiento", "Origen_Fondos", "Categoria",
        "Descripcion", "Monto", "Responsable", "Observaciones"
    ],
    "Configuracion": [
        "Parametro", "Valor", "Descripcion"
    ]
}

CONFIG_DEFAULT = pd.DataFrame([
    {"Parametro": "Peso_Meta_Kg", "Valor": "120.0", "Descripcion": "Peso objetivo del cerdo finalizado en pie (kg)"},
    {"Parametro": "Edad_Meta_Dias", "Valor": "153", "Descripcion": "Edad objetivo desde nacimiento para alcanzar 120 kg (150-155 días)"},
    {"Parametro": "Edad_Destete_Dias", "Valor": "21", "Descripcion": "Edad estándar de destete de lechones (días)"},
    {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": "1.4", "Descripcion": "Peso promedio al nacer (kg)"},
    {"Parametro": "Peso_Destete_21d_Kg", "Valor": "6.2", "Descripcion": "Peso objetivo al destete a los 21 días (kg)"},
    {"Parametro": "Margen_Venta_Cuido_Pct", "Valor": "18.0", "Descripcion": "Porcentaje de margen comercial por defecto para venta de concentrado (%)"},
    {"Parametro": "Intervalo_Bandas_Dias", "Valor": "15", "Descripcion": "Días entre inseminación de una banda y la siguiente (14, 15 o 21)"},
    {"Parametro": "Numero_Bandas", "Valor": "10", "Descripcion": "Total de bandas en el ciclo completo"},
    {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": "2", "Descripcion": "Máximo de cerdas a inseminar por banda"},
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

def peso_esperado_por_edad(edad_dias, peso_nac=1.4, peso_dest=6.2, edad_dest=21):
    if edad_dias <= 0:
        return peso_nac
    if edad_dias <= edad_dest:
        return peso_nac + ((peso_dest - peso_nac) / max(edad_dest, 1)) * edad_dias
    for f in PROGRAMA_FASES:
        if edad_dias <= f["edad_fin"]:
            span = max(f["edad_fin"] - f["edad_ini"], 1)
            prog = (edad_dias - f["edad_ini"]) / span
            return f["peso_ini"] + (f["peso_fin"] - f["peso_ini"]) * prog
    return 120.0 + (edad_dias - 153) * 0.85

def fase_por_edad(edad_dias, edad_dest=21):
    if edad_dias < edad_dest:
        return {"fase": "Lactancia", "edad_ini": 0, "edad_fin": edad_dest, "peso_ini": 1.4, "peso_fin": 6.2, "kg_cerdo": 0.0, "siguiente": "Fase 0"}
    for f in PROGRAMA_FASES:
        if edad_dias <= f["edad_fin"]:
            return f
    return PROGRAMA_FASES[-1]

def evaluar_fase_lote(lote_row, df_cons, hoy_date, edad_dest=21):
    lote_id = str(lote_row["ID_Lote"])
    cant = int(to_num(pd.Series([lote_row["Cant_Actual"]])).iloc[0])
    f_ini = parse_fecha(lote_row["Fecha_Inicio"])
    edad_ini = int(to_num(pd.Series([lote_row.get("Edad_Inicial_Dias", 0)])).iloc[0])
    dias_t = max((hoy_date - f_ini).days, 1) if f_ini else 1
    edad_act = edad_ini + dias_t

    etapa_actual = str(lote_row.get("Etapa", "")).strip()
    nombres_fases = [f["fase"] for f in PROGRAMA_FASES]

    if etapa_actual in nombres_fases:
        info_f = next(f for f in PROGRAMA_FASES if f["fase"] == etapa_actual)
    else:
        info_f = fase_por_edad(edad_act, edad_dest)

    kg_meta_fase_lote = round(info_f["kg_cerdo"] * max(cant, 1), 1)
    kg_consumidos_fase = 0.0
    if not df_cons.empty and info_f["kg_cerdo"] > 0:
        sub_c = df_cons[(df_cons["ID_Destino"] == lote_id) & (df_cons["Tipo_Concentrado"] == info_f["fase"])]
        kg_consumidos_fase = float(to_num(sub_c["Kilos_Servidos"]).sum())

    return {
        "edad_actual": edad_act,
        "fase_actual": info_f["fase"],
        "siguiente_fase": info_f["siguiente"],
        "kg_por_cerdo_fase": info_f["kg_cerdo"],
        "kg_meta_lote": kg_meta_fase_lote,
        "kg_consumidos_fase": round(kg_consumidos_fase, 1),
        "kg_restantes_fase": round(max(kg_meta_fase_lote - kg_consumidos_fase, 0.0), 1),
        "edad_fin_fase": info_f["edad_fin"],
        "peso_objetivo_fin_fase": info_f["peso_fin"]
    }

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
PESO_META = get_config(df_cfg, "Peso_Meta_Kg", 120.0)
EDAD_META = int(get_config(df_cfg, "Edad_Meta_Dias", 153))
EDAD_DESTETE = int(get_config(df_cfg, "Edad_Destete_Dias", 21))
PESO_NAC = get_config(df_cfg, "Peso_Nacimiento_Prom_Kg", 1.4)
PESO_DEST_21D = get_config(df_cfg, "Peso_Destete_21d_Kg", 6.2)
MARGEN_CUIDO_DEFAULT = get_config(df_cfg, "Margen_Venta_Cuido_Pct", 18.0)
INTERVALO_BANDAS = int(get_config(df_cfg, "Intervalo_Bandas_Dias", 15))
NUM_BANDAS = int(get_config(df_cfg, "Numero_Bandas", 10))
CUPO_BANDA = int(get_config(df_cfg, "Cupo_Max_Cerdas_Banda", 2))
MAX_DIAS_ABIERTOS = int(get_config(df_cfg, "Alerta_Dias_Abiertos", 7))
GDP_META_G = ((PESO_META - PESO_NAC) / max(EDAD_META, 1)) * 1000.0

# ==============================================================================
# CONTROL DE ACCESO (OPERARIO SIN CLAVE / ADMINISTRADOR CON CLAVE)
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "rol_usuario" not in st.session_state:
    st.session_state["rol_usuario"] = None

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
    with col_l2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div style="font-size: 135px; line-height: 1.05; margin-bottom: 10px;">🐽</div>', unsafe_allow_html=True)
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

# Barra lateral según perfil
with st.sidebar:
    st.markdown('<div style="font-size: 115px; line-height: 1.05; margin-bottom: 5px;">🐽</div>', unsafe_allow_html=True)
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
            "💼 Salud Financiera y Reinversión",
            "📅 Escalonamiento y Bandas",
            "🧬 Reproducción (IA y Partos)",
            "🐖 Lotes, Curva y Fases",
            "🌾 Bodega, Arqueo y Botiquín",
            "💉 Sanidad y Mortalidad",
            "💰 Ventas (Marranos / Cuido)",
            "📊 Analítica, Metas y Reportes",
            "⚙️ Gestionar Cerdas y Metas"
        ]
    else:
        opciones_menu = [
            "🚨 Tablero de Alertas",
            "🧬 Reproducción (IA y Partos)",
            "🐖 Lotes, Curva y Fases",
            "🌾 Bodega, Arqueo y Botiquín",
            "💉 Sanidad y Mortalidad"
        ]

    menu = st.radio("Módulos del Sistema", opciones_menu)

# ==============================================================================
# CÁLCULO CENTRAL DE ALERTAS OPERATIVAS Y NUTRICIONALES
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
                alertas_preventivas.append(f"🍼 **Destete Listo a los {EDAD_DESTETE} días ({cerda} / Lote {r.get('Nombre_Lote_Asignado','')}):** Fecha objetivo `{f_dest}`. Destetar e iniciar **Fase 0** (1.2 kg/cerdo).")

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
        ev_f = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE)
        edad_actual = ev_f["edad_actual"]
        peso_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
        peso_esperado_hoy = peso_esperado_por_edad(edad_actual, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE)
        dif_peso = peso_act - peso_esperado_hoy

        # Alerta 1: Cambio de fase nutricional por consumo acumulado
        if ev_f["kg_meta_lote"] > 0:
            if ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"]:
                alertas_criticas.append(
                    f"🌾 **Cambio de Fase Nutricional ({l['ID_Lote']}):** Ya completó los **{ev_f['kg_consumidos_fase']:.1f} kg de {ev_f['fase_actual']}** "
                    f"(meta: {ev_f['kg_meta_lote']:.1f} kg). **Iniciar inmediatamente con {ev_f['siguiente_fase']}** y actualizar la etapa del lote."
                )
            elif ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"] * 0.88:
                alertas_preventivas.append(
                    f"📦 **Próximo Cambio de Fase ({l['ID_Lote']}):** Lleva **{ev_f['kg_consumidos_fase']:.1f} / {ev_f['kg_meta_lote']:.1f} kg de {ev_f['fase_actual']}** "
                    f"(solo faltan **{ev_f['kg_restantes_fase']:.1f} kg**). Alistando bultos de **{ev_f['siguiente_fase']}**."
                )
            elif edad_actual > ev_f["edad_fin_fase"]:
                alertas_preventivas.append(
                    f"⏰ **Transición por Edad ({l['ID_Lote']} - {edad_actual} días):** Por edad ya superó los {ev_f['edad_fin_fase']} días de **{ev_f['fase_actual']}** "
                    f"(solo lleva {ev_f['kg_consumidos_fase']:.1f}/{ev_f['kg_meta_lote']:.1f} kg registrados). Verificar consumo y pasar a **{ev_f['siguiente_fase']}**."
                )

        # Alerta 2: Curva de crecimiento frente a meta 120 kg
        if edad_actual >= 28 and dif_peso < -3.5:
            alertas_preventivas.append(
                f"📉 **Bajo Ritmo de Crecimiento ({l['ID_Lote']}):** Pesa **{peso_act:.1f} kg** a los {edad_actual} días, cuando la curva a 120 kg exige **{peso_esperado_hoy:.1f} kg** ({dif_peso:+.1f} kg)."
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
# MÓDULO 1: TABLERO DE ALERTAS OPERATIVAS Y NUTRICIONALES
# ==============================================================================
if menu == "🚨 Tablero de Alertas":
    st.markdown('<div class="main-header">🚨 Centro de Control y Alertas Operativas</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Fecha del sistema: <b>{hoy.strftime("%d/%m/%Y")}</b> | Destete: <b>{EDAD_DESTETE} días</b> | Meta: <b>{PESO_META:.0f} kg a los {EDAD_META} días ({GDP_META_G:.0f} g/día)</b></div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    total_madres = len(df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (~df_rep["Estado"].str.contains("Descartada", na=False))]) if not df_rep.empty else 0
    total_cerdos_engorde = int(to_num(lotes_activos["Cant_Actual"]).sum()) if not lotes_activos.empty else 0
    utilidad_total_ventas = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0

    c1.metric("🐷 Cerdas Madres Activas", total_madres)
    c2.metric("🐖 Cerdos en Crecimiento", total_cerdos_engorde)
    if es_admin:
        c3.metric("💵 Utilidad Bruta Ventas", f"${utilidad_total_ventas:,.0f}")
    else:
        c3.metric("🌾 Lotes Activos en Galpón", len(lotes_activos))
    c4.metric("🔔 Alertas Operativas", len(alertas_criticas) + len(alertas_preventivas))

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
        st.subheader("🟡 Alertas Preventivas y Nutricionales")
        if alertas_preventivas:
            for a in alertas_preventivas:
                st.warning(a)
        else:
            st.info("✅ Todos los lotes y stocks están dentro de meta.")

# ==============================================================================
# MÓDULO 2: SALUD FINANCIERA, REINVERSIONES Y ALERTAS DE CAPITAL (SOLO ADMIN)
# ==============================================================================
elif menu == "💼 Salud Financiera y Reinversión":
    st.markdown('<div class="main-header">💼 Salud Financiera, Reinversión y Estabilidad</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Control independiente de líneas de ingreso, capital reinvertido, punto de estabilidad y excedentes libres para retiro.</div>', unsafe_allow_html=True)

    # Desglose de ingresos por línea de negocio
    ing_marranos, util_marranos = 0.0, 0.0
    ing_descarte, util_descarte = 0.0, 0.0
    ing_cuido, util_cuido = 0.0, 0.0
    ing_otros, util_otros = 0.0, 0.0

    if not df_ventas.empty:
        for _, v in df_ventas.iterrows():
            linea = str(v.get("Linea_Negocio", "Venta Cerdos Cebados"))
            ib = float(to_num(pd.Series([v["Ingreso_Bruto"]])).iloc[0])
            un = float(to_num(pd.Series([v["Utilidad_Neta"]])).iloc[0])
            if "Concentrado" in linea or "Cuido" in linea:
                ing_cuido += ib
                util_cuido += un
            elif "Descarte" in linea:
                ing_descarte += ib
                util_descarte += un
            elif "Cebados" in linea or "Pie" in linea or "Canal" in linea:
                ing_marranos += ib
                util_marranos += un
            else:
                ing_otros += ib
                util_otros += un

    util_total_negocio = util_marranos + util_descarte + util_cuido + util_otros

    # Movimientos de capital y gastos
    aportes_socio = 0.0
    capex_total = 0.0
    gastos_fijos_tot = 0.0
    reinversion_utilidades = 0.0
    retiros_ganancia = 0.0

    if not df_inv.empty:
        for _, m in df_inv.iterrows():
            t_mov = str(m["Tipo_Movimiento"])
            orig = str(m.get("Origen_Fondos", ""))
            monto = float(to_num(pd.Series([m["Monto"]])).iloc[0])

            if t_mov == "Aporte de Capital (Socio)":
                aportes_socio += monto
            elif t_mov == "Inversión Infraestructura / Activos":
                capex_total += monto
                if "Reinversión" in orig:
                    reinversion_utilidades += monto
            elif t_mov == "Reinversión en Inventario / Operaciones":
                reinversion_utilidades += monto
            elif t_mov == "Gasto Fijo Operativo":
                gastos_fijos_tot += monto
            elif t_mov == "Retiro de Utilidades (Ganancia Socio)":
                retiros_ganancia += monto

    utilidad_neta_real = util_total_negocio - gastos_fijos_tot
    caja_libre_acumulada = utilidad_neta_real - reinversion_utilidades - retiros_ganancia
    # Capital de trabajo sugerido de reserva (alimento para 1 lote de 12 cerdos completo ~2.600 kg)
    reserva_capital_trabajo = 4500000.0
    meta_estabilidad = capex_total + reserva_capital_trabajo
    avance_estabilidad_pct = (utilidad_neta_real / max(meta_estabilidad, 1.0)) * 100.0 if meta_estabilidad > 0 else 100.0
    excedente_retirable_hoy = max(caja_libre_acumulada - reserva_capital_trabajo, 0.0) if utilidad_neta_real >= capex_total else 0.0

    # --- ALERTAS FINANCIERAS INDEPENDIENTES ---
    st.subheader("🧭 Diagnóstico y Alertas Financieras de Estabilidad")
    if utilidad_neta_real < capex_total:
        faltante = capex_total - utilidad_neta_real
        st.error(
            f"🔴 **ETAPA DE CAPITALIZACIÓN (REINVERTIR 100 %):** La utilidad neta acumulada (**${utilidad_neta_real:,.0f}**) aún no cubre la inversión inicial en infraestructura/activos (**${capex_total:,.0f}**). "
            f"Faltan **${faltante:,.0f}** para recuperar el capital invertido. **Recomendación:** No retirar ganancias todavía; reinvertir toda utilidad en concentrado y crecimiento."
        )
    elif utilidad_neta_real < meta_estabilidad:
        st.warning(
            f"🟡 **INVERSIÓN RECUPERADA — CONSTRUYENDO COLCHÓN OPERATIVO:** Ya recuperaste el 100 % de la infraestructura (**${capex_total:,.0f}**), pero aún estás formando la reserva de capital de trabajo (**${reserva_capital_trabajo:,.0f}** para concentrado). "
            f"**Recomendación:** Reinvertir al menos el **70 %** de las utilidades y reservar el **30 %**."
        )
    else:
        st.success(
            f"🟢 **PUNTO DE ESTABILIDAD FINANCIERA ALCANZADO ({avance_estabilidad_pct:.0f} %):** La granja ya pagó su infraestructura, tiene cubierto su capital de trabajo y genera flujo positivo sostenible. "
            f"**Excedente libre disponible hoy para retirar como ganancia sin descapitalizar la granja: ${excedente_retirable_hoy:,.0f}**."
        )

    st.divider()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Utilidad Neta Real Granja", f"${utilidad_neta_real:,.0f}", help="Utilidad de todas las ventas menos gastos fijos")
    k2.metric("🔄 Utilidades Reinvertidas", f"${reinversion_utilidades:,.0f}")
    k3.metric("🏦 Caja Libre Disponible", f"${caja_libre_acumulada:,.0f}")
    k4.metric("💸 Retiros Socio Acumulados", f"${retiros_ganancia:,.0f}")

    st.subheader("📊 Rentabilidad por Línea de Negocio")
    df_lineas = pd.DataFrame([
        {"Línea de Ingreso": "🐖 Venta Cerdos Cebados (Pie / Canal)", "Ingresos Brutos ($)": ing_marranos, "Utilidad Neta ($)": util_marranos},
        {"Línea de Ingreso": "🐷 Venta Cerdas de Descarte", "Ingresos Brutos ($)": ing_descarte, "Utilidad Neta ($)": util_descarte},
        {"Línea de Ingreso": "🌾 Distribución / Venta de Cuido", "Ingresos Brutos ($)": ing_cuido, "Utilidad Neta ($)": util_cuido},
        {"Línea de Ingreso": "🍼 Venta Lechones / Otros", "Ingresos Brutos ($)": ing_otros, "Utilidad Neta ($)": util_otros}
    ])
    c_tab, c_graf = st.columns([1.2, 1])
    with c_tab:
        st.dataframe(df_lineas, use_container_width=True, hide_index=True)
    with c_graf:
        if df_lineas["Utilidad Neta ($)"].sum() > 0:
            fig_l = px.pie(df_lineas[df_lineas["Utilidad Neta ($)"] > 0], values="Utilidad Neta ($)", names="Línea de Ingreso", title="Participación en Utilidad Neta")
            fig_l.update_layout(height=280, margin=dict(t=40, b=10, l=10, r=10))
            st.plotly_chart(fig_l, use_container_width=True)

    st.divider()
    st.subheader("📝 Registrar Movimiento de Capital, Reinversión, Gasto Fijo o Retiro")
    c1, c2 = st.columns(2)
    tipo_mov = c1.selectbox("Tipo de Movimiento", [
        "Inversión Infraestructura / Activos",
        "Reinversión en Inventario / Operaciones",
        "Gasto Fijo Operativo",
        "Retiro de Utilidades (Ganancia Socio)",
        "Aporte de Capital (Socio)"
    ])
    origen_f = c2.selectbox("Origen de los Fondos", [
        "Reinversión de Utilidades (Caja Granja)",
        "Aporte Directo Socio / Capital Propio",
        "Crédito / Externo"
    ])

    if tipo_mov == "Inversión Infraestructura / Activos":
        cat_mov = c1.selectbox("Categoría", ["Construcción / Ampliación Corrales", "Jaulas Parideras / Gestación", "Comederos / Bebederos / Báscula", "Compra Pie de Cría", "Otra Infraestructura"])
    elif tipo_mov == "Reinversión en Inventario / Operaciones":
        cat_mov = c1.selectbox("Categoría", ["Compra Lote Concentrado para Reventa", "Capital Trabajo Alimento Cerdos", "Compra Nuevas Cerdas Reemplazo"])
    elif tipo_mov == "Gasto Fijo Operativo":
        cat_mov = c1.selectbox("Categoría", ["Nómina / Mano de Obra", "Energía / Agua / Gas", "Mantenimiento General", "Arriendo / Impuestos / Transporte Fijo"])
    elif tipo_mov == "Retiro de Utilidades (Ganancia Socio)":
        cat_mov = c1.selectbox("Categoría", ["Distribución Utilidades Socios", "Retiro Mensual Administrador"])
    else:
        cat_mov = c1.selectbox("Categoría", ["Inyección Capital Inicial", "Aporte Extraordinario Socio"])

    monto_mov = c2.number_input("Monto ($)", min_value=0.0, value=500000.0, step=50000.0)
    desc_mov = c1.text_input("Descripción Detallada", placeholder="Ej: Reinversión de utilidad venta Lote 01 en compra de 20 bultos")
    resp_mov = c2.text_input("Responsable / Socio", value="Administración")
    obs_mov = st.text_input("Observaciones", value="")

    if st.button("💾 Guardar Movimiento Financiero"):
        nuevo_m = pd.DataFrame([{
            "ID_Movimiento": f"FIN-{len(df_inv)+1:03d}",
            "Fecha": hoy.strftime("%Y-%m-%d"),
            "Tipo_Movimiento": tipo_mov,
            "Origen_Fondos": origen_f,
            "Categoria": cat_mov,
            "Descripcion": desc_mov,
            "Monto": monto_mov,
            "Responsable": resp_mov,
            "Observaciones": obs_mov
        }])
        df_inv = pd.concat([df_inv, nuevo_m], ignore_index=True)
        guardar_hoja("Inversiones_Gastos", df_inv)
        st.success("Movimiento financiero registrado correctamente.")
        st.rerun()

    st.subheader("Historial de Capital, Reinversiones y Gastos Fijos")
    st.dataframe(df_inv, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 3: ESCALONAMIENTO Y BANDAS DE PARTO
# ==============================================================================
elif menu == "📅 Escalonamiento y Bandas":
    st.markdown('<div class="main-header">📅 Planificador de Escalonamiento por Bandas</div>', unsafe_allow_html=True)
    ciclo_total = 114 + EDAD_DESTETE + 5
    st.markdown(
        f'<div class="sub-header">Ciclo biológico: <b>114 días gestación + {EDAD_DESTETE} días lactancia + 5 días destete-celo = {ciclo_total} días</b>. '
        f'Configuración: <b>{NUM_BANDAS} bandas cada {INTERVALO_BANDAS} días</b> (máx. {CUPO_BANDA} cerdas/banda).</div>',
        unsafe_allow_html=True
    )

    hembras = df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (~df_rep["Estado"].str.contains("Descartada", na=False))].copy() if not df_rep.empty else pd.DataFrame()
    lista_bandas = [f"Banda {i:02d}" for i in range(1, NUM_BANDAS + 1)]

    tab1, tab2, tab3 = st.tabs(["📊 Distribución Actual de Bandas", "🗓️ Cronograma Maestro de Rotación", "➕ Reorganizar Cerdas"])

    with tab1:
        if hembras.empty:
            st.info("No hay hembras activas registradas.")
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
            st.dataframe(pd.DataFrame(resumen_bandas), use_container_width=True, hide_index=True)

    with tab2:
        fecha_inicio_ciclo = st.date_input("Fecha de inicio de la Banda 01 (referencia)", value=hoy)
        filas_cron = []
        for idx, b in enumerate(lista_bandas):
            f_ins = fecha_inicio_ciclo + timedelta(days=idx * INTERVALO_BANDAS)
            f_par = f_ins + timedelta(days=114)
            f_des = f_par + timedelta(days=EDAD_DESTETE)
            f_reins = f_des + timedelta(days=5)
            f_venta = f_par + timedelta(days=EDAD_META)
            cerdas_b = hembras[hembras["Banda"] == b]["ID_Animal"].tolist() if not hembras.empty else []
            filas_cron.append({
                "Banda": b,
                "Cerdas": ", ".join(cerdas_b) if cerdas_b else "(Cupo libre)",
                "Inseminación": f_ins.strftime("%Y-%m-%d"),
                "Parto (114d)": f_par.strftime("%Y-%m-%d"),
                f"Destete ({EDAD_DESTETE}d)": f_des.strftime("%Y-%m-%d"),
                "Nueva Inseminación (+5d)": f_reins.strftime("%Y-%m-%d"),
                f"Salida Venta 120kg ({EDAD_META}d)": f_venta.strftime("%Y-%m-%d")
            })
        st.dataframe(pd.DataFrame(filas_cron), use_container_width=True, hide_index=True)

    with tab3:
        if hembras.empty:
            st.info("Agrega cerdas primero.")
        else:
            col_izq, col_der = st.columns(2)
            with col_izq:
                cerda_sel = st.selectbox("Seleccionar Cerda", hembras["ID_Animal"].tolist())
                banda_nueva = st.selectbox("Asignar a Banda", lista_bandas)
                if st.button("💾 Guardar Banda de esta Cerda"):
                    df_rep.loc[df_rep["ID_Animal"] == cerda_sel, "Banda"] = banda_nueva
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{cerda_sel} asignada a {banda_nueva}.")
                    st.rerun()

            with col_der:
                st.write("Distribuye automáticamente todas las cerdas activas entre las bandas sin sobrepasar el cupo por quincena.")
                if st.button("🪄 Distribuir Todas las Cerdas Automáticamente"):
                    idx_h = df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (~df_rep["Estado"].str.contains("Descartada", na=False))].index.tolist()
                    for i, row_idx in enumerate(idx_h):
                        df_rep.at[row_idx, "Banda"] = lista_bandas[i % NUM_BANDAS]
                    guardar_hoja("Reproductores", df_rep)
                    st.success("Cerdas distribuidas equilibradamente.")
                    st.rerun()

# ==============================================================================
# MÓDULO 4: REPRODUCCIÓN (IA Y PARTOS)
# ==============================================================================
elif menu == "🧬 Reproducción (IA y Partos)":
    st.markdown('<div class="main-header">🧬 Gestión Reproductiva (Inseminación Artificial)</div>', unsafe_allow_html=True)

    hembras = df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (~df_rep["Estado"].str.contains("Descartada", na=False))] if not df_rep.empty else pd.DataFrame()
    t1, t2, t3, t4 = st.tabs(["1️⃣ Registrar Inseminación", "2️⃣ Confirmar Preñez / Falla", "3️⃣ Registrar Parto", f"4️⃣ Registrar Destete ({EDAD_DESTETE}d)"])

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
            costo_ia = c1.number_input("Costo Total Servicio / Pajillas ($)", min_value=0.0, value=45000.0, step=5000.0) if es_admin else 45000.0

            f_parto_est = f_serv + timedelta(days=114)
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
                    "Fecha_Prob_Destete": (f_parto_est + timedelta(days=EDAD_DESTETE)).strftime("%Y-%m-%d"),
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
                st.success(f"Servicio {nuevo_id} registrado para {cerda_ia}. Parto probable: {f_parto_est}.")
                st.rerun()

    with t2:
        pendientes = df_repro[df_repro["Estado_Prenez"].isin(["Pendiente Eco", "Preñada"])] if not df_repro.empty else pd.DataFrame()
        if pendientes.empty:
            st.info("No hay servicios pendientes.")
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
                df_repro.loc[df_repro["ID_Evento"] == id_ev_p, "Fecha_Prob_Destete"] = (f_parto_r + timedelta(days=EDAD_DESTETE)).strftime("%Y-%m-%d")

                nuevo_lote = pd.DataFrame([{
                    "ID_Lote": nombre_lote,
                    "Madre_Origen": cerda_id,
                    "Raza_Cruce": cruce_txt,
                    "Corral": corral_parto,
                    "Etapa": "Lactancia",
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
                st.success(f"Parto registrado. Lote {nombre_lote} creado (destete programado a los {EDAD_DESTETE} días).")
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
            peso_tot_dest = c2.number_input("Peso Total Camada Destetada (kg)", min_value=1.0, value=float(round(dest_reales * PESO_DEST_21D, 1)))
            corral_dest = st.text_input("Corral de Precebo (Fase 0)", value="Precebo-1")

            if st.button("💾 Confirmar Destete e Iniciar Fase 0 (1.2 kg/cerdo)"):
                cerda_id = fila_d["ID_Cerda"]
                lote_id = fila_d["Nombre_Lote_Asignado"]
                costo_lact = float(to_num(df_rep.loc[df_rep["ID_Animal"] == cerda_id, "Costo_Acumulado_Ciclo"]).iloc[0])

                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Estado_Prenez"] = "Destetada"
                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Destetados_Reales"] = dest_reales
                df_repro.loc[df_repro["ID_Evento"] == id_ev_d, "Peso_Destete_Total_Kg"] = peso_tot_dest

                df_lotes.loc[df_lotes["ID_Lote"] == lote_id, "Etapa"] = "Fase 0"
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
                st.success(f"Destete completado. {lote_id} entra a **Fase 0** (Meta: {dest_reales * 1.2:.1f} kg totales).")
                st.rerun()

    st.divider()
    st.subheader("Historial Reproductivo")
    st.dataframe(df_repro, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 5: LOTES, CURVA DE CRECIMIENTO Y FASES NUTRICIONALES
# ==============================================================================
elif menu == "🐖 Lotes, Curva y Fases":
    st.markdown('<div class="main-header">🐖 Control Productivo, Curva 120 kg y Fases Nutricionales</div>', unsafe_allow_html=True)

    with st.expander("📋 Ver Tabla Maestra de Alimentación y Pesos Esperados (216.2 kg/cerdo -> 120 kg a los 153 días)"):
        df_prog_vis = pd.DataFrame(PROGRAMA_FASES)
        df_prog_vis["Acumulado_Kg"] = df_prog_vis["kg_cerdo"].cumsum()
        df_prog_vis.columns = ["Fase", "Edad Inicio (d)", "Edad Fin (d)", "Peso Inicio (kg)", "Peso Fin (kg)", "Kg / Cerdo", "Siguiente Fase", "Kg Acumulados"]
        st.dataframe(df_prog_vis, use_container_width=True, hide_index=True)

    if lotes_activos.empty:
        st.info("No hay lotes activos.")
    else:
        filas_ind = []
        for _, l in lotes_activos.iterrows():
            ev_f = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE)
            edad_dias = ev_f["edad_actual"]
            cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
            p_ini = float(to_num(pd.Series([l["Peso_Prom_Inicial_Kg"]])).iloc[0])
            p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
            kg_alim = float(to_num(pd.Series([l["Kg_Alimento_Consumido"]])).iloc[0])

            c_ini = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0])
            c_ali = float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0])
            c_san = float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
            c_tot = c_ini + c_ali + c_san

            dias_trans = max(edad_dias - int(to_num(pd.Series([l.get("Edad_Inicial_Dias", 0)])).iloc[0]), 1)
            gdp_real_g = ((p_act - p_ini) / dias_trans) * 1000.0
            peso_esp_hoy = peso_esperado_por_edad(edad_dias, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE)
            cumpl_pct = (p_act / max(peso_esp_hoy, 0.1)) * 100.0

            kg_ganados_tot = max((p_act - p_ini) * cant, 0.1)
            ca = kg_alim / kg_ganados_tot if kg_alim > 0 else 0.0
            punto_eq = c_tot / max(p_act * cant, 1.0)

            estado_fase_txt = (
                f"✅ Completada -> Pasar a {ev_f['siguiente_fase']}"
                if (ev_f["kg_meta_lote"] > 0 and ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"])
                else f"{ev_f['kg_consumidos_fase']:.1f} / {ev_f['kg_meta_lote']:.1f} kg"
            )

            fila_data = {
                "Lote": l["ID_Lote"],
                "Cerdos": cant,
                "Edad (d)": edad_dias,
                "Fase Actual": ev_f["fase_actual"],
                "Progreso Alimento Fase": estado_fase_txt,
                "Peso Real (kg)": round(p_act, 1),
                "Peso Curva Hoy (kg)": round(peso_esp_hoy, 1),
                "Cumpl. (%)": round(cumpl_pct, 1),
                "GDP (g/d)": round(gdp_real_g, 0),
                "Conv. Alim.": round(ca, 2)
            }
            if es_admin:
                fila_data["Costo Total ($)"] = round(c_tot, 0)
                fila_data["Punto Eq. ($/kg)"] = round(punto_eq, 0)

            filas_ind.append(fila_data)

        st.subheader("📊 Seguimiento de Lotes Activos y Avance de Fase")
        st.dataframe(pd.DataFrame(filas_ind), use_container_width=True, hide_index=True)

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("⚖️ Actualizar Pesaje o Avanzar de Fase")
            lote_edit = st.selectbox("Seleccionar Lote", lotes_activos["ID_Lote"].tolist())
            fila_le = lotes_activos[lotes_activos["ID_Lote"] == lote_edit].iloc[0]
            ev_le = evaluar_fase_lote(fila_le, df_consumo, hoy, EDAD_DESTETE)
            st.caption(f"Fase actual: **{ev_le['fase_actual']}** | Siguiente fase sugerida: **{ev_le['siguiente_fase']}**")

            nuevo_peso = st.number_input("Nuevo Peso Promedio Actual (kg)", min_value=1.0, value=float(to_num(pd.Series([fila_le["Peso_Prom_Actual_Kg"]])).iloc[0]), step=0.5)
            opciones_etapa = ["Lactancia"] + [f["fase"] for f in PROGRAMA_FASES]
            idx_et = opciones_etapa.index(fila_le["Etapa"]) if fila_le["Etapa"] in opciones_etapa else 1
            nueva_etapa = st.selectbox("Fase / Etapa Nutricional", opciones_etapa, index=idx_et)
            nuevo_corral = st.text_input("Corral", value=str(fila_le["Corral"]))

            if st.button("💾 Guardar Pesaje / Cambio de Fase"):
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Peso_Prom_Actual_Kg"] = nuevo_peso
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Etapa"] = nueva_etapa
                df_lotes.loc[df_lotes["ID_Lote"] == lote_edit, "Corral"] = nuevo_corral
                guardar_hoja("Lotes_Engorde", df_lotes)
                st.success(f"Lote {lote_edit} actualizado a {nueva_etapa} ({nuevo_peso} kg).")
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
# MÓDULO 6: BODEGA, VERIFICADOR DE ARQUEO POR BULTOS Y BOTIQUÍN
# ==============================================================================
elif menu == "🌾 Bodega, Arqueo y Botiquín":
    st.markdown('<div class="main-header">🌾 Bodega, Verificador de Bultos y Botiquín</div>', unsafe_allow_html=True)

    tab_b1, tab_b2, tab_b3 = st.tabs([
        "🔍 Verificador de Inventario Físico (Bultos)",
        "🍽️ Servir Alimento a Cerdos",
        "📦 Ingresar Compra a Bodega"
    ])

    # --- PESTAÑA 1: BOTÓN / VERIFICADOR EXPLÍCITO DE BULTOS EN BODEGA ---
    with tab_b1:
        st.subheader("🔍 Consulta Rápida y Arqueo Físico de Bodega")
        st.write("Selecciona cualquier referencia de concentrado o medicamento para ver cuántos bultos/unidades exactas deben quedar en bodega tras cruzar compras, consumo de cerdos y ventas externas.")

        c_q1, c_q2 = st.columns([1.2, 1])
        prod_consultar = c_q1.selectbox("Seleccionar Producto a Verificar:", LISTA_CONCENTRADOS + (df_bodega[df_bodega["Categoria_Insumo"] != "Concentrado"]["Nombre_Insumo"].unique().tolist() if not df_bodega.empty else []))
        consultar_btn = c_q2.button("🔎 Mostrar Saldo Exacto en Bodega", type="primary", use_container_width=True)

        if consultar_btn or prod_consultar:
            sub_b = df_bodega[df_bodega["Nombre_Insumo"] == prod_consultar] if not df_bodega.empty else pd.DataFrame()
            kg_comprados = float((to_num(sub_b["Unidades_Compradas"]) * to_num(sub_b["Contenido_Por_Unidad"])).sum()) if not sub_b.empty else 0.0
            saldo_kilos = float(to_num(sub_b["Saldo_Disponible"]).sum()) if not sub_b.empty else 0.0
            tam_bulto = float(to_num(sub_b["Contenido_Por_Unidad"]).iloc[-1]) if not sub_b.empty else 40.0
            und_m = str(sub_b["Unidad_Medida"].iloc[-1]) if not sub_b.empty else "kg"

            # Cruce explícito con consumos internos y ventas externas
            kg_consumo_interno = 0.0
            if not df_consumo.empty:
                kg_consumo_interno = float(to_num(df_consumo[df_consumo["Tipo_Concentrado"] == prod_consultar]["Kilos_Servidos"]).sum())

            kg_vendido_externo = 0.0
            if not df_ventas.empty:
                v_cuido = df_ventas[(df_ventas["Linea_Negocio"].str.contains("Concentrado|Cuido", na=False)) & (df_ventas["ID_Origen"] == prod_consultar)]
                kg_vendido_externo = float(to_num(v_cuido["Peso_Pie_O_Kilos"]).sum())

            bultos_totales = saldo_kilos / max(tam_bulto, 1.0)
            bultos_cerrados = int(saldo_kilos // max(tam_bulto, 1.0))
            kilos_sueltos = round(saldo_kilos % max(tam_bulto, 1.0), 1)

            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"📦 Deben quedar ({prod_consultar})", f"{bultos_totales:.2f} bultos", delta=f"{bultos_cerrados} cerrados + {kilos_sueltos} {und_m}")
            m2.metric("📥 Total Comprado", f"{kg_comprados:,.1f} {und_m}", delta=f"{kg_comprados/max(tam_bulto,1):.1f} bultos")
            m3.metric("🐖 Consumido por Cerdos", f"{kg_consumo_interno:,.1f} {und_m}")
            m4.metric("🚚 Vendido a Terceros", f"{kg_vendido_externo:,.1f} {und_m}")

            # Verificador de conteo físico en bodega
            st.markdown("#### 📋 Verificación contra Conteo Físico")
            cf1, cf2 = st.columns(2)
            bultos_fisicos = cf1.number_input(f"¿Cuántos bultos cerrados contaste físicamente de {prod_consultar}?", min_value=0, value=bultos_cerrados, step=1)
            kilos_fisicos_sueltos = cf2.number_input(f"¿Cuántos {und_m} sueltos (bulto abierto) hay en báscula?", min_value=0.0, value=float(kilos_sueltos), step=0.5)
            total_fisico = (bultos_fisicos * tam_bulto) + kilos_fisicos_sueltos
            diferencia_fisica = round(total_fisico - saldo_kilos, 1)

            if abs(diferencia_fisica) <= 0.5:
                st.success(f"✅ **INVENTARIO CUADRADO:** El conteo físico ({total_fisico:.1f} {und_m}) coincide exactamente con el sistema.")
            elif diferencia_fisica < 0:
                st.error(f"⚠️ **FALTANTE DETECTADO:** Faltan **{abs(diferencia_fisica):.1f} {und_m}** ({abs(diferencia_fisica)/tam_bulto:.2f} bultos) frente a lo registrado. Revisar si falta anotar una ración servida o una venta.")
            else:
                st.info(f"ℹ️ **SOBRANTE DETECTADO:** Hay **{diferencia_fisica:.1f} {und_m}** más en físico que en sistema. Revisar si faltó ingresar una factura de compra.")

    # --- PESTAÑA 2: SERVIR ALIMENTO A CERDOS CON GUÍA DE FASE ---
    with tab_b2:
        st.subheader("🍽️ Registrar Consumo Diario de Concentrado en Granja")
        conc_disp = df_bodega[(df_bodega["Categoria_Insumo"] == "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()
        if conc_disp.empty:
            st.warning("No hay concentrado con saldo disponible en bodega.")
        else:
            dest_tipo = st.selectbox("Destino del Alimento", ["Lote de Engorde", "Cerda Madre / Verraco"])
            if dest_tipo == "Lote de Engorde":
                lista_dest = lotes_activos["ID_Lote"].tolist() if not lotes_activos.empty else []
            else:
                lista_dest = df_rep["ID_Animal"].tolist() if not df_rep.empty else []

            id_dest = st.selectbox("Seleccionar Lote o Animal", lista_dest)

            # Guía automática de fase para el lote seleccionado
            fase_sugerida = conc_disp["Nombre_Insumo"].unique().tolist()[0]
            if dest_tipo == "Lote de Engorde" and id_dest and not lotes_activos.empty:
                fila_ld = lotes_activos[lotes_activos["ID_Lote"] == id_dest].iloc[0]
                ev_ld = evaluar_fase_lote(fila_ld, df_consumo, hoy, EDAD_DESTETE)
                st.info(
                    f"📌 **Guía Nutricional ({id_dest}):** Fase actual: **{ev_ld['fase_actual']}** | "
                    f"Consumido en esta fase: **{ev_ld['kg_consumidos_fase']:.1f} / {ev_ld['kg_meta_lote']:.1f} kg** "
                    f"(Faltan **{ev_ld['kg_restantes_fase']:.1f} kg** para pasar a **{ev_ld['siguiente_fase']}**)."
                )
                if ev_ld["fase_actual"] in conc_disp["Nombre_Insumo"].unique().tolist():
                    fase_sugerida = ev_ld["fase_actual"]

            tipos_conc = conc_disp["Nombre_Insumo"].unique().tolist()
            idx_c = tipos_conc.index(fase_sugerida) if fase_sugerida in tipos_conc else 0
            t_conc_sel = st.selectbox("Concentrado a Servir", tipos_conc, index=idx_c)
            saldo_tot_c = float(to_num(conc_disp[conc_disp["Nombre_Insumo"] == t_conc_sel]["Saldo_Disponible"]).sum())
            st.caption(f"Disponible en bodega de **{t_conc_sel}**: **{saldo_tot_c:.1f} kg** ({saldo_tot_c/40:.2f} bultos)")

            kg_serv = st.number_input("Kilos Servidos (kg)", min_value=0.5, max_value=max(saldo_tot_c, 0.5), value=min(10.0, max(saldo_tot_c, 0.5)), step=0.5)

            if st.button("💾 Descontar Alimento y Cargar Costo al Lote"):
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
                st.success(f"Descontados {kg_serv} kg de {t_conc_sel} y cargados a {id_dest}.")
                st.rerun()

    # --- PESTAÑA 3: INGRESAR COMPRA A BODEGA ---
    with tab_b3:
        st.subheader("📦 Registrar Compra de Concentrado o Medicamentos")
        cat_ins = st.selectbox("Categoría de Insumo", ["Concentrado", "Antibiótico / Inyectable", "Vacuna / Biológico", "Vitamina / Desparasitante"])
        if cat_ins == "Concentrado":
            nom_ins = st.selectbox("Tipo de Concentrado", LISTA_CONCENTRADOS)
            pres = "Bulto"
            und_med = "kg"
            cont_und = st.number_input("Kilos por Bulto", min_value=1.0, value=40.0)
        else:
            nom_ins = st.text_input("Nombre del Medicamento / Vacuna", placeholder="Ej: Florfenicol Inyectable")
            pres = st.selectbox("Presentación", ["Frasco", "Caja / Dosis", "Sobre / Bolsa"])
            und_med = st.selectbox("Unidad de Descuento", ["ml", "dosis", "g"])
            cont_und = st.number_input(f"Contenido por {pres} ({und_med})", min_value=1.0, value=100.0)

        cant_comp = st.number_input(f"Cantidad de {pres}s comprados", min_value=1, value=1)
        precio_u = st.number_input(f"Precio de Compra por {pres} ($)", min_value=0.0, value=85000.0, step=1000.0)
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

    st.divider()
    st.subheader("📊 Resumen General de Bodega y Botiquín")
    if not df_bodega.empty:
        df_bodega["Saldo_Disponible"] = to_num(df_bodega["Saldo_Disponible"])
        res_bod = df_bodega.groupby(["Categoria_Insumo", "Nombre_Insumo", "Unidad_Medida"], as_index=False).agg(
            Saldo_Total=("Saldo_Disponible", "sum"),
            Tam_Bulto=("Contenido_Por_Unidad", "last"),
            Costo_Promedio_Unidad=("Costo_Por_Medida", "mean")
        )
        res_bod["Bultos / Unidades"] = (res_bod["Saldo_Total"] / res_bod["Tam_Bulto"].replace(0, 1)).round(2)
        if es_admin:
            res_bod["Valor_Inventario ($)"] = (res_bod["Saldo_Total"] * res_bod["Costo_Promedio_Unidad"]).round(0)
            cols_ver = ["Categoria_Insumo", "Nombre_Insumo", "Saldo_Total", "Unidad_Medida", "Bultos / Unidades", "Valor_Inventario ($)"]
        else:
            cols_ver = ["Categoria_Insumo", "Nombre_Insumo", "Saldo_Total", "Unidad_Medida", "Bultos / Unidades"]
        st.dataframe(res_bod[cols_ver], use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 7: SANIDAD Y MORTALIDAD
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
            saldo_m = float(to_num(sub_m["Saldo_Disponible"]).sum())
            unidad_ap = str(sub_m.iloc[0]["Unidad_Medida"])
            costo_u_m = float(to_num(sub_m["Costo_Por_Medida"]).mean())
            cant_aplicada = c1.number_input(f"Cantidad Aplicada ({unidad_ap}) [Disponible: {saldo_m:.1f}]", min_value=0.1, max_value=max(saldo_m, 0.1), value=min(5.0, max(saldo_m, 0.1)), step=0.5)
            costo_san = round(cant_aplicada * costo_u_m, 0)
            if es_admin:
                st.info(f"Costo calculado automáticamente según botiquín: **${costo_san:,.0f}**")
        else:
            med_nombre = c2.text_input("Medicamento Aplicado (Externo)", value="Ninguno")
            costo_san = c1.number_input("Costo del Tratamiento ($)", min_value=0.0, value=0.0, step=1000.0) if es_admin else 0.0

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
# MÓDULO 8: VENTAS (MARRANOS, DESCARTE Y DISTRIBUCIÓN DE CUIDO) - SOLO ADMIN
# ==============================================================================
elif menu == "💰 Ventas (Marranos / Cuido)":
    st.markdown('<div class="main-header">💰 Registro de Ventas por Línea de Negocio</div>', unsafe_allow_html=True)

    linea_venta = st.selectbox("Selecciona la Línea de Ingreso:", [
        "🐖 Venta Cerdos Cebados (Pie / Canal)",
        "🌾 Venta Concentrado / Cuido (Distribución desde Bodega)",
        "🐷 Venta Cerda de Descarte",
        "🍼 Venta Lechones Destetos"
    ])

    # --- LÍNEA A: VENTA DE CUIDO / CONCENTRADO ---
    if "Concentrado" in linea_venta:
        st.subheader("🌾 Venta Comercial de Concentrado desde Bodega")
        conc_v = df_bodega[(df_bodega["Categoria_Insumo"] == "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()
        if conc_v.empty:
            st.warning("No hay concentrado con saldo disponible en bodega para vender.")
        else:
            c1, c2 = st.columns(2)
            prod_v = c1.selectbox("Concentrado a Vender", conc_v["Nombre_Insumo"].unique().tolist())
            sub_pv = conc_v[conc_v["Nombre_Insumo"] == prod_v]
            saldo_kg_v = float(to_num(sub_pv["Saldo_Disponible"]).sum())
            tam_b_v = float(to_num(sub_pv["Contenido_Por_Unidad"]).iloc[-1])
            costo_kg_prom = float(to_num(sub_pv["Costo_Por_Medida"]).mean())
            costo_bulto_prom = costo_kg_prom * tam_b_v

            st.info(f"Disponible: **{saldo_kg_v:.1f} kg** ({saldo_kg_v/tam_b_v:.1f} bultos de {tam_b_v:.0f} kg) | Costo Bodega: **${costo_bulto_prom:,.0f} / bulto** (${costo_kg_prom:,.1f}/kg)")

            bultos_v = c2.number_input("Cantidad de Bultos a Vender", min_value=0.25, max_value=max(saldo_kg_v / tam_b_v, 0.25), value=min(1.0, max(saldo_kg_v / tam_b_v, 0.25)), step=0.5)
            kilos_v = round(bultos_v * tam_b_v, 2)
            margen_pct = c1.number_input("Porcentaje de Margen Comercial (%) [Sugerido: 18%]", min_value=0.0, max_value=100.0, value=float(MARGEN_CUIDO_DEFAULT), step=1.0)
            costo_flete_v = c2.number_input("Costo Flete / Transporte / Peajes Asumido ($)", min_value=0.0, value=0.0, step=5000.0)
            cliente_txt = st.text_input("Cliente / Finca Destino", placeholder="Ej: Finca El Roble")

            costo_mercancia = kilos_v * costo_kg_prom
            precio_venta_bulto = costo_bulto_prom * (1 + margen_pct / 100.0)
            ingreso_b_cuido = bultos_v * precio_venta_bulto
            utilidad_n_cuido = ingreso_b_cuido - costo_mercancia - costo_flete_v

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Precio Venta / Bulto", f"${precio_venta_bulto:,.0f}")
            m2.metric("Ingreso Bruto Venta", f"${ingreso_b_cuido:,.0f}")
            m3.metric("Costo Bodega + Flete", f"${(costo_mercancia + costo_flete_v):,.0f}")
            m4.metric("Utilidad Neta Comercial", f"${utilidad_n_cuido:,.0f}")

            if st.button("💾 Confirmar Venta de Concentrado"):
                rest = kilos_v
                costo_real_desc = 0.0
                idxs = df_bodega[(df_bodega["Categoria_Insumo"] == "Concentrado") & (df_bodega["Nombre_Insumo"] == prod_v) & (to_num(df_bodega["Saldo_Disponible"]) > 0)].index
                for idx in idxs:
                    if rest <= 0:
                        break
                    disp = float(to_num(pd.Series([df_bodega.at[idx, "Saldo_Disponible"]])).iloc[0])
                    ckg = float(to_num(pd.Series([df_bodega.at[idx, "Costo_Por_Medida"]])).iloc[0])
                    tomar = min(disp, rest)
                    df_bodega.at[idx, "Saldo_Disponible"] = round(disp - tomar, 2)
                    costo_real_desc += tomar * ckg
                    rest -= tomar

                util_final_c = ingreso_b_cuido - costo_real_desc - costo_flete_v
                nueva_vc = pd.DataFrame([{
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Linea_Negocio": "Venta Concentrado / Cuido",
                    "ID_Origen": prod_v,
                    "Modalidad": f"{bultos_v} bultos (+{margen_pct:.0f}%) | {cliente_txt}",
                    "Cantidad_Animales_O_Unidades": bultos_v,
                    "Peso_Pie_O_Kilos": kilos_v,
                    "Peso_Carne_Canal_Kg": 0.0,
                    "Costo_Sacrificio_O_Flete": costo_flete_v,
                    "Ingreso_Bruto": round(ingreso_b_cuido, 0),
                    "Costo_Produccion_Asignado": round(costo_real_desc, 0),
                    "Utilidad_Neta": round(util_final_c, 0)
                }])
                df_ventas = pd.concat([df_ventas, nueva_vc], ignore_index=True)
                guardar_hoja("Bodega", df_bodega)
                guardar_hoja("Ventas", df_ventas)
                st.success(f"Venta registrada: {bultos_v} bultos de {prod_v} descontados de bodega.")
                st.rerun()

    # --- LÍNEA B: VENTA DE CERDA DE DESCARTE ---
    elif "Descarte" in linea_venta:
        st.subheader("🐷 Venta de Cerda de Descarte")
        hembras_act = df_rep[(df_rep["Tipo"].str.contains("Hembra", na=False)) & (~df_rep["Estado"].str.contains("Descartada", na=False))] if not df_rep.empty else pd.DataFrame()
        if hembras_act.empty:
            st.info("No hay cerdas activas para descarte.")
        else:
            c1, c2 = st.columns(2)
            cerda_d = c1.selectbox("Cerda a Descartar y Vender", hembras_act["ID_Animal"].tolist())
            peso_d = c2.number_input("Peso en Pie de la Cerda (kg)", min_value=80.0, max_value=350.0, value=190.0)
            precio_kg_d = c1.number_input("Precio Kilo Descarte ($)", min_value=0.0, value=6200.0, step=100.0)
            costo_acum_d = float(to_num(hembras_act[hembras_act["ID_Animal"] == cerda_d]["Costo_Acumulado_Ciclo"]).iloc[0])
            ing_d = peso_d * precio_kg_d
            util_d = ing_d - costo_acum_d
            c2.metric("Ingreso Neto Descarte", f"${util_d:,.0f}")

            if st.button("💾 Confirmar Venta de Cerda Descarte"):
                nueva_vd = pd.DataFrame([{
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Linea_Negocio": "Venta Cerda Descarte",
                    "ID_Origen": cerda_d,
                    "Modalidad": "Descarte en Pie",
                    "Cantidad_Animales_O_Unidades": 1,
                    "Peso_Pie_O_Kilos": peso_d,
                    "Peso_Carne_Canal_Kg": 0.0,
                    "Costo_Sacrificio_O_Flete": 0.0,
                    "Ingreso_Bruto": round(ing_d, 0),
                    "Costo_Produccion_Asignado": round(costo_acum_d, 0),
                    "Utilidad_Neta": round(util_d, 0)
                }])
                df_ventas = pd.concat([df_ventas, nueva_vd], ignore_index=True)
                df_rep.loc[df_rep["ID_Animal"] == cerda_d, "Estado"] = "Descartada / Vendida"
                df_rep.loc[df_rep["ID_Animal"] == cerda_d, "Observaciones"] = f"Vendida por descarte el {hoy.strftime('%Y-%m-%d')}"
                guardar_hoja("Ventas", df_ventas)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Cerda {cerda_d} vendida y marcada como Descartada.")
                st.rerun()

    # --- LÍNEA C: VENTA DE CERDOS CEBADOS O LECHONES ---
    else:
        es_desteto = "Lechones" in linea_venta
        st.subheader("🍼 Venta de Lechones Destetos" if es_desteto else "🐖 Venta de Cerdos Cebados (Pie / Canal)")
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
                        st.error(f"🚫 BLOQUEO SANITARIO: Tiempo de retiro vigente hasta `{fr}` por `{rs['Medicamento']}`.")
                        bloqueado = True

            cant_disp = int(to_num(pd.Series([fila_v["Cant_Actual"]])).iloc[0])
            c_tot_lote = float(to_num(pd.Series([fila_v["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Sanidad"]])).iloc[0])

            c1, c2 = st.columns(2)
            if es_desteto:
                modalidad = "Lechón Desteto por Unidad"
                cant_v = c1.number_input("Cantidad de Lechones Vendidos", min_value=1, max_value=max(cant_disp, 1), value=cant_disp)
                peso_pie_v = c2.number_input("Peso Total Camada (kg)", min_value=1.0, value=float(cant_v * to_num(pd.Series([fila_v["Peso_Prom_Actual_Kg"]])).iloc[0]))
                precio_und = c1.number_input("Precio Venta por Lechón ($)", min_value=0.0, value=240000.0, step=5000.0)
                peso_canal_v = 0.0
                costo_sac = 0.0
                ingreso_bruto = cant_v * precio_und
            else:
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
                    "Linea_Negocio": "Venta Lechones Destetos" if es_desteto else "Venta Cerdos Cebados",
                    "ID_Origen": lote_v,
                    "Modalidad": modalidad,
                    "Cantidad_Animales_O_Unidades": cant_v,
                    "Peso_Pie_O_Kilos": peso_pie_v,
                    "Peso_Carne_Canal_Kg": peso_canal_v,
                    "Costo_Sacrificio_O_Flete": costo_sac,
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
                st.success("Venta registrada y rentabilidad actualizada.")
                st.rerun()

    st.divider()
    st.subheader("Historial General de Ventas")
    st.dataframe(df_ventas, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 9: ANALÍTICA, METAS Y REPORTES (WORD / PDF) - SOLO ADMIN
# ==============================================================================
elif menu == "📊 Analítica, Metas y Reportes":
    st.markdown('<div class="main-header">📊 Inteligencia Productiva, Curva 120 kg y Reportes</div>', unsafe_allow_html=True)

    tab_a1, tab_a2, tab_a3 = st.tabs(["📈 Curva de Peso por Etapas (120 kg)", "🧬 Genética y Rentabilidad", "📄 Exportar Reporte Ejecutivo (Word / PDF)"])

    with tab_a1:
        st.subheader("Curva de Crecimiento Fija (21d Destete → 120 kg a los 153 días)")
        dias_eje = [0, 21, 28, 35, 45, 70, 90, 112, 132, 153]
        peso_eje = [peso_esperado_por_edad(d, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE) for d in dias_eje]

        fig_curva = go.Figure()
        fig_curva.add_trace(go.Scatter(x=dias_eje, y=peso_eje, mode="lines+markers", name="Curva Objetivo (120 kg)", line=dict(color="#1E8449", width=3, dash="dash")))

        if not lotes_activos.empty:
            for _, l in lotes_activos.iterrows():
                ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE)
                p_a = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                fig_curva.add_trace(go.Scatter(
                    x=[ev_l["edad_actual"]], y=[p_a], mode="markers+text",
                    name=l["ID_Lote"], text=[f"{l['ID_Lote']} ({p_a} kg - {ev_l['fase_actual']})"],
                    textposition="top center", marker=dict(size=14)
                ))

        fig_curva.update_layout(xaxis_title="Edad del Cerdo (Días)", yaxis_title="Peso Promedio (kg)", height=430)
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
            st.subheader("Utilidad Neta por Línea de Negocio")
            if not df_ventas.empty:
                df_ventas["Util_Num"] = to_num(df_ventas["Utilidad_Neta"])
                fig_v = px.bar(df_ventas.groupby("Linea_Negocio", as_index=False)["Util_Num"].sum(), x="Linea_Negocio", y="Util_Num", color="Linea_Negocio", title="Utilidad Neta Acumulada por Línea ($)")
                st.plotly_chart(fig_v, use_container_width=True)

    with tab_a3:
        st.subheader("📄 Generador de Informe Ejecutivo Integral")
        capex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Inversión Infraestructura / Activos"]["Monto"]).sum() if not df_inv.empty else 0.0
        opex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Gasto Fijo Operativo"]["Monto"]).sum() if not df_inv.empty else 0.0
        reinv_tot = to_num(df_inv[df_inv["Tipo_Movimiento"].str.contains("Reinversión", na=False)]["Monto"]).sum() if not df_inv.empty else 0.0
        ing_tot = to_num(df_ventas["Ingreso_Bruto"]).sum() if not df_ventas.empty else 0.0
        util_lotes = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0.0
        util_granja = util_lotes - opex_tot

        def generar_word():
            doc = Document()
            titulo = doc.add_heading("INFORME EJECUTIVO DE GESTIÓN PORCINA - PORCIGESTIÓN PRO", 0)
            titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_fecha = doc.add_paragraph(f"Fecha de emisión: {hoy.strftime('%d/%m/%Y')} | Programa: Destete {EDAD_DESTETE}d -> {PESO_META:.0f} kg a los {EDAD_META} días")
            p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_heading("1. Resumen Financiero, Reinversión e Inversiones", level=1)
            resumen_fin_txt = "\n".join([
                f"• Inversión acumulada en infraestructura y activos (CAPEX): ${capex_tot:,.0f}",
                f"• Gastos fijos operativos acumulados (OPEX): ${opex_tot:,.0f}",
                f"• Ingresos brutos totales (Cerdos + Descarte + Cuido): ${ing_tot:,.0f}",
                f"• Utilidad operativa bruta de ventas: ${util_lotes:,.0f}",
                f"• Utilidad neta real de la granja (descontando gastos fijos): ${util_granja:,.0f}",
                f"• Capital reinvertido desde utilidades: ${reinv_tot:,.0f}"
            ])
            doc.add_paragraph(resumen_fin_txt)

            doc.add_heading("2. Rendimiento de Lotes Activos y Fases Nutricionales", level=1)
            if not lotes_activos.empty:
                t_lot = doc.add_table(rows=1, cols=6)
                t_lot.style = "Table Grid"
                hl = t_lot.rows[0].cells
                hl[0].text = "Lote"
                hl[1].text = "Fase Actual"
                hl[2].text = "Cerdos"
                hl[3].text = "Peso Real (kg)"
                hl[4].text = "Alimento Fase"
                hl[5].text = "Costo Total ($)"
                for _, l in lotes_activos.iterrows():
                    ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE)
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    ctot = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
                    rc = t_lot.add_row().cells
                    rc[0].text = str(l["ID_Lote"])
                    rc[1].text = str(ev_l["fase_actual"])
                    rc[2].text = str(cant)
                    rc[3].text = f"{p_act:.1f} kg"
                    rc[4].text = f"{ev_l['kg_consumidos_fase']:.1f}/{ev_l['kg_meta_lote']:.1f} kg"
                    rc[5].text = f"${ctot:,.0f}"

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
            pdf.cell(0, 8, "1. RESUMEN FINANCIERO Y REINVERSION", ln=True)
            pdf.set_font("Arial", "", 10)
            lines_fin = [
                f"- Inversion Infraestructura (CAPEX): ${capex_tot:,.0f}",
                f"- Gastos Fijos Operativos (OPEX): ${opex_tot:,.0f}",
                f"- Ingresos Brutos Totales: ${ing_tot:,.0f}",
                f"- Utilidad Neta Real Granja: ${util_granja:,.0f}",
                f"- Utilidades Reinvertidas: ${reinv_tot:,.0f}"
            ]
            for lf in lines_fin:
                pdf.cell(0, 6, lf, ln=True)
            pdf.ln(4)

            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 8, "2. LOTES ACTIVOS Y FASES NUTRICIONALES", ln=True)
            pdf.set_font("Arial", "", 9)
            if not lotes_activos.empty:
                for _, l in lotes_activos.iterrows():
                    ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE)
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    txt = f"{l['ID_Lote']} | {ev_l['fase_actual']} ({ev_l['kg_consumidos_fase']:.1f}/{ev_l['kg_meta_lote']:.1f} kg) | {cant} cerdos | Peso: {p_act:.1f} kg"
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
# MÓDULO 10: GESTIONAR CERDAS Y METAS - SOLO ADMIN
# ==============================================================================
elif menu == "⚙️ Gestionar Cerdas y Metas":
    st.markdown('<div class="main-header">⚙️ Administración de Reproductores y Parámetros</div>', unsafe_allow_html=True)

    t_cfg1, t_cfg2 = st.tabs(["🎯 Configurar Metas y Margen Cuido", "🐷 Agregar / Editar Reproductores"])

    with t_cfg1:
        st.subheader("Parámetros Técnicos y Comerciales")
        c1, c2, c3 = st.columns(3)
        n_peso_meta = c1.number_input("Peso Objetivo Finalización (kg)", min_value=70.0, max_value=150.0, value=float(PESO_META), step=1.0)
        n_edad_meta = c2.number_input("Edad Objetivo Mercado (días)", min_value=100, max_value=220, value=int(EDAD_META), step=1)
        n_edad_dest = c3.number_input("Edad Destete Lechones (días)", min_value=18, max_value=35, value=int(EDAD_DESTETE), step=1)

        c4, c5, c6 = st.columns(3)
        n_margen_cuido = c4.number_input("Margen Venta Cuido por Defecto (%)", min_value=0.0, max_value=100.0, value=float(MARGEN_CUIDO_DEFAULT), step=1.0)
        n_int_bandas = c5.selectbox("Ritmo entre Bandas (días)", [14, 15, 21], index=[14, 15, 21].index(INTERVALO_BANDAS) if INTERVALO_BANDAS in [14, 15, 21] else 1)
        n_cupo_banda = c6.number_input("Cupo Máx. Cerdas por Banda", min_value=1, max_value=10, value=int(CUPO_BANDA))

        if st.button("💾 Guardar Configuración"):
            nueva_cfg = pd.DataFrame([
                {"Parametro": "Peso_Meta_Kg", "Valor": str(n_peso_meta), "Descripcion": "Peso objetivo del cerdo finalizado en pie (kg)"},
                {"Parametro": "Edad_Meta_Dias", "Valor": str(n_edad_meta), "Descripcion": "Edad objetivo desde nacimiento (días)"},
                {"Parametro": "Edad_Destete_Dias", "Valor": str(n_edad_dest), "Descripcion": "Edad estándar de destete (días)"},
                {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": str(PESO_NAC), "Descripcion": "Peso promedio al nacer (kg)"},
                {"Parametro": "Peso_Destete_21d_Kg", "Valor": str(PESO_DEST_21D), "Descripcion": "Peso objetivo al destete a 21 días (kg)"},
                {"Parametro": "Margen_Venta_Cuido_Pct", "Valor": str(n_margen_cuido), "Descripcion": "Margen comercial por defecto venta concentrado (%)"},
                {"Parametro": "Intervalo_Bandas_Dias", "Valor": str(n_int_bandas), "Descripcion": "Días entre bandas"},
                {"Parametro": "Numero_Bandas", "Valor": str(NUM_BANDAS), "Descripcion": "Número de bandas en ciclo completo"},
                {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": str(n_cupo_banda), "Descripcion": "Máximo de cerdas por banda"},
                {"Parametro": "Alerta_Dias_Abiertos", "Valor": str(MAX_DIAS_ABIERTOS), "Descripcion": "Días máximos vacía antes de alerta"}
            ])
            guardar_hoja("Configuracion", nueva_cfg)
            st.success("Configuración guardada correctamente.")
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
            banda_n = st.selectbox(f"Banda Asignada (Sugerida: {banda_sug})", lista_b, index=lista_b.index(banda_sug))
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