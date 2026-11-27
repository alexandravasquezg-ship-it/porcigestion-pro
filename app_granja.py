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
import re

st.set_page_config(
    page_title="PorciGestión Pro v3 | Control Integral Porcino",
    page_icon="🐽",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {font-size: 2rem; font-weight: 800; color: #1E8449; margin-bottom: 0.2rem;}
    .sub-header {font-size: 1rem; color: #566573; margin-bottom: 1.5rem;}
    .metric-card {background-color: #F8F9F9; padding: 15px; border-radius: 10px; border-left: 5px solid #1E8449;}
</style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# ==============================================================================
# LISTAS MAESTRAS Y ESQUEMA DE HOJAS
# ==============================================================================
RAZAS_SEMEN = [
    "PIC 337", "PIC 337 Superior", "PIC 408", "PIC 410",
    "Pietrain", "Duroc", "Large White", "Landrace", "L02 PIC", "L03 PIC"
]

RAZAS_MARRANAS = [
    "L02 PIC", "L03 PIC", "Camborough PIC", "F1 (Landrace x Yorkshire)",
    "Landrace", "Large White", "PIC 337", "PIC 337 Superior",
    "PIC 408", "PIC 410", "Pietrain", "Duroc"
]

PROGRAMA_FASES_BASE = [
    {"fase": "Fase 0", "duracion": 7, "peso_ini": 6.2, "peso_fin": 7.8, "kg_cerdo": 1.2, "siguiente": "Fase 1"},
    {"fase": "Fase 1", "duracion": 7, "peso_ini": 7.8, "peso_fin": 10.2, "kg_cerdo": 2.0, "siguiente": "Fase 2"},
    {"fase": "Fase 2", "duracion": 10, "peso_ini": 10.2, "peso_fin": 14.5, "kg_cerdo": 5.0, "siguiente": "Iniciación"},
    {"fase": "Iniciación", "duracion": 25, "peso_ini": 14.5, "peso_fin": 30.0, "kg_cerdo": 21.0, "siguiente": "Prelevante"},
    {"fase": "Prelevante", "duracion": 20, "peso_ini": 30.0, "peso_fin": 46.0, "kg_cerdo": 27.0, "siguiente": "Levante"},
    {"fase": "Levante", "duracion": 22, "peso_ini": 46.0, "peso_fin": 68.0, "kg_cerdo": 40.0, "siguiente": "Engorde"},
    {"fase": "Engorde", "duracion": 20, "peso_ini": 68.0, "peso_fin": 92.0, "kg_cerdo": 40.0, "siguiente": "Finalizador"},
    {"fase": "Finalizador", "duracion": 21, "peso_ini": 92.0, "peso_fin": 120.0, "kg_cerdo": 80.0, "siguiente": "Venta / Mercado (120 kg)"}
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
    {"Parametro": "Edad_Meta_Dias", "Valor": "153", "Descripcion": "Edad objetivo desde nacimiento para alcanzar 120 kg"},
    {"Parametro": "Edad_Destete_Dias", "Valor": "21", "Descripcion": "Edad estándar de destete de lechones (días)"},
    {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": "1.4", "Descripcion": "Peso promedio al nacer (kg)"},
    {"Parametro": "Peso_Destete_21d_Kg", "Valor": "6.2", "Descripcion": "Peso objetivo al destete (kg)"},
    {"Parametro": "Margen_Venta_Cuido_Pct", "Valor": "18.0", "Descripcion": "Porcentaje de margen comercial por defecto para venta de concentrado (%)"},
    {"Parametro": "Meta_Cerdos_Mes", "Valor": "40", "Descripcion": "Meta final mensual de cerdos cebados vendidos (unidades/mes)"},
    {"Parametro": "Meta_Intermedia_Cerdos_Mes", "Valor": "20", "Descripcion": "Meta intermedia o primer escalón mensual de cerdos vendidos (unidades/mes)"},
    {"Parametro": "Meta_Utilidad_Por_Cerdo", "Valor": "200000", "Descripcion": "Utilidad neta objetivo por cada cerdo vendido ($)"},
    {"Parametro": "Intervalo_Bandas_Dias", "Valor": "15", "Descripcion": "Días entre inseminación de una banda y la siguiente"},
    {"Parametro": "Numero_Bandas", "Valor": "10", "Descripcion": "Total de bandas en el ciclo completo"},
    {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": "2", "Descripcion": "Máximo de cerdas a inseminar por banda"},
    {"Parametro": "Alerta_Dias_Abiertos", "Valor": "7", "Descripcion": "Días máximos que una cerda vacía puede estar sin servicio"}
])

# ==============================================================================
# UTILIDADES DE INTEGRIDAD, LECTURA SEGURA Y CÁLCULO
# ==============================================================================
def asegurar_columnas(df, cols):
    if df is None:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]

@st.cache_data(ttl=60)
def cargar_datos_seguros():
    datos = {}
    errores = []
    for hoja, cols in ESQUEMA_HOJAS.items():
        try:
            df = conn.read(worksheet=hoja, ttl=0)
            if df is None:
                errores.append(f"La hoja '{hoja}' devolvió una respuesta nula.")
                continue
            df = df.dropna(how="all")
            datos[hoja] = asegurar_columnas(df, cols)
        except Exception as e:
            errores.append(f"Error leyendo hoja '{hoja}': {e}")
    return datos, errores

def guardar_hoja(hoja, df):
    if st.session_state.get("lectura_corrupta", False):
        st.error("🚫 Guardado bloqueado por seguridad: hubo un error al leer Google Sheets.")
        st.stop()
    df_clean = asegurar_columnas(df.copy(), ESQUEMA_HOJAS[hoja])
    conn.update(worksheet=hoja, data=df_clean)
    st.cache_data.clear()

def generar_siguiente_id(df, col_id, prefijo):
    max_num = 0
    if df is not None and not df.empty and col_id in df.columns:
        for val in df[col_id].dropna().astype(str):
            m = re.search(rf"^{re.escape(prefijo)}-(\d+)", val.strip())
            if m:
                max_num = max(max_num, int(m.group(1)))
    return f"{prefijo}-{max_num + 1:03d}"

def to_num(series, default=0.0):
    return pd.to_numeric(series, errors="coerce").fillna(default)

def parse_fecha(val):
    if pd.isna(val) or str(val).strip() in ("", "None", "nan", "NaT"):
        return None
    try:
        return pd.to_datetime(str(val).strip()[:10]).date()
    except Exception:
        return None

def get_config(df_cfg, param, default_val):
    if df_cfg is None or df_cfg.empty:
        return default_val
    fila = df_cfg[df_cfg["Parametro"].astype(str).str.strip() == param]
    if fila.empty:
        return default_val
    try:
        return float(fila.iloc[0]["Valor"])
    except Exception:
        return default_val

def opciones_limpias(serie):
    if serie is None or len(serie) == 0:
        return []
    vals = [str(x).strip() for x in serie.dropna().tolist() if str(x).strip() not in ("", "nan", "None")]
    return list(dict.fromkeys(vals))

def construir_programa_fases(edad_destete, peso_destete, peso_meta):
    """Vincula dinámicamente las edades y pesos de cada fase con la edad/peso de destete configurados."""
    fases = []
    edad_cursor = int(edad_destete)
    factor_peso = peso_meta / 120.0
    for idx, f in enumerate(PROGRAMA_FASES_BASE):
        e_ini = edad_cursor
        e_fin = e_ini + f["duracion"]
        p_ini = peso_destete if idx == 0 else round(f["peso_ini"] * factor_peso, 2)
        p_fin = peso_meta if idx == len(PROGRAMA_FASES_BASE) - 1 else round(f["peso_fin"] * factor_peso, 2)
        fases.append({
            "fase": f["fase"],
            "edad_ini": e_ini,
            "edad_fin": e_fin,
            "peso_ini": p_ini,
            "peso_fin": p_fin,
            "kg_cerdo": f["kg_cerdo"],
            "siguiente": f["siguiente"]
        })
        edad_cursor = e_fin
    return fases

def peso_esperado_por_edad(edad_dias, peso_nac, peso_dest, edad_dest, fases_din):
    if edad_dias <= 0:
        return peso_nac
    if edad_dias <= edad_dest:
        return peso_nac + ((peso_dest - peso_nac) / max(edad_dest, 1)) * edad_dias
    for f in fases_din:
        if edad_dias <= f["edad_fin"]:
            span = max(f["edad_fin"] - f["edad_ini"], 1)
            prog = (edad_dias - f["edad_ini"]) / span
            return f["peso_ini"] + (f["peso_fin"] - f["peso_ini"]) * prog
    ultima = fases_din[-1]
    return ultima["peso_fin"] + (edad_dias - ultima["edad_fin"]) * 0.85

def fase_por_edad(edad_dias, edad_dest, peso_dest, fases_din):
    if edad_dias < edad_dest:
        return {"fase": "Lactancia", "edad_ini": 0, "edad_fin": edad_dest, "peso_ini": 1.4, "peso_fin": peso_dest, "kg_cerdo": 0.0, "siguiente": "Fase 0"}
    for f in fases_din:
        if edad_dias <= f["edad_fin"]:
            return f
    return fases_din[-1]

def evaluar_fase_lote(lote_row, df_cons, hoy_date, edad_dest, peso_dest, fases_din):
    lote_id = str(lote_row["ID_Lote"]).strip()
    cant_ini = int(to_num(pd.Series([lote_row["Cant_Inicial"]])).iloc[0])
    cant_act = int(to_num(pd.Series([lote_row["Cant_Actual"]])).iloc[0])
    f_ini = parse_fecha(lote_row["Fecha_Inicio"])
    edad_ini = int(to_num(pd.Series([lote_row.get("Edad_Inicial_Dias", 0)])).iloc[0])
    dias_t = max((hoy_date - f_ini).days, 1) if f_ini else 1
    edad_act = edad_ini + dias_t

    etapa_reg = str(lote_row.get("Etapa", "")).strip()
    nombres_fases = [f["fase"] for f in fases_din]
    info_edad = fase_por_edad(edad_act, edad_dest, peso_dest, fases_din)

    if etapa_reg in nombres_fases:
        info_f = next(f for f in fases_din if f["fase"] == etapa_reg)
    else:
        info_f = info_edad

    # Población efectiva promedio para evitar saltos artificiales por bajas o ventas parciales
    poblacion_efectiva = max((cant_ini + cant_act) / 2.0, float(cant_act), 1.0)
    kg_meta_fase_lote = round(info_f["kg_cerdo"] * poblacion_efectiva, 1)

    kg_consumidos_fase = 0.0
    if not df_cons.empty and info_f["kg_cerdo"] > 0:
        sub_c = df_cons[(df_cons["ID_Destino"].astype(str).str.strip() == lote_id) & (df_cons["Tipo_Concentrado"].astype(str).str.strip() == info_f["fase"])]
        kg_consumidos_fase = float(to_num(sub_c["Kilos_Servidos"]).sum())

    return {
        "edad_actual": edad_act,
        "fase_registrada": info_f["fase"],
        "fase_por_edad": info_edad["fase"],
        "siguiente_fase": info_f["siguiente"],
        "kg_por_cerdo_fase": info_f["kg_cerdo"],
        "kg_meta_lote": kg_meta_fase_lote,
        "kg_consumidos_fase": round(kg_consumidos_fase, 1),
        "kg_restantes_fase": round(max(kg_meta_fase_lote - kg_consumidos_fase, 0.0), 1),
        "edad_fin_fase": info_f["edad_fin"],
        "peso_objetivo_fin_fase": info_f["peso_fin"]
    }

def calcular_fifo_bodega(df_bod, cat, nombre, pres, unidad, cantidad_requerida):
    """Calcula el costo real exacto FIFO de una salida antes y durante la confirmación."""
    sub = df_bod[
        (df_bod["Categoria_Insumo"].astype(str).str.strip() == cat) &
        (df_bod["Nombre_Insumo"].astype(str).str.strip() == nombre) &
        (df_bod["Presentacion"].astype(str).str.strip() == pres) &
        (df_bod["Unidad_Medida"].astype(str).str.strip() == unidad) &
        (to_num(df_bod["Saldo_Disponible"]) > 0)
    ]
    # FIFO real: ordenar primero por fecha de compra y, en empate, por ID.
    # Si la fecha es inválida, se coloca al final para no consumirla antes que lotes fechados.
    sub = sub.copy()
    sub["_fecha_fifo"] = pd.to_datetime(sub["Fecha"], errors="coerce")
    sub["_id_fifo"] = sub["ID_Compra"].astype(str).str.strip()
    sub = sub.sort_values(["_fecha_fifo", "_id_fifo"], kind="stable", na_position="last")

    restante = max(float(cantidad_requerida), 0.0)
    costo_total_fifo = 0.0
    plan_descuento = []
    for idx, row in sub.iterrows():
        if restante <= 1e-6:
            break
        disp = float(to_num(pd.Series([row["Saldo_Disponible"]])).iloc[0])
        c_unit = float(to_num(pd.Series([row["Costo_Por_Medida"]])).iloc[0])
        tomar = min(disp, restante)
        costo_total_fifo += tomar * c_unit
        plan_descuento.append((idx, round(disp - tomar, 4)))
        restante -= tomar
    return round(costo_total_fifo, 2), plan_descuento, round(restante, 4)

# ==============================================================================
# CARGA SEGURA DE DATOS (DETIENE EJECUCIÓN SI FALLA GOOGLE SHEETS)
# ==============================================================================
datos, errores_lectura = cargar_datos_seguros()
if errores_lectura:
    st.session_state["lectura_corrupta"] = True
    st.error("🚨 **ERROR CRÍTICO DE CONEXIÓN CON GOOGLE SHEETS**")
    st.warning("Por seguridad e integridad de tus datos, la aplicación ha bloqueado todas las operaciones de escritura hasta restablecer la lectura completa de las hojas.")
    for err in errores_lectura:
        st.code(err)
    if st.button("🔄 Reintentar Conexión con Google Sheets"):
        st.cache_data.clear()
        st.session_state["lectura_corrupta"] = False
        st.rerun()
    st.stop()
else:
    st.session_state["lectura_corrupta"] = False

df_rep = datos["Reproductores"]
df_repro = datos["Reproduccion"]
df_lotes = datos["Lotes_Engorde"]
df_bodega = datos["Bodega"]
df_consumo = datos["Consumo_Alimento"]
df_sanidad = datos["Sanidad_Eventos"]
df_ventas = datos["Ventas"]
df_inv = datos["Inversiones_Gastos"]
df_cfg = datos["Configuracion"] if not datos["Configuracion"].empty else CONFIG_DEFAULT.copy()

hoy = date.today()

# Parámetros globales dinámicos
PESO_META = get_config(df_cfg, "Peso_Meta_Kg", 120.0)
EDAD_DESTETE = int(get_config(df_cfg, "Edad_Destete_Dias", 21))
PESO_NAC = get_config(df_cfg, "Peso_Nacimiento_Prom_Kg", 1.4)
PESO_DEST_21D = get_config(df_cfg, "Peso_Destete_21d_Kg", 6.2)
PROGRAMA_FASES = construir_programa_fases(EDAD_DESTETE, PESO_DEST_21D, PESO_META)
EDAD_META = PROGRAMA_FASES[-1]["edad_fin"]
MARGEN_CUIDO_DEFAULT = get_config(df_cfg, "Margen_Venta_Cuido_Pct", 18.0)
META_CERDOS_MES = int(get_config(df_cfg, "Meta_Cerdos_Mes", 40))
META_INTERMEDIA_MES = int(get_config(df_cfg, "Meta_Intermedia_Cerdos_Mes", 20))
META_UTIL_CERDO = float(get_config(df_cfg, "Meta_Utilidad_Por_Cerdo", 200000.0))
META_UTIL_MENSUAL = META_CERDOS_MES * META_UTIL_CERDO
INTERVALO_BANDAS = int(get_config(df_cfg, "Intervalo_Bandas_Dias", 15))
NUM_BANDAS = int(get_config(df_cfg, "Numero_Bandas", 10))
CUPO_BANDA = int(get_config(df_cfg, "Cupo_Max_Cerdas_Banda", 2))
MAX_DIAS_ABIERTOS = int(get_config(df_cfg, "Alerta_Dias_Abiertos", 7))

# ==============================================================================
# CONTROL DE ACCESO SEGURO (USANDO st.secrets SI ESTÁ CONFIGURADO)
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "rol_usuario" not in st.session_state:
    st.session_state["rol_usuario"] = None

try:
    ADMIN_PASS_REAL = st.secrets.get("admin_password", "admin123")
except Exception:
    # Streamlit puede no tener secrets.toml configurado; se conserva el fallback previo.
    ADMIN_PASS_REAL = "admin123"

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
    with col_l2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div style="font-size: 135px; line-height: 1.05; margin-bottom: 10px;">🐽</div>', unsafe_allow_html=True)
        st.markdown('<div class="main-header">PorciGestión Pro v3</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Sistema Blindado de Administración y Analítica Porcina</div>', unsafe_allow_html=True)

        perfil_sel = st.selectbox(
            "Selecciona tu perfil de ingreso:",
            ["Operario de Granja (Trabajador)", "Administrador (Gerencia)"]
        )

        if perfil_sel == "Operario de Granja (Trabajador)":
            st.info("👷 Ingreso directo para tareas operativas diarias. No requiere contraseña.")
            if st.button("Ingresar como Operario", type="primary", use_container_width=True):
                st.session_state["autenticado"] = True
                st.session_state["rol_usuario"] = "operario"
                st.rerun()
        else:
            clave_ing = st.text_input("Contraseña de Administrador:", type="password", placeholder="Ingresa la clave de administrador...")
            if st.button("Ingresar como Administrador", type="primary", use_container_width=True):
                if clave_ing == ADMIN_PASS_REAL:
                    st.session_state["autenticado"] = True
                    st.session_state["rol_usuario"] = "admin"
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta.")
    st.stop()

es_admin = (st.session_state["rol_usuario"] == "admin")

with st.sidebar:
    st.markdown('<div style="font-size: 115px; line-height: 1.05; margin-bottom: 5px;">🐽</div>', unsafe_allow_html=True)
    st.title("PorciGestión Pro v3")
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
            "🔮 Simulador y Proyección a Metas",
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
# FILTROS DE ESTADO ANIMAL Y CÁLCULO CENTRAL DE ALERTAS
# ==============================================================================
ESTADOS_INACTIVOS_REP = {"Descartada / Vendida", "Muerta", "Descartada"}

hembras_activas_df = df_rep[
    (df_rep["Tipo"].astype(str).str.contains("Hembra", na=False)) &
    (~df_rep["Estado"].astype(str).str.strip().isin(ESTADOS_INACTIVOS_REP))
].copy() if not df_rep.empty else pd.DataFrame()

madres_activas_tot = len(hembras_activas_df)

# Eventos reproductivos abiertos por cerda
eventos_abiertos_df = df_repro[
    df_repro["Estado_Prenez"].astype(str).str.strip().isin(["Pendiente Eco", "Preñada", "En Lactancia"])
].copy() if not df_repro.empty else pd.DataFrame()
cerdas_con_ciclo_abierto = set(opciones_limpias(eventos_abiertos_df["ID_Cerda"])) if not eventos_abiertos_df.empty else set()

alertas_criticas = []
alertas_preventivas = []

if not eventos_abiertos_df.empty:
    ids_hembras_activas_set = set(opciones_limpias(hembras_activas_df["ID_Animal"]))
    for _, r in eventos_abiertos_df.iterrows():
        cerda = str(r["ID_Cerda"]).strip()
        if cerda not in ids_hembras_activas_set:
            continue
        est = str(r["Estado_Prenez"]).strip()
        f_parto = parse_fecha(r["Fecha_Prob_Parto"])
        f_eco = parse_fecha(r["Fecha_Eco_28d"])
        f_dest = parse_fecha(r["Fecha_Prob_Destete"])

        if est == "Preñada" and f_parto:
            dias_p = (f_parto - hoy).days
            if dias_p < 0:
                alertas_criticas.append(f"🚨 **Parto Atrasado ({cerda} - {r.get('Banda','')}):** Tenía fecha probable `{f_parto}` (**{abs(dias_p)} días de atraso**). Revisar urgente.")
            elif dias_p <= 5:
                alertas_criticas.append(f"🔴 **Parto Próximo ({cerda} - {r.get('Banda','')}):** Programado para `{f_parto}` (en {dias_p} días). Preparar sala de maternidad.")
        if est == "Pendiente Eco" and f_eco and hoy >= f_eco:
            alertas_preventivas.append(f"🔵 **Confirmar Preñez ({cerda}):** Cumplió 28 días desde inseminación (`{f_eco}`). Realizar ecografía o chequeo de retorno.")
        if est == "En Lactancia" and f_dest and (hoy - f_dest).days >= -2:
            alertas_preventivas.append(f"🍼 **Destete Listo a los {EDAD_DESTETE} días ({cerda} / Lote {r.get('Nombre_Lote_Asignado','')}):** Fecha objetivo `{f_dest}`.")

if not hembras_activas_df.empty:
    hembras_vacias = hembras_activas_df[hembras_activas_df["Estado"].astype(str).str.strip() == "Vacía"]
    for _, h in hembras_vacias.iterrows():
        f_cambio = parse_fecha(h.get("Fecha_Ultimo_Cambio_Estado", ""))
        if f_cambio:
            dias_ab = (hoy - f_cambio).days
            if dias_ab > MAX_DIAS_ABIERTOS:
                alertas_criticas.append(f"⏳ **Cerda con Días Abiertos ({h['ID_Animal']} - {h.get('Banda','')}):** Lleva **{dias_ab} días vacía** (límite: {MAX_DIAS_ABIERTOS} días).")

lotes_activos = df_lotes[df_lotes["Estado"].astype(str).str.strip() == "Activo"].copy() if not df_lotes.empty else pd.DataFrame()
if not lotes_activos.empty:
    for _, l in lotes_activos.iterrows():
        ev_f = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
        edad_actual = ev_f["edad_actual"]
        peso_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
        peso_esperado_hoy = peso_esperado_por_edad(edad_actual, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE, PROGRAMA_FASES)
        dif_peso = peso_act - peso_esperado_hoy

        if ev_f["kg_meta_lote"] > 0:
            if ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"]:
                alertas_criticas.append(
                    f"🌾 **Cambio de Fase Nutricional ({l['ID_Lote']}):** Ya completó **{ev_f['kg_consumidos_fase']:.1f} / {ev_f['kg_meta_lote']:.1f} kg de {ev_f['fase_registrada']}**. "
                    f"Pasar a **{ev_f['siguiente_fase']}**."
                )
            elif ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"] * 0.88:
                alertas_preventivas.append(
                    f"📦 **Próximo Cambio de Fase ({l['ID_Lote']}):** Lleva **{ev_f['kg_consumidos_fase']:.1f} / {ev_f['kg_meta_lote']:.1f} kg de {ev_f['fase_registrada']}** "
                    f"(faltan **{ev_f['kg_restantes_fase']:.1f} kg**)."
                )
            elif edad_actual > ev_f["edad_fin_fase"]:
                alertas_preventivas.append(
                    f"⏰ **Desfase Edad vs Etapa ({l['ID_Lote']}):** Tiene {edad_actual} días (corresponde a **{ev_f['fase_por_edad']}**), pero sigue registrado en **{ev_f['fase_registrada']}**."
                )

        if edad_actual >= EDAD_DESTETE + 7 and dif_peso < -3.5:
            alertas_preventivas.append(
                f"📉 **Bajo Ritmo de Crecimiento ({l['ID_Lote']}):** Pesa **{peso_act:.1f} kg** a los {edad_actual} días vs **{peso_esperado_hoy:.1f} kg** esperados ({dif_peso:+.1f} kg)."
            )

if not df_sanidad.empty:
    for _, s in df_sanidad.iterrows():
        f_ret = parse_fecha(s.get("Fecha_Fin_Retiro", ""))
        if f_ret and f_ret >= hoy:
            alertas_criticas.append(f"🚫 **Tiempo de Retiro Vigente ({s['ID_Destino']}):** `{s['Medicamento']}` hasta `{f_ret}`.")

if not df_bodega.empty:
    df_bodega["Saldo_Num"] = to_num(df_bodega["Saldo_Disponible"])
    resumen_stock = df_bodega.groupby(["Categoria_Insumo", "Nombre_Insumo", "Presentacion", "Unidad_Medida"], as_index=False)["Saldo_Num"].sum()
    for _, st_row in resumen_stock.iterrows():
        if st_row["Categoria_Insumo"] == "Concentrado" and st_row["Saldo_Num"] < 120:
            alertas_preventivas.append(f"🌾 **Stock Bajo Concentrado ({st_row['Nombre_Insumo']}):** Quedan **{st_row['Saldo_Num']:.1f} kg** ({st_row['Saldo_Num']/40:.1f} bultos).")
        elif st_row["Categoria_Insumo"] != "Concentrado" and st_row["Saldo_Num"] < 20:
            alertas_preventivas.append(f"💊 **Botiquín Bajo ({st_row['Nombre_Insumo']} - {st_row['Presentacion']}):** Quedan **{st_row['Saldo_Num']:.1f} {st_row['Unidad_Medida']}**.")

# Ventas últimos 30 días (validando 0 <= (hoy - fecha) <= 30 para ignorar fechas futuras)
cerdos_vendidos_30d = 0
utilidad_cerdos_30d = 0.0
utilidad_prom_por_cerdo_hist = 0.0
costo_prod_prom_cerdo_hist = 0.0

if not df_ventas.empty:
    v_cebados = df_ventas[df_ventas["Linea_Negocio"].astype(str).str.contains("Cebados", na=False)].copy()
    if not v_cebados.empty:
        tot_cerdos_hist = float(to_num(v_cebados["Cantidad_Animales_O_Unidades"]).sum())
        tot_util_hist = float(to_num(v_cebados["Utilidad_Neta"]).sum())
        tot_costo_hist = float(to_num(v_cebados["Costo_Produccion_Asignado"]).sum())
        utilidad_prom_por_cerdo_hist = tot_util_hist / max(tot_cerdos_hist, 1.0)
        costo_prod_prom_cerdo_hist = tot_costo_hist / max(tot_cerdos_hist, 1.0)

        for _, vc in v_cebados.iterrows():
            fv = parse_fecha(vc["Fecha"])
            if fv and 0 <= (hoy - fv).days <= 30:
                cerdos_vendidos_30d += int(to_num(pd.Series([vc["Cantidad_Animales_O_Unidades"]])).iloc[0])
                utilidad_cerdos_30d += float(to_num(pd.Series([vc["Utilidad_Neta"]])).iloc[0])

# ==============================================================================
# MÓDULO 1: TABLERO DE ALERTAS
# ==============================================================================
if menu == "🚨 Tablero de Alertas":
    st.markdown('<div class="main-header">🚨 Centro de Control y Alertas Operativas</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="sub-header">Fecha: <b>{hoy.strftime("%d/%m/%Y")}</b> | Escalón 1: <b>{META_INTERMEDIA_MES} cerdos/mes</b> → '
        f'Meta Final: <b>{META_CERDOS_MES} cerdos/mes</b> (<b>${META_UTIL_CERDO:,.0f}/cerdo</b>)</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4 = st.columns(4)
    total_cerdos_engorde = int(to_num(lotes_activos["Cant_Actual"]).sum()) if not lotes_activos.empty else 0
    c1.metric("🐷 Cerdas Madres Activas", madres_activas_tot, delta="9 madres = 20/mes | 18 madres = 40/mes")
    c2.metric("🐖 Cerdos en Crecimiento", total_cerdos_engorde)
    if es_admin:
        c3.metric("🎯 Ventas (0–30d) vs Escalón 1", f"{cerdos_vendidos_30d} / {META_INTERMEDIA_MES}", delta=f"Meta final: {META_CERDOS_MES}/mes")
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
# MÓDULO 2: SIMULADOR INTERACTIVO Y PROYECCIÓN REALISTA (SOLO ADMIN)
# ==============================================================================
elif menu == "🔮 Simulador y Proyección a Metas":
    st.markdown('<div class="main-header">🔮 Simulador Realista de Crecimiento y Diagnóstico</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">Proyección simplificada: las madres actuales se consideran en régimen productivo; las madres añadidas se incorporan con rezago de 9 meses. No reemplaza un modelo individual de servicios, partos, destetes y mortalidad.</div>',
        unsafe_allow_html=True
    )

    col_m1, col_m2, col_m3 = st.columns([1.2, 1, 1])
    modo_meta = col_m1.selectbox(
        "🎯 Objetivo a Simular:",
        [f"Primer Escalón ({META_INTERMEDIA_MES} cerdos/mes)", f"Meta Final ({META_CERDOS_MES} cerdos/mes)", "Meta Personalizada"]
    )
    default_obj = META_INTERMEDIA_MES if "Primer Escalón" in modo_meta else (META_CERDOS_MES if "Meta Final" in modo_meta else 20)
    objetivo_cerdos_sim = col_m2.number_input("Cerdos Vendidos / Mes Objetivo", min_value=2, max_value=100, value=int(default_obj), step=2)
    objetivo_util_cerdo_sim = col_m3.number_input("Utilidad Neta Meta / Cerdo ($)", min_value=50000.0, max_value=500000.0, value=float(META_UTIL_CERDO), step=10000.0)

    st.divider()
    s1, s2, s3 = st.columns(3)
    madres_base_sim = s1.number_input("Cerdas Madres Activas Hoy", min_value=1, max_value=50, value=max(int(madres_activas_tot), 2), step=1)
    estrategia_crec = s2.selectbox(
        "Estrategia de Crecimiento Pie de Cría",
        [
            "Acelerada: Comprar cerdas listas ($1.8M c/u)",
            "Mixta: Compra inicial + retención de primerizas propias",
            "Orgánica: Solo retener primerizas propias (sin compra)"
        ]
    )
    cerdas_comprar_ya = s3.slider("Cerdas listas a comprar en Mes 1", min_value=0, max_value=20, value=4 if "Acelerada" in estrategia_crec else (2 if "Mixta" in estrategia_crec else 0), step=1)

    s4, s5, s6 = st.columns(3)
    ajuste_costo_cuido_pct = s4.slider("Variación Costo Concentrado (%)", min_value=-15, max_value=15, value=0, step=1)
    bultos_cuido_externo_mes = s5.slider("Venta Cuido a Terceros (bultos/mes)", min_value=0, max_value=150, value=25, step=5)
    aporte_socio_sim = s6.number_input("Aporte Capital Extra Socio ($)", min_value=0.0, max_value=150000000.0, value=0.0, step=2000000.0)

    COSTO_CERDA_REEMPLAZO = 1800000.0
    cerdos_mes_por_madre = 2.25
    madres_necesarias_meta = int(np.ceil(objetivo_cerdos_sim / cerdos_mes_por_madre))
    madres_faltantes = max(madres_necesarias_meta - madres_base_sim, 0)

    costo_base_cerdo = costo_prod_prom_cerdo_hist if costo_prod_prom_cerdo_hist > 0 else 720000.0
    ahorro_cuido_cerdo = (costo_base_cerdo * 0.65) * (-ajuste_costo_cuido_pct / 100.0)
    util_base_cerdo = utilidad_prom_por_cerdo_hist if utilidad_prom_por_cerdo_hist > 0 else 280000.0
    util_cerdo_simulada = util_base_cerdo + ahorro_cuido_cerdo
    util_bulto_cuido_sim = 13500.0

    filas_sim = []
    madres_act_m = float(madres_base_sim)
    # Desembolso real de compra inicial en Mes 1
    inversion_inicial_madres = cerdas_comprar_ya * COSTO_CERDA_REEMPLAZO
    caja_sim = float(aporte_socio_sim) - inversion_inicial_madres
    caja_minima = caja_sim
    mes_logro_meta = None

    # Historial mensual de madres activas para aplicar rezago aproximado de 9 meses (gestación + engorde)
    historial_madres = [float(madres_base_sim)]

    for m in range(0, 25):
        egreso_madres_m = 0.0
        if m == 1 and cerdas_comprar_ya > 0:
            madres_act_m = min(madres_act_m + cerdas_comprar_ya, madres_necesarias_meta)
            egreso_madres_m = inversion_inicial_madres
        elif m >= 2 and madres_act_m < madres_necesarias_meta:
            if "Acelerada" in estrategia_crec:
                # Solo compra otra cerda si hay flujo positivo real en caja
                if m % 2 == 0 and caja_sim >= COSTO_CERDA_REEMPLAZO + 2500000:
                    madres_act_m = min(madres_act_m + 1, madres_necesarias_meta)
                    caja_sim -= COSTO_CERDA_REEMPLAZO
                    egreso_madres_m = COSTO_CERDA_REEMPLAZO
            elif "Mixta" in estrategia_crec:
                # Primerizas retenidas tardan 7 meses en entrar a servicio (costo crianza ~$650.000)
                if m >= 7 and m % 2 == 1 and caja_sim >= 1500000:
                    madres_act_m = min(madres_act_m + 1, madres_necesarias_meta)
                    caja_sim -= 650000.0
                    egreso_madres_m = 650000.0
            else:
                # Orgánica pura: solo retiene si hubo partos previos suficientes
                if m >= 7 and historial_madres[m - 7] >= 2 and caja_sim >= 1500000:
                    madres_act_m = min(madres_act_m + 1, madres_necesarias_meta)
                    caja_sim -= 650000.0
                    egreso_madres_m = 650000.0

        historial_madres.append(madres_act_m)

        # Las ventas del mes m provienen de las madres que estaban en servicio 9 meses atrás
        madres_que_producen_venta = historial_madres[max(0, m - 9)]
        cerdos_vendidos_m = int(round(madres_que_producen_venta * cerdos_mes_por_madre))

        if cerdos_vendidos_m >= objetivo_cerdos_sim and mes_logro_meta is None:
            mes_logro_meta = m

        util_marranos_m = cerdos_vendidos_m * util_cerdo_simulada
        util_cuido_m = bultos_cuido_externo_mes * util_bulto_cuido_sim
        gasto_fijo_m = 2200000.0 if cerdos_vendidos_m <= 22 else 3400000.0
        flujo_neto_m = util_marranos_m + util_cuido_m - gasto_fijo_m
        caja_sim += flujo_neto_m
        caja_minima = min(caja_minima, caja_sim)

        filas_sim.append({
            "Mes": f"Mes {m:02d}",
            "Madres Activas": int(madres_act_m),
            "Cerdos Vendidos / Mes": cerdos_vendidos_m,
            "Meta Objetivo": objetivo_cerdos_sim,
            "Inversión Reemplazos ($)": round(egreso_madres_m, 0),
            "Flujo Operativo Mensual ($)": round(flujo_neto_m, 0),
            "Caja Acumulada ($)": round(caja_sim, 0)
        })

    df_sim = pd.DataFrame(filas_sim)

    st.divider()
    k1, k2, k3, k4 = st.columns(4)
    tiempo_txt = f"En {mes_logro_meta} meses" if mes_logro_meta is not None else "> 24 meses (falta caja/madres)"
    k1.metric("⏱️ Tiempo Real a Meta", tiempo_txt, delta=f"Meta: {objetivo_cerdos_sim} cerdos/mes")
    k2.metric("🐷 Madres Necesarias", f"{madres_necesarias_meta} cerdas", delta=f"{madres_faltantes:+d} vs base ({madres_base_sim})")
    k3.metric("💵 Utilidad / Cerdo Simulada", f"${util_cerdo_simulada:,.0f}", delta=f"{util_cerdo_simulada - objetivo_util_cerdo_sim:+,.0f} vs meta")
    k4.metric("🏦 Punto Mínimo de Caja", f"${caja_minima:,.0f}", delta="Deficitario (requiere aporte)" if caja_minima < 0 else "Caja positiva")

    fig_sim = go.Figure()
    fig_sim.add_trace(go.Scatter(x=df_sim["Mes"], y=df_sim["Cerdos Vendidos / Mes"], mode="lines+markers", name="Cerdos Vendidos / Mes", line=dict(color="#1E8449", width=3)))
    fig_sim.add_trace(go.Scatter(x=df_sim["Mes"], y=df_sim["Meta Objetivo"], mode="lines", name=f"Meta ({objetivo_cerdos_sim} cerdos/mes)", line=dict(color="#D35400", width=2, dash="dash")))
    fig_sim.update_layout(title=f"Curva Biológica y Financiera hacia {objetivo_cerdos_sim} Cerdos/Mes (24 Meses)", yaxis_title="Cerdos Vendidos / Mes", height=380)
    st.plotly_chart(fig_sim, use_container_width=True)

    st.subheader("🧠 Diagnóstico Automático")
    if caja_minima < 0:
        st.error(
            f"🚨 **Déficit de Caja Detectado (${abs(caja_minima):,.0f}):** Comprar {cerdas_comprar_ya} cerdas y sostener el concentrado mientras llegan las primeras ventas "
            f"deja la caja en negativo. **Recomendación:** Realizar un aporte de capital de al menos **${abs(caja_minima) + 2000000:,.0f}** o reducir la compra inicial de madres."
        )
    else:
        st.success("✅ **Liquidez Sostenible:** La caja se mantiene positiva durante toda la curva de crecimiento hacia la meta.")

    with st.expander("📋 Ver Tabla Detallada Mes a Mes (24 Meses)"):
        st.dataframe(df_sim, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 3: SALUD FINANCIERA Y CONCILIACIÓN REAL DE CAJA (SOLO ADMIN)
# ==============================================================================
elif menu == "💼 Salud Financiera y Reinversión":
    st.markdown('<div class="main-header">💼 Salud Financiera y Flujo de Caja Registrado</div>', unsafe_allow_html=True)

    ing_marranos, util_marranos = 0.0, 0.0
    ing_descarte, util_descarte = 0.0, 0.0
    ing_cuido, util_cuido = 0.0, 0.0
    ing_otros, util_otros = 0.0, 0.0
    ingresos_brutos_totales = 0.0
    fletes_sacrificio_totales = 0.0

    if not df_ventas.empty:
        for _, v in df_ventas.iterrows():
            linea = str(v.get("Linea_Negocio", "Venta Cerdos Cebados"))
            ib = float(to_num(pd.Series([v["Ingreso_Bruto"]])).iloc[0])
            un = float(to_num(pd.Series([v["Utilidad_Neta"]])).iloc[0])
            fl = float(to_num(pd.Series([v["Costo_Sacrificio_O_Flete"]])).iloc[0])
            ingresos_brutos_totales += ib
            fletes_sacrificio_totales += fl
            if "Concentrado" in linea or "Cuido" in linea:
                ing_cuido += ib
                util_cuido += un
            elif "Descarte" in linea:
                ing_descarte += ib
                util_descarte += un
            elif "Cebados" in linea:
                ing_marranos += ib
                util_marranos += un
            else:
                ing_otros += ib
                util_otros += un

    util_bruta_ventas = util_marranos + util_descarte + util_cuido + util_otros

    aportes_socio = 0.0
    capex_total = 0.0
    gastos_fijos_tot = 0.0
    reinversion_registrada = 0.0
    retiros_ganancia = 0.0

    if not df_inv.empty:
        for _, m in df_inv.iterrows():
            t_mov = str(m["Tipo_Movimiento"]).strip()
            monto = float(to_num(pd.Series([m["Monto"]])).iloc[0])
            if t_mov == "Aporte de Capital (Socio)":
                aportes_socio += monto
            elif t_mov == "Inversión Infraestructura / Activos":
                capex_total += monto
            elif t_mov == "Reinversión en Inventario / Operaciones":
                reinversion_registrada += monto
            elif t_mov == "Gasto Fijo Operativo":
                gastos_fijos_tot += monto
            elif t_mov == "Retiro de Utilidades (Ganancia Socio)":
                retiros_ganancia += monto

    # Desembolsos reales en compras de bodega y servicios reproductivos
    compras_bodega_tot = float((to_num(df_bodega["Unidades_Compradas"]) * to_num(df_bodega["Precio_Unitario"])).sum()) if not df_bodega.empty else 0.0
    servicios_ia_tot = float(to_num(df_repro["Costo_Servicio"]).sum()) if not df_repro.empty else 0.0
    valor_inventario_actual = float((to_num(df_bodega["Saldo_Disponible"]) * to_num(df_bodega["Costo_Por_Medida"])).sum()) if not df_bodega.empty else 0.0

    # Ecuación real de flujo de efectivo (sin doble conteo de reinversiones)
    entradas_efectivo = aportes_socio + ingresos_brutos_totales
    salidas_efectivo = capex_total + compras_bodega_tot + servicios_ia_tot + fletes_sacrificio_totales + gastos_fijos_tot + retiros_ganancia
    caja_real_disponible = entradas_efectivo - salidas_efectivo

    # Utilidad contable neta integral (descontando gastos fijos operativos)
    utilidad_contable_neta = util_bruta_ventas - gastos_fijos_tot
    tot_cerdos_v = float(to_num(df_ventas[df_ventas["Linea_Negocio"].astype(str).str.contains("Cebados", na=False)]["Cantidad_Animales_O_Unidades"]).sum()) if not df_ventas.empty else 0.0
    util_integral_por_cerdo = (utilidad_contable_neta / tot_cerdos_v) if tot_cerdos_v > 0 else 0.0

    st.subheader("💵 Flujo de Caja Registrado (Estimado) vs. Rentabilidad Registrada")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🏦 Caja Registrada Estimada", f"${caja_real_disponible:,.0f}", delta="Según movimientos registrados")
    k2.metric("📦 Valor Inventario Bodega", f"${valor_inventario_actual:,.0f}")
    k3.metric("📈 Utilidad Contable Neta", f"${utilidad_contable_neta:,.0f}", delta=f"Margen ventas - OPEX (${gastos_fijos_tot:,.0f})")
    k4.metric("🐖 Utilidad Integral / Cerdo", f"${util_integral_por_cerdo:,.0f}", delta=f"Margen lote: ${utilidad_prom_por_cerdo_hist:,.0f}")

    reserva_capital_trabajo = 4500000.0
    if caja_real_disponible < reserva_capital_trabajo:
        st.error(
            f"🔴 **PRIORIDAD LIQUIDEZ (NO RETIRAR UTILIDADES):** La caja real estimada (**${caja_real_disponible:,.0f}**) está por debajo del colchón operativo mínimo "
            f"(**${reserva_capital_trabajo:,.0f}**). Mantén los recursos en capital de trabajo."
        )
    elif utilidad_contable_neta < capex_total:
        st.warning(
            f"🟡 **ETAPA DE AMORTIZACIÓN DE INVERSIÓN:** Tienes buena liquidez en caja (**${caja_real_disponible:,.0f}**), pero la utilidad contable acumulada "
            f"(**${utilidad_contable_neta:,.0f}**) aún amortiza el CAPEX inicial (**${capex_total:,.0f}**)."
        )
    else:
        excedente = max(caja_real_disponible - reserva_capital_trabajo, 0.0)
        st.success(f"🟢 **ESTABILIDAD FINANCIERA:** Inversión amortizada y liquidez cubierta. Excedente distribuible: **${excedente:,.0f}**.")

    st.divider()
    df_lineas = pd.DataFrame([
        {"Línea de Ingreso": "🐖 Venta Cerdos Cebados", "Ingresos Brutos ($)": ing_marranos, "Margen Bruto Lote ($)": util_marranos},
        {"Línea de Ingreso": "🐷 Venta Cerdas Descarte", "Ingresos Brutos ($)": ing_descarte, "Margen Bruto Lote ($)": util_descarte},
        {"Línea de Ingreso": "🌾 Distribución Cuido", "Ingresos Brutos ($)": ing_cuido, "Margen Bruto Lote ($)": util_cuido},
        {"Línea de Ingreso": "🍼 Lechones / Otros", "Ingresos Brutos ($)": ing_otros, "Margen Bruto Lote ($)": util_otros}
    ])
    st.dataframe(df_lineas, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("📝 Registrar Movimiento de Capital, Gasto Fijo o Retiro")
    c1, c2 = st.columns(2)
    tipo_mov = c1.selectbox("Tipo de Movimiento", [
        "Gasto Fijo Operativo",
        "Inversión Infraestructura / Activos",
        "Reinversión en Inventario / Operaciones",
        "Retiro de Utilidades (Ganancia Socio)",
        "Aporte de Capital (Socio)"
    ])
    origen_f = c2.selectbox("Origen de los Fondos", [
        "Caja Operativa Granja",
        "Aporte Directo Socio / Capital Propio",
        "Crédito / Externo"
    ])
    cat_mov = c1.text_input("Categoría", value="Nómina / Mano de Obra" if "Gasto" in tipo_mov else "Movimiento General")
    monto_mov = c2.number_input("Monto ($)", min_value=0.0, value=500000.0, step=50000.0)
    desc_mov = c1.text_input("Descripción Detallada", placeholder="Ej: Pago nómina operario mes actual")
    resp_mov = c2.text_input("Responsable", value="Administración")
    obs_mov = st.text_input("Observaciones", value="")

    if st.button("💾 Guardar Movimiento Financiero"):
        nuevo_id_fin = generar_siguiente_id(df_inv, "ID_Movimiento", "FIN")
        nuevo_m = pd.DataFrame([{
            "ID_Movimiento": nuevo_id_fin,
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
        st.success(f"Movimiento {nuevo_id_fin} registrado correctamente.")
        st.rerun()

    st.dataframe(df_inv, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 4: ESCALONAMIENTO Y BANDAS (RESPETANDO CAPACIDAD ESTRICTA)
# ==============================================================================
elif menu == "📅 Escalonamiento y Bandas":
    st.markdown('<div class="main-header">📅 Planificador de Escalonamiento por Bandas</div>', unsafe_allow_html=True)
    ciclo_biologico = 114 + EDAD_DESTETE + 5
    ciclo_bandas = NUM_BANDAS * INTERVALO_BANDAS
    capacidad_total_granja = NUM_BANDAS * CUPO_BANDA

    st.info(
        f"📐 **Sincronización Biológica vs. Bandas:** Ciclo cerda = **{ciclo_biologico} días** (114d gest + {EDAD_DESTETE}d lact + 5d celo) | "
        f"Ciclo de {NUM_BANDAS} bandas cada {INTERVALO_BANDAS}d = **{ciclo_bandas} días** (margen de **{ciclo_bandas - ciclo_biologico:+d} días** para absorción de retornos a celo). "
        f"Capacidad máxima instalada: **{capacidad_total_granja} madres**."
    )

    if madres_activas_tot > capacidad_total_granja:
        st.warning(f"⚠️ Tienes **{madres_activas_tot} madres activas**, lo cual supera la capacidad instalada de **{capacidad_total_granja} cupos** ({NUM_BANDAS} bandas × {CUPO_BANDA} cerdas/banda).")

    lista_bandas = [f"Banda {i:02d}" for i in range(1, NUM_BANDAS + 1)]
    tab1, tab2, tab3 = st.tabs(["📊 Distribución Actual de Bandas", "🗓️ Cronograma Maestro", "➕ Asignar / Reorganizar Cerdas"])

    with tab1:
        if hembras_activas_df.empty:
            st.info("No hay hembras activas registradas.")
        else:
            hembras_activas_df["Banda_Limpia"] = hembras_activas_df["Banda"].astype(str).str.strip().replace("", "Sin Asignar")
            resumen_bandas = []
            for b in lista_bandas:
                sub = hembras_activas_df[hembras_activas_df["Banda_Limpia"] == b]
                cupo_disp = CUPO_BANDA - len(sub)
                estado_cupo = "🟢 Disponible" if cupo_disp > 0 else ("✅ Completa" if cupo_disp == 0 else "🔴 Sobrecupo")
                resumen_bandas.append({
                    "Banda": b,
                    "Cerdas Asignadas": len(sub),
                    "Cupo Máximo": CUPO_BANDA,
                    "Espacios Libres": max(cupo_disp, 0),
                    "Estado Banda": estado_cupo,
                    "IDs Cerdas": ", ".join(opciones_limpias(sub["ID_Animal"])) if not sub.empty else "—"
                })
            st.dataframe(pd.DataFrame(resumen_bandas), use_container_width=True, hide_index=True)

            sin_b = hembras_activas_df[~hembras_activas_df["Banda_Limpia"].isin(lista_bandas)]
            if not sin_b.empty:
                st.warning(f"⏳ Cerdas en espera / sin banda asignada: **{', '.join(opciones_limpias(sin_b['ID_Animal']))}**")

    with tab2:
        fecha_inicio_ciclo = st.date_input("Fecha de inicio de la Banda 01", value=hoy)
        filas_cron = []
        for idx, b in enumerate(lista_bandas):
            f_ins = fecha_inicio_ciclo + timedelta(days=idx * INTERVALO_BANDAS)
            f_par = f_ins + timedelta(days=114)
            f_des = f_par + timedelta(days=EDAD_DESTETE)
            f_reins = f_des + timedelta(days=5)
            f_venta = f_par + timedelta(days=EDAD_META)
            cerdas_b = opciones_limpias(hembras_activas_df[hembras_activas_df["Banda"].astype(str).str.strip() == b]["ID_Animal"]) if not hembras_activas_df.empty else []
            filas_cron.append({
                "Banda": b,
                "Cerdas": ", ".join(cerdas_b) if cerdas_b else "(Cupo libre)",
                "Inseminación": f_ins.strftime("%Y-%m-%d"),
                "Parto (114d)": f_par.strftime("%Y-%m-%d"),
                f"Destete ({EDAD_DESTETE}d)": f_des.strftime("%Y-%m-%d"),
                "Re-Servicio (+5d)": f_reins.strftime("%Y-%m-%d"),
                f"Salida 120kg ({EDAD_META}d)": f_venta.strftime("%Y-%m-%d")
            })
        st.dataframe(pd.DataFrame(filas_cron), use_container_width=True, hide_index=True)

    with tab3:
        lista_h_ids = opciones_limpias(hembras_activas_df["ID_Animal"]) if not hembras_activas_df.empty else []
        if not lista_h_ids:
            st.info("Agrega cerdas primero.")
        else:
            col_izq, col_der = st.columns(2)
            with col_izq:
                cerda_sel = st.selectbox("Seleccionar Cerda", lista_h_ids)
                banda_nueva = st.selectbox("Asignar a Banda", lista_bandas + ["Sin Asignar"])
                ocupadas_en_banda = len(hembras_activas_df[(hembras_activas_df["Banda"].astype(str).str.strip() == banda_nueva) & (hembras_activas_df["ID_Animal"].astype(str).str.strip() != cerda_sel)])
                sobrecupo = (banda_nueva != "Sin Asignar" and ocupadas_en_banda >= CUPO_BANDA)
                if sobrecupo:
                    st.error(f"🚫 La {banda_nueva} ya alcanzó su cupo máximo ({CUPO_BANDA} cerdas).")
                if st.button("💾 Guardar Banda", disabled=sobrecupo):
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_sel, "Banda"] = banda_nueva
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{cerda_sel} asignada a {banda_nueva}.")
                    st.rerun()

            with col_der:
                st.write(f"Distribuye respetando estrictamente el cupo máximo de **{CUPO_BANDA} cerdas por banda** (máx. {capacidad_total_granja} cerdas).")
                if st.button("🪄 Distribuir Respetando Cupo Máximo"):
                    idx_h = hembras_activas_df.index.tolist()
                    asignadas = 0
                    for i, row_idx in enumerate(idx_h):
                        if asignadas < capacidad_total_granja:
                            df_rep.at[row_idx, "Banda"] = lista_bandas[asignadas % NUM_BANDAS]
                            asignadas += 1
                        else:
                            df_rep.at[row_idx, "Banda"] = "Sin Asignar"
                    guardar_hoja("Reproductores", df_rep)
                    if len(idx_h) > capacidad_total_granja:
                        st.warning(f"Se asignaron {capacidad_total_granja} cerdas al cupo máximo y {len(idx_h) - capacidad_total_granja} quedaron en 'Sin Asignar'.")
                    else:
                        st.success(f"Distribución completada respetando el cupo de {CUPO_BANDA} cerdas/banda.")
                    st.rerun()

# ==============================================================================
# MÓDULO 5: REPRODUCCIÓN (IA Y PARTOS) CON VALIDACIÓN ESTRICTA DE ESTADOS
# ==============================================================================
elif menu == "🧬 Reproducción (IA y Partos)":
    st.markdown('<div class="main-header">🧬 Gestión Reproductiva (Inseminación Artificial)</div>', unsafe_allow_html=True)

    # SOLO cerdas en estado 'Vacía' y SIN ningún evento reproductivo abierto
    cerdas_elegibles_ia_df = hembras_activas_df[
        (hembras_activas_df["Estado"].astype(str).str.strip() == "Vacía") &
        (~hembras_activas_df["ID_Animal"].astype(str).str.strip().isin(cerdas_con_ciclo_abierto))
    ] if not hembras_activas_df.empty else pd.DataFrame()

    lista_hembras_ia = opciones_limpias(cerdas_elegibles_ia_df["ID_Animal"]) if not cerdas_elegibles_ia_df.empty else []
    verracos_activos = df_rep[(df_rep["Tipo"].astype(str).str.contains("Verraco", na=False)) & (~df_rep["Estado"].astype(str).str.strip().isin(ESTADOS_INACTIVOS_REP))] if not df_rep.empty else pd.DataFrame()
    verracos_lista = opciones_limpias(verracos_activos["ID_Animal"]) if not verracos_activos.empty else []
    opciones_recelador = [f"Sí ({v})" for v in verracos_lista] + ["Sí (Verraco Externo)", "No"] if verracos_lista else ["Sí (Verraco Externo)", "No"]

    t1, t2, t3, t4 = st.tabs(["1️⃣ Registrar Inseminación", "2️⃣ Confirmar Preñez / Falla", "3️⃣ Registrar Parto", f"4️⃣ Registrar Destete ({EDAD_DESTETE}d)"])

    with t1:
        if not lista_hembras_ia:
            st.warning("⚠️ No hay cerdas en estado **Vacía** (sin ciclo abierto) elegibles para inseminar.")
        else:
            c1, c2 = st.columns(2)
            cerda_ia = c1.selectbox("Cerda Elegible a Inseminar (Solo Vacías)", lista_hembras_ia)
            info_c = cerdas_elegibles_ia_df[cerdas_elegibles_ia_df["ID_Animal"].astype(str).str.strip() == cerda_ia].iloc[0]
            banda_c = str(info_c.get("Banda", "Banda 01")).strip() or "Banda 01"
            raza_c = str(info_c["Linea_Genetica"]).strip()
            st.caption(f"Genética **{cerda_ia}**: `{raza_c}` | Banda: `{banda_c}` | Estado: `Vacía`")

            f_serv = c2.date_input("Fecha de Inseminación", value=hoy)
            semen_raza = c1.selectbox("Línea / Raza del Semen (Pajilla)", RAZAS_SEMEN)
            recelador = c2.selectbox("Estímulo con Verraco Recelador", opciones_recelador)
            costo_ia = c1.number_input("Costo Total Servicio / Pajillas ($)", min_value=0.0, value=45000.0, step=5000.0) if es_admin else 45000.0

            f_parto_est = f_serv + timedelta(days=114)
            if st.button("💾 Guardar Inseminación"):
                # Doble verificación contra concurrencia
                if cerda_ia in cerdas_con_ciclo_abierto:
                    st.error(f"🚫 {cerda_ia} ya tiene un ciclo reproductivo abierto.")
                else:
                    nuevo_id = generar_siguiente_id(df_repro, "ID_Evento", "INS")
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
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_ia, "Estado"] = "Inseminada"
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_ia, "Fecha_Ultimo_Cambio_Estado"] = f_serv.strftime("%Y-%m-%d")
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_ia, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_ia, "Costo_Acumulado_Ciclo"]) + costo_ia
                    guardar_hoja("Reproduccion", df_repro)
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"Servicio {nuevo_id} registrado para {cerda_ia}. Parto probable: {f_parto_est}.")
                    st.rerun()

    with t2:
        pendientes = df_repro[df_repro["Estado_Prenez"].astype(str).str.strip().isin(["Pendiente Eco", "Preñada"])] if not df_repro.empty else pd.DataFrame()
        if pendientes.empty:
            st.info("No hay servicios pendientes de diagnóstico.")
        else:
            opciones_ev = (pendientes["ID_Evento"].astype(str).str.strip() + " | " + pendientes["ID_Cerda"].astype(str).str.strip() + " [" + pendientes["Estado_Prenez"].astype(str).str.strip() + "]").tolist()
            ev_sel = st.selectbox("Seleccionar Evento", opciones_ev)
            id_ev = ev_sel.split(" | ")[0].strip()
            fila_ev_sel = pendientes[pendientes["ID_Evento"].astype(str).str.strip() == id_ev].iloc[0]
            est_previo = str(fila_ev_sel["Estado_Prenez"]).strip()

            if est_previo == "Pendiente Eco":
                res_eco = st.selectbox("Resultado Ecografía / Chequeo", ["Preñada", "Repitió Celo"])
                confirmar_cambio = True
            else:
                st.warning("⚠️ Esta cerda ya tiene **Preñez Confirmada**. Solo modifica este estado si ocurrió un aborto o pérdida gestacional.")
                res_eco = st.selectbox("Registrar Novedad en Preñez", ["Aborto", "Repitió Celo"])
                confirmar_cambio = st.checkbox("Confirmo cerrar esta preñez confirmada por novedad sanitaria/reproductiva")

            if st.button("💾 Actualizar Estado Reproductivo", disabled=not confirmar_cambio):
                cerda_id = str(fila_ev_sel["ID_Cerda"]).strip()
                df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev, "Estado_Prenez"] = res_eco
                if res_eco == "Preñada":
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Estado"] = "Preñada"
                else:
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Estado"] = "Vacía"
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")
                    if res_eco == "Repitió Celo":
                        df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Repeticiones"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Repeticiones"]) + 1
                    elif res_eco == "Aborto":
                        df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Abortos"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Abortos"]) + 1
                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Reproductores", df_rep)
                st.success("Estado reproductivo actualizado.")
                st.rerun()

    with t3:
        prenadas = df_repro[df_repro["Estado_Prenez"].astype(str).str.strip() == "Preñada"] if not df_repro.empty else pd.DataFrame()
        if prenadas.empty:
            st.info("No hay cerdas preñadas pendientes de parto.")
        else:
            opciones_p = (prenadas["ID_Evento"].astype(str).str.strip() + " | " + prenadas["ID_Cerda"].astype(str).str.strip() + " (" + prenadas["Macho_O_Semen"].astype(str).str.strip() + ")").tolist()
            ev_p = st.selectbox("Cerda que Parió", opciones_p)
            id_ev_p = ev_p.split(" | ")[0].strip()
            fila_p = prenadas[prenadas["ID_Evento"].astype(str).str.strip() == id_ev_p].iloc[0]
            c1, c2, c3 = st.columns(3)
            f_parto_r = c1.date_input("Fecha Real de Parto", value=hoy)
            vivos = c2.number_input("Nacidos Vivos", min_value=1, max_value=25, value=12)
            mortinatos = c3.number_input("Mortinatos / Momias", min_value=0, max_value=10, value=0)
            sig_lote_num = len(df_lotes) + 1
            nombre_lote = st.text_input("Nombre del Nuevo Lote", value=f"LOTE-{sig_lote_num:02d}-{str(fila_p['ID_Cerda']).strip()}")
            corral_parto = st.text_input("Corral / Paridera", value=f"Paridera-{str(fila_p['ID_Cerda']).strip()}")

            if st.button("💾 Registrar Parto y Crear Lote Lactante"):
                if nombre_lote.strip() in opciones_limpias(df_lotes["ID_Lote"]):
                    st.error("Ya existe un lote con ese nombre. Elige otro identificador.")
                else:
                    cerda_id = str(fila_p["ID_Cerda"]).strip()
                    costo_ciclo = float(to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Costo_Acumulado_Ciclo"]).iloc[0])
                    cruce_txt = f"{fila_p['Raza_Cerda']} × {fila_p['Macho_O_Semen']}"

                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Estado_Prenez"] = "En Lactancia"
                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Fecha_Real_Parto"] = f_parto_r.strftime("%Y-%m-%d")
                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Nacidos_Vivos"] = vivos
                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Mortinatos"] = mortinatos
                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Nombre_Lote_Asignado"] = nombre_lote.strip()
                    df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_p, "Fecha_Prob_Destete"] = (f_parto_r + timedelta(days=EDAD_DESTETE)).strftime("%Y-%m-%d")

                    nuevo_lote = pd.DataFrame([{
                        "ID_Lote": nombre_lote.strip(),
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

                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Estado"] = "Lactante"
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Partos"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Partos"]) + 1
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Costo_Acumulado_Ciclo"] = 0.0
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = f_parto_r.strftime("%Y-%m-%d")

                    guardar_hoja("Reproduccion", df_repro)
                    guardar_hoja("Lotes_Engorde", df_lotes)
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"Parto registrado. Lote {nombre_lote} creado.")
                    st.rerun()

    with t4:
        lactantes = df_repro[df_repro["Estado_Prenez"].astype(str).str.strip() == "En Lactancia"] if not df_repro.empty else pd.DataFrame()
        if lactantes.empty:
            st.info("No hay camadas en lactancia.")
        else:
            opciones_d = (lactantes["ID_Evento"].astype(str).str.strip() + " | " + lactantes["ID_Cerda"].astype(str).str.strip() + " -> " + lactantes["Nombre_Lote_Asignado"].astype(str).str.strip()).tolist()
            ev_d = st.selectbox("Camada a Destetar", opciones_d)
            id_ev_d = ev_d.split(" | ")[0].strip()
            fila_d = lactantes[lactantes["ID_Evento"].astype(str).str.strip() == id_ev_d].iloc[0]
            c1, c2 = st.columns(2)
            dest_reales = c1.number_input("Lechones Destetados", min_value=1, max_value=25, value=int(to_num(pd.Series([fila_d["Nacidos_Vivos"]])).iloc[0]))
            peso_tot_dest = c2.number_input("Peso Total Camada Destetada (kg)", min_value=1.0, value=float(round(dest_reales * PESO_DEST_21D, 1)))
            corral_dest = st.text_input("Corral de Precebo (Fase 0)", value="Precebo-1")

            if st.button("💾 Confirmar Destete e Iniciar Fase 0"):
                cerda_id = str(fila_d["ID_Cerda"]).strip()
                lote_id = str(fila_d["Nombre_Lote_Asignado"]).strip()
                costo_lact = float(to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Costo_Acumulado_Ciclo"]).iloc[0])

                df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_d, "Estado_Prenez"] = "Destetada"
                df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_d, "Destetados_Reales"] = dest_reales
                df_repro.loc[df_repro["ID_Evento"].astype(str).str.strip() == id_ev_d, "Peso_Destete_Total_Kg"] = peso_tot_dest

                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Etapa"] = "Fase 0"
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Corral"] = corral_dest
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Cant_Actual"] = dest_reales
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Peso_Prom_Actual_Kg"] = round(peso_tot_dest / dest_reales, 2)
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Costo_Inicial_Lechones"] = to_num(df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_id, "Costo_Inicial_Lechones"]) + costo_lact

                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Estado"] = "Vacía"
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Costo_Acumulado_Ciclo"] = 0.0
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_id, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")

                guardar_hoja("Reproduccion", df_repro)
                guardar_hoja("Lotes_Engorde", df_lotes)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Destete completado. {lote_id} entra a Fase 0.")
                st.rerun()

    st.divider()
    st.subheader("Historial Reproductivo")
    st.dataframe(df_repro, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 6: LOTES, CURVA DINÁMICA Y FUSIÓN CON TRAZABILIDAD COMPLETA
# ==============================================================================
elif menu == "🐖 Lotes, Curva y Fases":
    st.markdown('<div class="main-header">🐖 Control Productivo, Curva Dinámica y Fases Nutricionales</div>', unsafe_allow_html=True)

    with st.expander(f"📋 Tabla Dinámica de Fases (Destete: {EDAD_DESTETE}d -> Meta {PESO_META:.0f} kg a los {EDAD_META} días)"):
        df_prog_vis = pd.DataFrame(PROGRAMA_FASES)
        df_prog_vis["Acumulado_Kg"] = df_prog_vis["kg_cerdo"].cumsum()
        df_prog_vis.columns = ["Fase", "Edad Inicio (d)", "Edad Fin (d)", "Peso Inicio (kg)", "Peso Fin (kg)", "Kg / Cerdo", "Siguiente Fase", "Kg Acumulados"]
        st.dataframe(df_prog_vis, use_container_width=True, hide_index=True)

    if lotes_activos.empty:
        st.info("No hay lotes activos.")
    else:
        filas_ind = []
        for _, l in lotes_activos.iterrows():
            lote_id_str = str(l["ID_Lote"]).strip()
            ev_f = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
            edad_dias = ev_f["edad_actual"]
            cant_ini = int(to_num(pd.Series([l["Cant_Inicial"]])).iloc[0])
            cant_act = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
            p_ini = float(to_num(pd.Series([l["Peso_Prom_Inicial_Kg"]])).iloc[0])
            p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
            kg_alim = float(to_num(pd.Series([l["Kg_Alimento_Consumido"]])).iloc[0])

            c_ini = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0])
            c_ali = float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0])
            c_san = float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
            c_tot = c_ini + c_ali + c_san

            dias_trans = max(edad_dias - int(to_num(pd.Series([l.get("Edad_Inicial_Dias", 0)])).iloc[0]), 1)
            gdp_real_g = ((p_act - p_ini) / dias_trans) * 1000.0
            peso_esp_hoy = peso_esperado_por_edad(edad_dias, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE, PROGRAMA_FASES)
            cumpl_pct = (p_act / max(peso_esp_hoy, 0.1)) * 100.0

            # Conversión real considerando biomasa vendida parcialmente de este lote + biomasa viva actual - biomasa inicial
            kg_vendidos_parcial = 0.0
            if not df_ventas.empty:
                v_parc = df_ventas[df_ventas["ID_Origen"].astype(str).str.strip() == lote_id_str]
                kg_vendidos_parcial = float(to_num(v_parc["Peso_Pie_O_Kilos"]).sum())

            biomasa_salida_y_viva = (p_act * cant_act) + kg_vendidos_parcial
            biomasa_inicial = p_ini * max(cant_ini, cant_act)
            kg_ganados_reales = max(biomasa_salida_y_viva - biomasa_inicial, 0.1)
            ca = kg_alim / kg_ganados_reales if kg_alim > 0 else 0.0
            punto_eq = c_tot / max(p_act * cant_act, 1.0)

            estado_fase_txt = (
                f"✅ Completada -> {ev_f['siguiente_fase']}"
                if (ev_f["kg_meta_lote"] > 0 and ev_f["kg_consumidos_fase"] >= ev_f["kg_meta_lote"])
                else f"{ev_f['kg_consumidos_fase']:.1f} / {ev_f['kg_meta_lote']:.1f} kg"
            )

            fila_data = {
                "Lote": lote_id_str,
                "Cerdos (Act/Ini)": f"{cant_act}/{cant_ini}",
                "Edad (d)": edad_dias,
                "Fase Reg.": ev_f["fase_registrada"],
                "Fase Edad": ev_f["fase_por_edad"],
                "Alimento Fase": estado_fase_txt,
                "Peso Real (kg)": round(p_act, 1),
                "Curva Hoy (kg)": round(peso_esp_hoy, 1),
                "Cumpl. (%)": round(cumpl_pct, 1),
                "GDP (g/d)": round(gdp_real_g, 0),
                "Conv. Alim.": round(ca, 2)
            }
            if es_admin:
                fila_data["Costo Activo ($)"] = round(c_tot, 0)
                fila_data["Punto Eq. ($/kg)"] = round(punto_eq, 0)

            filas_ind.append(fila_data)

        st.subheader("📊 Seguimiento de Lotes Activos y Trazabilidad Nutricional")
        st.dataframe(pd.DataFrame(filas_ind), use_container_width=True, hide_index=True)

        st.divider()
        c1, c2 = st.columns(2)
        lista_lotes_ids = opciones_limpias(lotes_activos["ID_Lote"])
        with c1:
            st.subheader("⚖️ Actualizar Pesaje o Avanzar de Fase")
            lote_edit = st.selectbox("Seleccionar Lote", lista_lotes_ids)
            fila_le = lotes_activos[lotes_activos["ID_Lote"].astype(str).str.strip() == lote_edit].iloc[0]
            ev_le = evaluar_fase_lote(fila_le, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
            st.caption(f"Fase registrada: **{ev_le['fase_registrada']}** | Fase por edad: **{ev_le['fase_por_edad']}** | Siguiente: **{ev_le['siguiente_fase']}**")

            nuevo_peso = st.number_input("Nuevo Peso Promedio Actual (kg)", min_value=1.0, value=float(to_num(pd.Series([fila_le["Peso_Prom_Actual_Kg"]])).iloc[0]), step=0.5)
            opciones_etapa = ["Lactancia"] + [f["fase"] for f in PROGRAMA_FASES]
            etapa_str = str(fila_le["Etapa"]).strip()
            idx_et = opciones_etapa.index(etapa_str) if etapa_str in opciones_etapa else 1
            nueva_etapa = st.selectbox("Fase / Etapa Nutricional", opciones_etapa, index=idx_et)
            nuevo_corral = st.text_input("Corral", value=str(fila_le["Corral"]))

            if st.button("💾 Guardar Pesaje / Cambio de Fase"):
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_edit, "Peso_Prom_Actual_Kg"] = nuevo_peso
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_edit, "Etapa"] = nueva_etapa
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_edit, "Corral"] = nuevo_corral
                guardar_hoja("Lotes_Engorde", df_lotes)
                st.success(f"Lote {lote_edit} actualizado.")
                st.rerun()

        with c2:
            st.subheader("🔗 Fusionar Lotes (Con Consolidación Histórica)")
            if len(lista_lotes_ids) >= 2:
                l1 = st.selectbox("Lote A (Principal)", lista_lotes_ids, key="u1")
                l2 = st.selectbox("Lote B (Se integrará al A)", [x for x in lista_lotes_ids if x != l1], key="u2")
                r1 = lotes_activos[lotes_activos["ID_Lote"].astype(str).str.strip() == l1].iloc[0]
                r2 = lotes_activos[lotes_activos["ID_Lote"].astype(str).str.strip() == l2].iloc[0]
                ev1 = evaluar_fase_lote(r1, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                ev2 = evaluar_fase_lote(r2, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                dif_edad_lotes = abs(ev1["edad_actual"] - ev2["edad_actual"])

                if dif_edad_lotes > 14:
                    st.warning(f"⚠️ Diferencia de edad alta ({dif_edad_lotes} días: {ev1['edad_actual']}d vs {ev2['edad_actual']}d). Mezclar cerdos de distinta edad puede generar competencia y desfase nutricional.")
                nuevo_nom_union = st.text_input("Nombre Lote Unificado", value=f"{l1}+{l2}")

                if st.button("🔗 Fusionar y Migrar Historiales"):
                    c1_n = int(to_num(pd.Series([r1["Cant_Actual"]])).iloc[0])
                    c2_n = int(to_num(pd.Series([r2["Cant_Actual"]])).iloc[0])
                    tot_c = c1_n + c2_n
                    p_prom = (c1_n * float(to_num(pd.Series([r1["Peso_Prom_Actual_Kg"]])).iloc[0]) + c2_n * float(to_num(pd.Series([r2["Peso_Prom_Actual_Kg"]])).iloc[0])) / max(tot_c, 1)

                    # 1. Consolidar tabla Lotes_Engorde
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == l1, "ID_Lote"] = nuevo_nom_union
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Madre_Origen"] = f"{r1['Madre_Origen']} + {r2['Madre_Origen']}"
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Raza_Cruce"] = f"Mixto ({r1['Raza_Cruce']} / {r2['Raza_Cruce']})"
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Cant_Inicial"] = int(to_num(pd.Series([r1["Cant_Inicial"]])).iloc[0]) + int(to_num(pd.Series([r2["Cant_Inicial"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Cant_Actual"] = tot_c
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Peso_Prom_Actual_Kg"] = round(p_prom, 2)
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Kg_Alimento_Consumido"] = float(to_num(pd.Series([r1["Kg_Alimento_Consumido"]])).iloc[0]) + float(to_num(pd.Series([r2["Kg_Alimento_Consumido"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Costo_Inicial_Lechones"] = float(to_num(pd.Series([r1["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Inicial_Lechones"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Costo_Alimento"] = float(to_num(pd.Series([r1["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Alimento"]])).iloc[0])
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == nuevo_nom_union, "Costo_Sanidad"] = float(to_num(pd.Series([r1["Costo_Sanidad"]])).iloc[0]) + float(to_num(pd.Series([r2["Costo_Sanidad"]])).iloc[0])
                    df_lotes = df_lotes[df_lotes["ID_Lote"].astype(str).str.strip() != l2]

                    # 2. Migrar historial completo en Consumo_Alimento y Sanidad_Eventos
                    if not df_consumo.empty:
                        df_consumo.loc[df_consumo["ID_Destino"].astype(str).str.strip().isin([l1, l2]), "ID_Destino"] = nuevo_nom_union
                    if not df_sanidad.empty:
                        df_sanidad.loc[df_sanidad["ID_Destino"].astype(str).str.strip().isin([l1, l2]), "ID_Destino"] = nuevo_nom_union

                    guardar_hoja("Lotes_Engorde", df_lotes)
                    guardar_hoja("Consumo_Alimento", df_consumo)
                    guardar_hoja("Sanidad_Eventos", df_sanidad)
                    st.success(f"Lotes unidos en {nuevo_nom_union} ({tot_c} cerdos) con trazabilidad histórica completa.")
                    st.rerun()

# ==============================================================================
# MÓDULO 7: BODEGA, VALORACIÓN PONDERADA/FIFO ÚNICA Y ARQUEO CONTABLE
# ==============================================================================
elif menu == "🌾 Bodega, Arqueo y Botiquín":
    st.markdown('<div class="main-header">🌾 Bodega, Conciliación de Inventario y Botiquín</div>', unsafe_allow_html=True)

    tab_b1, tab_b2, tab_b3 = st.tabs([
        "🔍 Conciliación y Arqueo Físico",
        "🍽️ Servir Alimento a Cerdos",
        "📦 Ingresar Compra a Bodega"
    ])

    with tab_b1:
        st.subheader("🔍 Conciliación Contable vs. Conteo Físico")
        if df_bodega.empty:
            st.info("Aún no hay productos registrados en bodega.")
        else:
            df_bodega["Clave_Prod"] = (
                df_bodega["Categoria_Insumo"].astype(str).str.strip() + " | " +
                df_bodega["Nombre_Insumo"].astype(str).str.strip() + " (" +
                df_bodega["Presentacion"].astype(str).str.strip() + " - " +
                df_bodega["Unidad_Medida"].astype(str).str.strip() + ")"
            )
            claves_unicas = opciones_limpias(df_bodega["Clave_Prod"])
            clave_sel = st.selectbox("Seleccionar Producto Exacto (Categoría | Nombre | Presentación | Unidad):", claves_unicas)

            sub_b = df_bodega[df_bodega["Clave_Prod"] == clave_sel]
            cat_p = str(sub_b.iloc[0]["Categoria_Insumo"]).strip()
            nom_p = str(sub_b.iloc[0]["Nombre_Insumo"]).strip()
            pres_p = str(sub_b.iloc[0]["Presentacion"]).strip()
            und_p = str(sub_b.iloc[0]["Unidad_Medida"]).strip()
            tam_p = float(to_num(sub_b["Contenido_Por_Unidad"]).iloc[-1])

            comprado_tot = float((to_num(sub_b["Unidades_Compradas"]) * to_num(sub_b["Contenido_Por_Unidad"])).sum())
            saldo_bodega_reg = float(to_num(sub_b["Saldo_Disponible"]).sum())

            consumo_int = 0.0
            if cat_p == "Concentrado" and not df_consumo.empty:
                consumo_int = float(to_num(df_consumo[df_consumo["Tipo_Concentrado"].astype(str).str.strip() == nom_p]["Kilos_Servidos"]).sum())
            elif cat_p != "Concentrado" and not df_sanidad.empty:
                consumo_int = float(to_num(df_sanidad[(df_sanidad["Medicamento"].astype(str).str.strip() == nom_p) & (df_sanidad["Unidad_Medida"].astype(str).str.strip() == und_p)]["Cantidad_Aplicada"]).sum())

            vendido_ext = 0.0
            if cat_p == "Concentrado" and not df_ventas.empty:
                v_c = df_ventas[(df_ventas["Linea_Negocio"].astype(str).str.contains("Concentrado|Cuido", na=False)) & (df_ventas["ID_Origen"].astype(str).str.strip() == nom_p)]
                vendido_ext = float(to_num(v_c["Peso_Pie_O_Kilos"]).sum())

            saldo_teorico_ecuacion = round(comprado_tot - consumo_int - vendido_ext, 2)
            bultos_cerrados = int(saldo_teorico_ecuacion // max(tam_p, 1.0))
            sueltos_teoricos = round(saldo_teorico_ecuacion % max(tam_p, 1.0), 2)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📐 Saldo Teórico (Ecuación)", f"{saldo_teorico_ecuacion:,.1f} {und_p}", delta=f"{bultos_cerrados} {pres_p}s + {sueltos_teoricos} {und_p}")
            m2.metric("💾 Saldo en Tabla Bodega", f"{saldo_bodega_reg:,.1f} {und_p}", delta=f"Descuadre libro: {saldo_bodega_reg - saldo_teorico_ecuacion:+.1f}")
            m3.metric("🐖 Consumo Interno Granja", f"{consumo_int:,.1f} {und_p}")
            m4.metric("🚚 Vendido a Terceros", f"{vendido_ext:,.1f} {und_p}")

            cf1, cf2 = st.columns(2)
            und_fisicas = cf1.number_input(f"¿Cuántos {pres_p}s cerrados contaste físicamente?", min_value=0, value=max(bultos_cerrados, 0), step=1)
            sueltos_fisicos = cf2.number_input(f"¿Cuántos {und_p} sueltos hay en báscula/frasco abierto?", min_value=0.0, value=max(float(sueltos_teoricos), 0.0), step=0.5)
            total_fisico = (und_fisicas * tam_p) + sueltos_fisicos
            dif_fis = round(total_fisico - saldo_teorico_ecuacion, 2)

            if abs(dif_fis) <= 0.5:
                st.success(f"✅ **INVENTARIO CONCILIADO:** El conteo físico ({total_fisico:.1f} {und_p}) cuadra con la ecuación de movimientos.")
            elif dif_fis < 0:
                st.error(f"⚠️ **FALTANTE FÍSICO:** Faltan **{abs(dif_fis):.1f} {und_p}** respecto al saldo teórico.")
            else:
                st.info(f"ℹ️ **SOBRANTE FÍSICO:** Hay **{dif_fis:.1f} {und_p}** más en físico que en movimientos.")

    with tab_b2:
        st.subheader("🍽️ Registrar Consumo Diario de Concentrado")
        conc_disp = df_bodega[(df_bodega["Categoria_Insumo"].astype(str).str.strip() == "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()
        if conc_disp.empty:
            st.warning("No hay concentrado con saldo disponible en bodega.")
        else:
            dest_tipo = st.selectbox("Destino del Alimento", ["Lote de Engorde", "Cerda Madre / Verraco"])
            lista_dest = opciones_limpias(lotes_activos["ID_Lote"]) if dest_tipo == "Lote de Engorde" else opciones_limpias(hembras_activas_df["ID_Animal"])
            if not lista_dest:
                st.warning("No hay animales/lotes activos para este destino.")
            else:
                id_dest = st.selectbox("Seleccionar Destino", lista_dest)
                tipos_conc = opciones_limpias(conc_disp["Nombre_Insumo"])
                fase_sug = tipos_conc[0]
                if dest_tipo == "Lote de Engorde" and not lotes_activos.empty:
                    fila_ld = lotes_activos[lotes_activos["ID_Lote"].astype(str).str.strip() == id_dest].iloc[0]
                    ev_ld = evaluar_fase_lote(fila_ld, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                    st.info(f"📌 Fase registrada: **{ev_ld['fase_registrada']}** | Consumido: **{ev_ld['kg_consumidos_fase']:.1f} / {ev_ld['kg_meta_lote']:.1f} kg**")
                    if ev_ld["fase_registrada"] in tipos_conc:
                        fase_sug = ev_ld["fase_registrada"]

                t_conc_sel = st.selectbox("Concentrado a Servir", tipos_conc, index=tipos_conc.index(fase_sug) if fase_sug in tipos_conc else 0)
                saldo_tot_c = float(to_num(conc_disp[conc_disp["Nombre_Insumo"].astype(str).str.strip() == t_conc_sel]["Saldo_Disponible"]).sum())
                st.caption(f"Disponible real en bodega: **{saldo_tot_c:.2f} kg**")

                min_kg_val = min(0.1, saldo_tot_c)
                kg_serv = st.number_input("Kilos Servidos (kg)", min_value=min_kg_val, max_value=saldo_tot_c, value=min(10.0, saldo_tot_c), step=0.5)
                costo_fifo_serv, plan_desc, faltante_fifo = calcular_fifo_bodega(df_bodega, "Concentrado", t_conc_sel, "Bulto", "kg", kg_serv)

                if st.button("💾 Descontar Alimento (FIFO) y Cargar Costo"):
                    if faltante_fifo > 0 or kg_serv > saldo_tot_c:
                        st.error("🚫 Saldo insuficiente en bodega para cubrir esa cantidad.")
                    else:
                        for idx, nuevo_s in plan_desc:
                            df_bodega.at[idx, "Saldo_Disponible"] = nuevo_s
                        nuevo_cons = pd.DataFrame([{
                            "Fecha": hoy.strftime("%Y-%m-%d"),
                            "Destino_Tipo": dest_tipo,
                            "ID_Destino": id_dest,
                            "Tipo_Concentrado": t_conc_sel,
                            "Kilos_Servidos": kg_serv,
                            "Costo_Total": round(costo_fifo_serv, 0)
                        }])
                        df_consumo = pd.concat([df_consumo, nuevo_cons], ignore_index=True)

                        if dest_tipo == "Lote de Engorde":
                            df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == id_dest, "Kg_Alimento_Consumido"] = to_num(df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == id_dest, "Kg_Alimento_Consumido"]) + kg_serv
                            df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == id_dest, "Costo_Alimento"] = to_num(df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == id_dest, "Costo_Alimento"]) + costo_fifo_serv
                            guardar_hoja("Lotes_Engorde", df_lotes)
                        else:
                            df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == id_dest, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == id_dest, "Costo_Acumulado_Ciclo"]) + costo_fifo_serv
                            guardar_hoja("Reproductores", df_rep)

                        guardar_hoja("Bodega", df_bodega)
                        guardar_hoja("Consumo_Alimento", df_consumo)
                        st.success(f"Descontados {kg_serv} kg de {t_conc_sel} (Costo FIFO: ${costo_fifo_serv:,.0f}).")
                        st.rerun()

    with tab_b3:
        st.subheader("📦 Registrar Compra de Concentrado o Medicamentos")
        cat_ins = st.selectbox("Categoría de Insumo", ["Concentrado", "Antibiótico / Inyectable", "Vacuna / Biológico", "Vitamina / Desparasitante"])
        if cat_ins == "Concentrado":
            nom_ins = st.selectbox("Tipo de Concentrado", LISTA_CONCENTRADOS)
            pres = "Bulto"
            und_med = "kg"
            cont_und = st.number_input("Kilos por Bulto", min_value=1.0, value=40.0)
        else:
            nom_ins = st.text_input("Nombre del Medicamento / Vacuna", placeholder="Ej: Florfenicol Inyectable 30%")
            pres = st.selectbox("Presentación", ["Frasco", "Caja / Dosis", "Sobre / Bolsa"])
            und_med = st.selectbox("Unidad de Descuento", ["ml", "dosis", "g"])
            cont_und = st.number_input(f"Contenido por {pres} ({und_med})", min_value=1.0, value=100.0)

        cant_comp = st.number_input(f"Cantidad de {pres}s comprados", min_value=1, value=1)
        precio_u = st.number_input(f"Precio de Compra por {pres} ($)", min_value=0.0, value=85000.0, step=1000.0)
        tot_medida = cant_comp * cont_und
        costo_med = precio_u / max(cont_und, 1.0)

        if st.button("💾 Ingresar Compra a Bodega"):
            if not str(nom_ins).strip():
                st.error("Escribe el nombre del insumo.")
            else:
                nuevo_id_b = generar_siguiente_id(df_bodega, "ID_Compra", "INS")
                nueva_c = pd.DataFrame([{
                    "ID_Compra": nuevo_id_b,
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Categoria_Insumo": cat_ins,
                    "Nombre_Insumo": str(nom_ins).strip(),
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
                st.success(f"Compra {nuevo_id_b} ingresada a bodega.")
                st.rerun()

    st.divider()
    st.subheader("📊 Resumen Ponderado de Bodega y Botiquín (Sin Mezclar Unidades)")
    if not df_bodega.empty:
        df_bodega["Saldo_Num"] = to_num(df_bodega["Saldo_Disponible"])
        df_bodega["Valor_Fila"] = df_bodega["Saldo_Num"] * to_num(df_bodega["Costo_Por_Medida"])
        res_bod = df_bodega.groupby(["Categoria_Insumo", "Nombre_Insumo", "Presentacion", "Unidad_Medida"], as_index=False).agg(
            Saldo_Total=("Saldo_Num", "sum"),
            Tam_Envase=("Contenido_Por_Unidad", "last"),
            Valor_Total_Stock=("Valor_Fila", "sum")
        )
        res_bod["Unidades / Bultos"] = (res_bod["Saldo_Total"] / res_bod["Tam_Envase"].replace(0, 1)).round(2)
        res_bod["Costo_Ponderado_Real ($)"] = np.where(res_bod["Saldo_Total"] > 0, (res_bod["Valor_Total_Stock"] / res_bod["Saldo_Total"]).round(2), 0.0)
        if es_admin:
            res_bod["Valor_Total_Stock"] = res_bod["Valor_Total_Stock"].round(0)
            cols_v = ["Categoria_Insumo", "Nombre_Insumo", "Presentacion", "Saldo_Total", "Unidad_Medida", "Unidades / Bultos", "Costo_Ponderado_Real ($)", "Valor_Total_Stock"]
        else:
            cols_v = ["Categoria_Insumo", "Nombre_Insumo", "Presentacion", "Saldo_Total", "Unidad_Medida", "Unidades / Bultos"]
        st.dataframe(res_bod[cols_v], use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 8: SANIDAD Y MORTALIDAD (INCLUYE BAJA REAL DE REPRODUCTORES)
# ==============================================================================
elif menu == "💉 Sanidad y Mortalidad":
    st.markdown('<div class="main-header">💉 Sanidad, Botiquín y Control de Mortalidad</div>', unsafe_allow_html=True)

    reproductores_vivos_ids = opciones_limpias(df_rep[~df_rep["Estado"].astype(str).str.strip().isin(ESTADOS_INACTIVOS_REP)]["ID_Animal"]) if not df_rep.empty else []
    destinos_posibles = opciones_limpias(lotes_activos["ID_Lote"]) + reproductores_vivos_ids
    meds_disp = df_bodega[(df_bodega["Categoria_Insumo"].astype(str).str.strip() != "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)].copy() if not df_bodega.empty else pd.DataFrame()

    if not destinos_posibles:
        st.info("No hay lotes ni reproductores activos registrados.")
    else:
        c1, c2 = st.columns(2)
        dest_san = c1.selectbox("Lote o Reproductor Activo", destinos_posibles)
        cat_san = c2.selectbox("Tipo de Evento", ["Tratamiento / Enfermedad", "Vacunación Preventiva", "Mortalidad (Baja)"])
        sintoma = c1.text_input("Detalle / Síntoma / Causa", placeholder="Ej: Tos respiratoria / Muerte súbita")

        usar_botiquin = False
        med_nombre, med_pres, unidad_ap = "Ninguno", "", ""
        cant_aplicada, costo_san = 0.0, 0.0
        plan_desc_med = []

        if cat_san != "Mortalidad (Baja)":
            if not meds_disp.empty:
                usar_botiquin = st.checkbox("💊 Descontar medicamento del Botiquín en Bodega", value=True)
            if usar_botiquin and not meds_disp.empty:
                meds_disp["Clave_Med"] = meds_disp["Nombre_Insumo"].astype(str).str.strip() + " | " + meds_disp["Presentacion"].astype(str).str.strip() + " (" + meds_disp["Unidad_Medida"].astype(str).str.strip() + ")"
                clave_med_sel = c2.selectbox("Medicamento Exacto (Nombre | Presentación | Unidad)", opciones_limpias(meds_disp["Clave_Med"]))
                sub_m = meds_disp[meds_disp["Clave_Med"] == clave_med_sel]
                med_nombre = str(sub_m.iloc[0]["Nombre_Insumo"]).strip()
                med_pres = str(sub_m.iloc[0]["Presentacion"]).strip()
                unidad_ap = str(sub_m.iloc[0]["Unidad_Medida"]).strip()
                cat_med = str(sub_m.iloc[0]["Categoria_Insumo"]).strip()
                saldo_m = float(to_num(sub_m["Saldo_Disponible"]).sum())

                cant_aplicada = c1.number_input(f"Cantidad Aplicada ({unidad_ap}) [Máx: {saldo_m:.2f}]", min_value=min(0.1, saldo_m), max_value=saldo_m, value=min(5.0, saldo_m), step=0.5)
                costo_san, plan_desc_med, _ = calcular_fifo_bodega(df_bodega, cat_med, med_nombre, med_pres, unidad_ap, cant_aplicada)
                if es_admin:
                    st.info(f"Costo exacto FIFO a cargar: **${costo_san:,.0f}**")
            else:
                med_nombre = c2.text_input("Medicamento Aplicado (Externo)", value="Ninguno")
                costo_san = c1.number_input("Costo del Tratamiento ($)", min_value=0.0, value=0.0, step=1000.0) if es_admin else 0.0
        else:
            if dest_san in reproductores_vivos_ids:
                st.error(f"⚠️ **MORTALIDAD DE REPRODUCTOR ({dest_san}):** Al guardar, el animal pasará permanentemente a estado **Muerta**, se cerrará cualquier preñez/ciclo activo y saldrá del inventario.")

        dias_ret = c2.number_input("Días de Retiro (Prohibición de venta)", min_value=0, max_value=60, value=0)
        f_fin_ret = (hoy + timedelta(days=int(dias_ret))).strftime("%Y-%m-%d") if dias_ret > 0 else ""

        if st.button("💾 Registrar Evento Sanitario"):
            if usar_botiquin and cant_aplicada > 0:
                for idx, nuevo_s in plan_desc_med:
                    df_bodega.at[idx, "Saldo_Disponible"] = nuevo_s
                guardar_hoja("Bodega", df_bodega)

            nuevo_san = pd.DataFrame([{
                "Fecha": hoy.strftime("%Y-%m-%d"),
                "ID_Destino": dest_san,
                "Categoria": cat_san,
                "Detalle_Sintoma": sintoma,
                "Medicamento": med_nombre,
                "Cantidad_Aplicada": cant_aplicada,
                "Unidad_Medida": unidad_ap if usar_botiquin else "",
                "Costo": round(costo_san, 0),
                "Dias_Retiro": dias_ret,
                "Fecha_Fin_Retiro": f_fin_ret,
                "Estado_Caso": "En Tratamiento" if dias_ret > 0 else "Cerrado"
            }])
            df_sanidad = pd.concat([df_sanidad, nuevo_san], ignore_index=True)

            if not lotes_activos.empty and dest_san in opciones_limpias(lotes_activos["ID_Lote"]):
                if cat_san == "Mortalidad (Baja)":
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == dest_san, "Cant_Actual"] = (to_num(df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == dest_san, "Cant_Actual"]) - 1).clip(lower=0)
                df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == dest_san, "Costo_Sanidad"] = to_num(df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == dest_san, "Costo_Sanidad"]) + costo_san
                guardar_hoja("Lotes_Engorde", df_lotes)
            elif dest_san in reproductores_vivos_ids:
                if cat_san == "Mortalidad (Baja)":
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Estado"] = "Muerta"
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Banda"] = "N/A"
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Observaciones"] = f"Fallecida el {hoy.strftime('%Y-%m-%d')}: {sintoma}"
                    # Cerrar eventos reproductivos abiertos de esa cerda
                    if not df_repro.empty:
                        df_repro.loc[(df_repro["ID_Cerda"].astype(str).str.strip() == dest_san) & (df_repro["Estado_Prenez"].astype(str).str.strip().isin(["Pendiente Eco", "Preñada", "En Lactancia"])), "Estado_Prenez"] = "Muerte Madre"
                        guardar_hoja("Reproduccion", df_repro)
                else:
                    df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Costo_Acumulado_Ciclo"] = to_num(df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == dest_san, "Costo_Acumulado_Ciclo"]) + costo_san
                guardar_hoja("Reproductores", df_rep)

            guardar_hoja("Sanidad_Eventos", df_sanidad)
            st.success("Evento sanitario registrado e inventarios actualizados.")
            st.rerun()

    st.divider()
    st.subheader("Historial Sanitario y Mortalidad")
    st.dataframe(df_sanidad, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 9: VENTAS (FIFO EXACTO EN CUIDO, BLOQUEO REPRODUCTIVO Y RETIRO GLOBAL)
# ==============================================================================
elif menu == "💰 Ventas (Marranos / Cuido)":
    st.markdown('<div class="main-header">💰 Registro de Ventas por Línea de Negocio</div>', unsafe_allow_html=True)

    linea_venta = st.selectbox("Selecciona la Línea de Ingreso:", [
        "🐖 Venta Cerdos Cebados (Pie / Canal)",
        "🌾 Venta Concentrado / Cuido (Valoración FIFO Exacta)",
        "🐷 Venta Cerda de Descarte (Solo Elegibles)",
        "🍼 Venta Lechones Destetos"
    ])

    def verificar_retiro_sanitario(id_objetivo):
        if df_sanidad.empty:
            return None
        sub_r = df_sanidad[df_sanidad["ID_Destino"].astype(str).str.strip() == id_objetivo]
        for _, rs in sub_r.iterrows():
            fr = parse_fecha(rs.get("Fecha_Fin_Retiro", ""))
            if fr and fr >= hoy:
                return f"Tiempo de retiro vigente hasta `{fr}` por `{rs['Medicamento']}`."
        return None

    if "Concentrado" in linea_venta:
        st.subheader("🌾 Venta Comercial de Concentrado (Costo FIFO Exacto)")
        conc_v = df_bodega[(df_bodega["Categoria_Insumo"].astype(str).str.strip() == "Concentrado") & (to_num(df_bodega["Saldo_Disponible"]) > 0)] if not df_bodega.empty else pd.DataFrame()
        if conc_v.empty:
            st.warning("No hay concentrado disponible en bodega.")
        else:
            c1, c2 = st.columns(2)
            conc_v = conc_v.copy()
            conc_v["_ClaveVenta"] = (conc_v["Nombre_Insumo"].astype(str).str.strip() + " | " +
                                    conc_v["Presentacion"].astype(str).str.strip() + " | " +
                                    conc_v["Contenido_Por_Unidad"].astype(str).str.strip() + " " +
                                    conc_v["Unidad_Medida"].astype(str).str.strip())
            clave_v = c1.selectbox("Concentrado / presentación a vender", opciones_limpias(conc_v["_ClaveVenta"]))
            sub_pv = conc_v[conc_v["_ClaveVenta"] == clave_v]
            prod_v = str(sub_pv.iloc[0]["Nombre_Insumo"]).strip()
            pres_v = str(sub_pv.iloc[0]["Presentacion"]).strip()
            unidad_v = str(sub_pv.iloc[0]["Unidad_Medida"]).strip()
            tam_b_v = float(to_num(pd.Series([sub_pv.iloc[0]["Contenido_Por_Unidad"]])).iloc[0])
            saldo_kg_v = float(to_num(sub_pv["Saldo_Disponible"]).sum())
            max_bultos_v = saldo_kg_v / max(tam_b_v, 1.0)

            bultos_v = c2.number_input("Cantidad de Bultos a Vender", min_value=min(0.1, max_bultos_v), max_value=max_bultos_v, value=min(1.0, max_bultos_v), step=0.25)
            kilos_v = round(bultos_v * tam_b_v, 2)

            # Cálculo FIFO exacto ANTES de confirmar (coincide 100% con el guardado)
            costo_fifo_venta, plan_desc_venta, _ = calcular_fifo_bodega(df_bodega, "Concentrado", prod_v, pres_v, unidad_v, kilos_v)
            costo_bulto_fifo = costo_fifo_venta / max(bultos_v, 0.0001)

            margen_pct = c1.number_input("Margen Comercial (%) [Sugerido: 18%]", min_value=0.0, max_value=100.0, value=float(MARGEN_CUIDO_DEFAULT), step=1.0)
            costo_flete_v = c2.number_input("Costo Flete / Transporte Asumido ($)", min_value=0.0, value=0.0, step=5000.0)
            cliente_txt = st.text_input("Cliente / Finca Destino", placeholder="Ej: Finca El Roble")

            precio_venta_bulto = costo_bulto_fifo * (1 + margen_pct / 100.0)
            ingreso_b_cuido = round(bultos_v * precio_venta_bulto, 0)
            utilidad_n_cuido = round(ingreso_b_cuido - costo_fifo_venta - costo_flete_v, 0)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Costo FIFO Bodega / Bulto", f"${costo_bulto_fifo:,.0f}")
            m2.metric("Precio Venta / Bulto", f"${precio_venta_bulto:,.0f}")
            m3.metric("Ingreso Bruto Venta", f"${ingreso_b_cuido:,.0f}")
            m4.metric("Utilidad Neta Exacta", f"${utilidad_n_cuido:,.0f}")

            if st.button("💾 Confirmar Venta de Concentrado"):
                for idx, nuevo_s in plan_desc_venta:
                    df_bodega.at[idx, "Saldo_Disponible"] = nuevo_s
                nueva_vc = pd.DataFrame([{
                    "Fecha": hoy.strftime("%Y-%m-%d"),
                    "Linea_Negocio": "Venta Concentrado / Cuido",
                    "ID_Origen": prod_v,
                    "Modalidad": f"{bultos_v} unidades de {pres_v} ({tam_b_v:g} {unidad_v} c/u, +{margen_pct:.0f}%) | {cliente_txt}",
                    "Cantidad_Animales_O_Unidades": bultos_v,
                    "Peso_Pie_O_Kilos": kilos_v,
                    "Peso_Carne_Canal_Kg": 0.0,
                    "Costo_Sacrificio_O_Flete": costo_flete_v,
                    "Ingreso_Bruto": ingreso_b_cuido,
                    "Costo_Produccion_Asignado": round(costo_fifo_venta, 0),
                    "Utilidad_Neta": utilidad_n_cuido
                }])
                df_ventas = pd.concat([df_ventas, nueva_vc], ignore_index=True)
                guardar_hoja("Bodega", df_bodega)
                guardar_hoja("Ventas", df_ventas)
                st.success(f"Venta registrada con utilidad exacta de ${utilidad_n_cuido:,.0f}.")
                st.rerun()

    elif "Descarte" in linea_venta:
        st.subheader("🐷 Venta de Cerda de Descarte")
        # Bloquea venta de cerdas Preñadas, Lactantes o con ciclo activo
        hembras_elegibles_desc = hembras_activas_df[
            (hembras_activas_df["Estado"].astype(str).str.strip() == "Vacía") &
            (~hembras_activas_df["ID_Animal"].astype(str).str.strip().isin(cerdas_con_ciclo_abierto))
        ] if not hembras_activas_df.empty else pd.DataFrame()

        lista_hd = opciones_limpias(hembras_elegibles_desc["ID_Animal"]) if not hembras_elegibles_desc.empty else []
        if not lista_hd:
            st.warning("⚠️ No hay cerdas en estado **Vacía** elegibles para descarte (las cerdas Preñadas, Lactantes o Inseminadas están protegidas).")
        else:
            c1, c2 = st.columns(2)
            cerda_d = c1.selectbox("Cerda Elegible a Descartar (Solo Vacías)", lista_hd)
            err_ret_d = verificar_retiro_sanitario(cerda_d)
            if err_ret_d:
                st.error(f"🚫 BLOQUEO SANITARIO: {err_ret_d}")

            peso_d = c2.number_input("Peso en Pie de la Cerda (kg)", min_value=80.0, max_value=350.0, value=190.0)
            precio_kg_d = c1.number_input("Precio Kilo Descarte ($)", min_value=0.0, value=6200.0, step=100.0)
            costo_acum_d = float(to_num(hembras_elegibles_desc[hembras_elegibles_desc["ID_Animal"].astype(str).str.strip() == cerda_d]["Costo_Acumulado_Ciclo"]).iloc[0])
            ing_d = peso_d * precio_kg_d
            util_d = ing_d - costo_acum_d
            c2.metric("Ingreso Neto Descarte", f"${util_d:,.0f}")

            if st.button("💾 Confirmar Venta de Cerda Descarte", disabled=bool(err_ret_d)):
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
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_d, "Estado"] = "Descartada / Vendida"
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_d, "Banda"] = "N/A"
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_d, "Fecha_Ultimo_Cambio_Estado"] = hoy.strftime("%Y-%m-%d")
                df_rep.loc[df_rep["ID_Animal"].astype(str).str.strip() == cerda_d, "Observaciones"] = f"Vendida por descarte el {hoy.strftime('%Y-%m-%d')}"
                guardar_hoja("Ventas", df_ventas)
                guardar_hoja("Reproductores", df_rep)
                st.success(f"Cerda {cerda_d} vendida y retirada del inventario activo.")
                st.rerun()

    else:
        es_desteto = "Lechones" in linea_venta
        st.subheader("🍼 Venta de Lechones Destetos" if es_desteto else "🐖 Venta de Cerdos Cebados (Pie / Canal)")

        if es_desteto:
            lotes_filtrados_v = lotes_activos[lotes_activos["Etapa"].astype(str).str.strip().isin(["Fase 0", "Fase 1"])] if not lotes_activos.empty else pd.DataFrame()
        else:
            lotes_filtrados_v = lotes_activos if not lotes_activos.empty else pd.DataFrame()

        lista_lv = opciones_limpias(lotes_filtrados_v["ID_Lote"]) if not lotes_filtrados_v.empty else []
        if not lista_lv:
            st.info("No hay lotes elegibles para esta modalidad de venta.")
        else:
            lote_v = st.selectbox("Seleccionar Lote", lista_lv)
            fila_v = lotes_filtrados_v[lotes_filtrados_v["ID_Lote"].astype(str).str.strip() == lote_v].iloc[0]
            err_ret_l = verificar_retiro_sanitario(lote_v)
            if err_ret_l:
                st.error(f"🚫 BLOQUEO SANITARIO: {err_ret_l}")

            cant_disp = int(to_num(pd.Series([fila_v["Cant_Actual"]])).iloc[0])
            c_tot_lote = float(to_num(pd.Series([fila_v["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([fila_v["Costo_Sanidad"]])).iloc[0])

            c1, c2 = st.columns(2)
            if es_desteto:
                modalidad = "Lechón Desteto por Unidad"
                cant_v = c1.number_input("Cantidad de Lechones Vendidos", min_value=1, max_value=max(cant_disp, 1), value=cant_disp)
                peso_pie_v = c2.number_input("Peso Total Vendido (kg)", min_value=1.0, value=float(cant_v * to_num(pd.Series([fila_v["Peso_Prom_Actual_Kg"]])).iloc[0]))
                precio_und = c1.number_input("Precio Venta por Lechón ($)", min_value=0.0, value=240000.0, step=5000.0)
                peso_canal_v, costo_sac = 0.0, 0.0
                ingreso_bruto = cant_v * precio_und
            else:
                modalidad = c1.selectbox("Modalidad de Venta", ["En Pie (Cerdo entero en báscula)", "Beneficiado (Carne en canal por kilos)"])
                cant_v = c2.number_input("Cantidad de Cerdos Vendidos", min_value=1, max_value=max(cant_disp, 1), value=cant_disp)
                peso_pie_v = c1.number_input("Peso Total en Pie Vendido (kg)", min_value=1.0, value=float(cant_v * to_num(pd.Series([fila_v["Peso_Prom_Actual_Kg"]])).iloc[0]))

                if "En Pie" in modalidad:
                    precio_kg = c2.number_input("Precio Venta por Kilo en Pie ($)", min_value=0.0, value=8500.0, step=100.0)
                    peso_canal_v, costo_sac = 0.0, 0.0
                    ingreso_bruto = peso_pie_v * precio_kg
                else:
                    rend_pct = c2.number_input("Rendimiento en Canal (%)", min_value=50.0, max_value=90.0, value=78.0)
                    peso_canal_v = round(peso_pie_v * (rend_pct / 100.0), 2)
                    precio_canal = c1.number_input("Precio Venta Kilo en Canal ($)", min_value=0.0, value=14500.0, step=100.0)
                    costo_sac = c2.number_input("Costo Sacrificio / Desposte / Transporte ($)", min_value=0.0, value=float(cant_v * 65000), step=5000.0)
                    ingreso_bruto = peso_canal_v * precio_canal

            costo_prop = (c_tot_lote / max(cant_disp, 1)) * cant_v
            utilidad_neta = ingreso_bruto - costo_prop - costo_sac
            util_por_cerdo_venta = utilidad_neta / max(cant_v, 1)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ingreso Bruto", f"${ingreso_bruto:,.0f}")
            m2.metric("Costo Prod. + Sacrificio", f"${(costo_prop + costo_sac):,.0f}")
            m3.metric("Utilidad Neta Total", f"${utilidad_neta:,.0f}")
            m4.metric("Utilidad / Cerdo", f"${util_por_cerdo_venta:,.0f}", delta=f"{util_por_cerdo_venta - META_UTIL_CERDO:+,.0f} vs meta")

            if st.button("💾 Confirmar Venta", disabled=bool(err_ret_l)):
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
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Cant_Actual"] = 0
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Estado"] = "Vendido"
                else:
                    ratio_rem = nuevo_saldo_cerdos / cant_disp
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Cant_Actual"] = nuevo_saldo_cerdos
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Costo_Inicial_Lechones"] = round(float(to_num(pd.Series([fila_v["Costo_Inicial_Lechones"]])).iloc[0]) * ratio_rem, 0)
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Costo_Alimento"] = round(float(to_num(pd.Series([fila_v["Costo_Alimento"]])).iloc[0]) * ratio_rem, 0)
                    df_lotes.loc[df_lotes["ID_Lote"].astype(str).str.strip() == lote_v, "Costo_Sanidad"] = round(float(to_num(pd.Series([fila_v["Costo_Sanidad"]])).iloc[0]) * ratio_rem, 0)

                guardar_hoja("Ventas", df_ventas)
                guardar_hoja("Lotes_Engorde", df_lotes)
                st.success("Venta registrada y costos prorrateados correctamente.")
                st.rerun()

    st.divider()
    st.subheader("Historial General de Ventas")
    st.dataframe(df_ventas, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 10: ANALÍTICA E INFORMES INTEGRALES (KPIs REPRODUCTIVOS Y SANITARIOS)
# ==============================================================================
elif menu == "📊 Analítica, Metas y Reportes":
    st.markdown('<div class="main-header">📊 Inteligencia Productiva e Informes Ejecutivos</div>', unsafe_allow_html=True)

    tab_a1, tab_a2, tab_a3 = st.tabs(["📈 Curva Dinámica de Peso", "🧬 KPIs Reproductivos y Sanitarios", "📄 Reporte Ejecutivo Integral (Word / PDF)"])

    with tab_a1:
        dias_eje = [0, EDAD_DESTETE] + [f["edad_fin"] for f in PROGRAMA_FASES]
        peso_eje = [peso_esperado_por_edad(d, PESO_NAC, PESO_DEST_21D, EDAD_DESTETE, PROGRAMA_FASES) for d in dias_eje]

        fig_curva = go.Figure()
        fig_curva.add_trace(go.Scatter(x=dias_eje, y=peso_eje, mode="lines+markers", name=f"Curva Objetivo ({PESO_META:.0f} kg)", line=dict(color="#1E8449", width=3, dash="dash")))

        if not lotes_activos.empty:
            for _, l in lotes_activos.iterrows():
                ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                p_a = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                fig_curva.add_trace(go.Scatter(
                    x=[ev_l["edad_actual"]], y=[p_a], mode="markers+text",
                    name=l["ID_Lote"], text=[f"{l['ID_Lote']} ({p_a} kg)"],
                    textposition="top center", marker=dict(size=14)
                ))

        fig_curva.update_layout(xaxis_title="Edad del Cerdo (Días)", yaxis_title="Peso Promedio (kg)", height=430)
        st.plotly_chart(fig_curva, use_container_width=True)

    # Cálculo de KPIs reproductivos y mortalidad globales para Analítica y Reportes
    partos_realizados = df_repro[to_num(df_repro["Nacidos_Vivos"]) > 0] if not df_repro.empty else pd.DataFrame()
    total_partos = len(partos_realizados)
    prom_nacidos_vivos = float(to_num(partos_realizados["Nacidos_Vivos"]).mean()) if total_partos > 0 else 0.0
    destetados_df = df_repro[to_num(df_repro["Destetados_Reales"]) > 0] if not df_repro.empty else pd.DataFrame()
    prom_destetados = float(to_num(destetados_df["Destetados_Reales"]).mean()) if not destetados_df.empty else 0.0
    tot_servicios = len(df_repro) if not df_repro.empty else 0
    tot_repeticiones = int(to_num(df_rep["Repeticiones"]).sum()) if not df_rep.empty else 0
    tot_abortos = int(to_num(df_rep["Abortos"]).sum()) if not df_rep.empty else 0

    # Indicador por cohorte de servicio: se excluyen eventos aún pendientes de diagnóstico.
    estados_diagnosticados = ["Preñada", "Repitió Celo", "Aborto", "En Lactancia", "Destetada", "Muerte Madre"]
    servicios_diagnosticados = df_repro[df_repro["Estado_Prenez"].astype(str).str.strip().isin(estados_diagnosticados)] if not df_repro.empty else pd.DataFrame()
    estados_confirmacion = ["Preñada", "Aborto", "En Lactancia", "Destetada", "Muerte Madre"]
    servicios_confirmados = servicios_diagnosticados[servicios_diagnosticados["Estado_Prenez"].astype(str).str.strip().isin(estados_confirmacion)] if not servicios_diagnosticados.empty else pd.DataFrame()
    tasa_prenez_pct = (len(servicios_confirmados) / len(servicios_diagnosticados) * 100.0) if not servicios_diagnosticados.empty else 0.0

    # La base no permite calcular una tasa de mortalidad por cohorte fiable (no separa
    # mortalidades de ventas ni vincula bajas a edad/parto); mostramos conteo verificable.
    bajas_lotes = int(df_sanidad["Categoria"].astype(str).str.contains("Mortalidad", na=False).sum()) if not df_sanidad.empty else 0

    with tab_a2:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("🐷 Partos Registrados", total_partos, delta=f"Prom. Nac. Vivos: {prom_nacidos_vivos:.1f}")
        k2.metric("🍼 Prom. Destetados", f"{prom_destetados:.1f} lechones")
        k3.metric("🎯 Preñez confirmada / servicios diagnosticados", f"{tasa_prenez_pct:.1f}%", delta=f"Diagn.: {len(servicios_diagnosticados)} | Confirm.: {len(servicios_confirmados)}")
        k4.metric("📉 Mortalidades registradas", f"{bajas_lotes}")

    with tab_a3:
        capex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Inversión Infraestructura / Activos"]["Monto"]).sum() if not df_inv.empty else 0.0
        opex_tot = to_num(df_inv[df_inv["Tipo_Movimiento"] == "Gasto Fijo Operativo"]["Monto"]).sum() if not df_inv.empty else 0.0
        ing_tot = to_num(df_ventas["Ingreso_Bruto"]).sum() if not df_ventas.empty else 0.0
        util_lotes = to_num(df_ventas["Utilidad_Neta"]).sum() if not df_ventas.empty else 0.0
        util_granja = util_lotes - opex_tot

        def generar_word():
            doc = Document()
            titulo = doc.add_heading("INFORME EJECUTIVO INTEGRAL - PORCIGESTIÓN PRO v3", 0)
            titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_fecha = doc.add_paragraph(
                f"Fecha: {hoy.strftime('%d/%m/%Y')} | Escalón 1: {META_INTERMEDIA_MES} cerdos/mes | Meta Final: {META_CERDOS_MES} cerdos/mes"
            )
            p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_heading("1. Indicadores Reproductivos y Sanitarios", level=1)
            doc.add_paragraph(
                f"• Madres activas en inventario: {madres_activas_tot}\n"
                f"• Total partos registrados: {total_partos} | Promedio nacidos vivos: {prom_nacidos_vivos:.1f} | Promedio destetados: {prom_destetados:.1f}\n"
                f"• Preñez confirmada / servicios diagnosticados: {tasa_prenez_pct:.1f}% ({len(servicios_confirmados)}/{len(servicios_diagnosticados)})\n"
                f"• Mortalidades registradas (eventos): {bajas_lotes}"
            )

            doc.add_heading("2. Resumen Financiero y Metas Comerciales", level=1)
            doc.add_paragraph(
                f"• Inversión infraestructura y pie de cría (CAPEX): ${capex_tot:,.0f}\n"
                f"• Gastos fijos operativos (OPEX): ${opex_tot:,.0f}\n"
                f"• Ingresos brutos totales: ${ing_tot:,.0f}\n"
                f"• Resultado estimado de ventas menos gastos fijos registrados: ${util_granja:,.0f}\n"
                f"• Margen promedio por cerdo vendido: ${utilidad_prom_por_cerdo_hist:,.0f} (Meta: ${META_UTIL_CERDO:,.0f})\n"
                f"• Ventas últimos 30 días: {cerdos_vendidos_30d} cerdos"
            )

            doc.add_heading("3. Lotes Activos y Fases Nutricionales", level=1)
            if not lotes_activos.empty:
                t_lot = doc.add_table(rows=1, cols=6)
                t_lot.style = "Table Grid"
                hl = t_lot.rows[0].cells
                hl[0].text, hl[1].text, hl[2].text, hl[3].text, hl[4].text, hl[5].text = "Lote", "Fase", "Cerdos", "Peso Real", "Alimento Fase", "Costo Total"
                for _, l in lotes_activos.iterrows():
                    ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    ctot = float(to_num(pd.Series([l["Costo_Inicial_Lechones"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Alimento"]])).iloc[0]) + float(to_num(pd.Series([l["Costo_Sanidad"]])).iloc[0])
                    rc = t_lot.add_row().cells
                    rc[0].text = str(l["ID_Lote"])
                    rc[1].text = str(ev_l["fase_registrada"])
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
            pdf.set_font("Helvetica", "B", 13)
            pdf.multi_cell(0, 8, "INFORME EJECUTIVO INTEGRAL - PORCIGESTION PRO v3", align="C")
            pdf.set_font("Helvetica", "", 9)
            pdf.multi_cell(0, 6, f"Fecha: {hoy.strftime('%d/%m/%Y')} | Escalon 1: {META_INTERMEDIA_MES}/mes | Meta Final: {META_CERDOS_MES}/mes", align="C")
            pdf.ln(4)

            secciones = [
                ("1. INDICADORES REPRODUCTIVOS Y SANITARIOS", [
                    f"- Madres activas: {madres_activas_tot} | Partos: {total_partos} | Prom. Nacidos Vivos: {prom_nacidos_vivos:.1f} | Destetados: {prom_destetados:.1f}",
                    f"- Preñez confirmada/diagnosticados: {tasa_prenez_pct:.1f}% ({len(servicios_confirmados)}/{len(servicios_diagnosticados)}) | Mortalidades registradas: {bajas_lotes}"
                ]),
                ("2. RESUMEN FINANCIERO INTEGRAL", [
                    f"- CAPEX: ${capex_tot:,.0f} | OPEX: ${opex_tot:,.0f} | Ingresos Totales: ${ing_tot:,.0f}",
                    f"- Resultado estimado (registros disponibles): ${util_granja:,.0f} | Margen/Cerdo: ${utilidad_prom_por_cerdo_hist:,.0f} | Ventas 30d: {cerdos_vendidos_30d} cerdos"
                ])
            ]
            for titulo_s, lineas in secciones:
                pdf.set_font("Helvetica", "B", 10)
                pdf.multi_cell(0, 7, titulo_s)
                pdf.set_font("Helvetica", "", 9)
                for ln in lineas:
                    pdf.multi_cell(0, 6, ln.encode("latin-1", "replace").decode("latin-1"))
                pdf.ln(3)

            pdf.set_font("Helvetica", "B", 10)
            pdf.multi_cell(0, 7, "3. LOTES ACTIVOS")
            pdf.set_font("Helvetica", "", 9)
            if not lotes_activos.empty:
                for _, l in lotes_activos.iterrows():
                    ev_l = evaluar_fase_lote(l, df_consumo, hoy, EDAD_DESTETE, PESO_DEST_21D, PROGRAMA_FASES)
                    cant = int(to_num(pd.Series([l["Cant_Actual"]])).iloc[0])
                    p_act = float(to_num(pd.Series([l["Peso_Prom_Actual_Kg"]])).iloc[0])
                    txt = f"{l['ID_Lote']} | {ev_l['fase_registrada']} ({ev_l['kg_consumidos_fase']:.1f}/{ev_l['kg_meta_lote']:.1f} kg) | {cant} cerdos | {p_act:.1f} kg"
                    pdf.multi_cell(0, 6, txt.encode("latin-1", "replace").decode("latin-1"))

            return bytes(pdf.output())

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="📘 Descargar Informe Integral en Word (.docx)",
                data=generar_word(),
                file_name=f"Reporte_PorciGestion_v2_{hoy.strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_d2:
            st.download_button(
                label="📕 Descargar Informe Integral en PDF (.pdf)",
                data=generar_pdf(),
                file_name=f"Reporte_PorciGestion_v2_{hoy.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# ==============================================================================
# MÓDULO 11: GESTIONAR CERDAS Y METAS (PROTECCIÓN CONTRA BORRADO CON HISTORIAL)
# ==============================================================================
elif menu == "⚙️ Gestionar Cerdas y Metas":
    st.markdown('<div class="main-header">⚙️ Administración de Reproductores y Parámetros</div>', unsafe_allow_html=True)

    t_cfg1, t_cfg2 = st.tabs(["🎯 Configurar Metas y Parámetros", "🐷 Agregar / Gestionar Reproductores"])

    with t_cfg1:
        c1, c2, c3 = st.columns(3)
        n_meta_interm = c1.number_input("Escalón 1 / Meta Intermedia (cerdos/mes)", min_value=2, max_value=200, value=int(META_INTERMEDIA_MES), step=2)
        n_meta_cerdos_mes = c2.number_input("Meta Final Cerdos Vendidos / Mes", min_value=4, max_value=500, value=int(META_CERDOS_MES), step=2)
        n_meta_util_cerdo = c3.number_input("Meta Utilidad Neta / Cerdo ($)", min_value=10000.0, max_value=2000000.0, value=float(META_UTIL_CERDO), step=10000.0)

        c4, c5, c6 = st.columns(3)
        n_margen_cuido = c4.number_input("Margen Venta Cuido (%)", min_value=0.0, max_value=100.0, value=float(MARGEN_CUIDO_DEFAULT), step=1.0)
        n_peso_meta = c5.number_input("Peso Objetivo Finalización (kg)", min_value=70.0, max_value=150.0, value=float(PESO_META), step=1.0)
        n_edad_dest = c6.number_input("Edad Destete Lechones (días)", min_value=18, max_value=35, value=int(EDAD_DESTETE), step=1)

        c7, c8 = st.columns(2)
        n_int_bandas = c7.selectbox("Ritmo entre Bandas (días)", [14, 15, 21], index=[14, 15, 21].index(INTERVALO_BANDAS) if INTERVALO_BANDAS in [14, 15, 21] else 1)
        n_cupo_banda = c8.number_input("Cupo Máx. Cerdas por Banda", min_value=1, max_value=10, value=int(CUPO_BANDA))

        if st.button("💾 Guardar Configuración"):
            edad_fin_calc = construir_programa_fases(n_edad_dest, PESO_DEST_21D, n_peso_meta)[-1]["edad_fin"]
            nueva_cfg = pd.DataFrame([
                {"Parametro": "Peso_Meta_Kg", "Valor": str(n_peso_meta), "Descripcion": "Peso objetivo del cerdo finalizado en pie (kg)"},
                {"Parametro": "Edad_Meta_Dias", "Valor": str(edad_fin_calc), "Descripcion": "Edad objetivo calculada desde destete + fases"},
                {"Parametro": "Edad_Destete_Dias", "Valor": str(n_edad_dest), "Descripcion": "Edad estándar de destete (días)"},
                {"Parametro": "Peso_Nacimiento_Prom_Kg", "Valor": str(PESO_NAC), "Descripcion": "Peso promedio al nacer (kg)"},
                {"Parametro": "Peso_Destete_21d_Kg", "Valor": str(PESO_DEST_21D), "Descripcion": "Peso objetivo al destete (kg)"},
                {"Parametro": "Margen_Venta_Cuido_Pct", "Valor": str(n_margen_cuido), "Descripcion": "Margen comercial por defecto venta concentrado (%)"},
                {"Parametro": "Meta_Cerdos_Mes", "Valor": str(n_meta_cerdos_mes), "Descripcion": "Meta final mensual de cerdos cebados vendidos"},
                {"Parametro": "Meta_Intermedia_Cerdos_Mes", "Valor": str(n_meta_interm), "Descripcion": "Meta intermedia mensual de cerdos vendidos"},
                {"Parametro": "Meta_Utilidad_Por_Cerdo", "Valor": str(n_meta_util_cerdo), "Descripcion": "Utilidad neta objetivo por cerdo ($)"},
                {"Parametro": "Intervalo_Bandas_Dias", "Valor": str(n_int_bandas), "Descripcion": "Días entre bandas"},
                {"Parametro": "Numero_Bandas", "Valor": str(NUM_BANDAS), "Descripcion": "Número de bandas en ciclo completo"},
                {"Parametro": "Cupo_Max_Cerdas_Banda", "Valor": str(n_cupo_banda), "Descripcion": "Máximo de cerdas por banda"},
                {"Parametro": "Alerta_Dias_Abiertos", "Valor": str(MAX_DIAS_ABIERTOS), "Descripcion": "Días máximos vacía antes de alerta"}
            ])
            guardar_hoja("Configuracion", nueva_cfg)
            st.success("Configuración guardada y vinculada dinámicamente a la curva de fases.")
            st.rerun()

    with t_cfg2:
        col_add, col_del = st.columns(2)
        with col_add:
            st.subheader("➕ Agregar Nuevo Reproductor")
            id_nuevo = st.text_input("ID Animal", placeholder="Ej: M-15")
            tipo_n = st.selectbox("Tipo", ["Hembra Madre", "Verraco Recelador (Estímulo Celo)"])
            raza_n = st.selectbox("Línea Genética / Raza", RAZAS_MARRANAS if "Hembra" in tipo_n else RAZAS_SEMEN)
            lista_b = [f"Banda {i:02d}" for i in range(1, NUM_BANDAS + 1)]
            banda_n = st.selectbox("Banda Asignada", lista_b + ["Sin Asignar"])
            obs_n = st.text_input("Observaciones", value="Activa")

            if st.button("➕ Agregar Animal"):
                id_clean = id_nuevo.strip()
                if not id_clean or id_clean in opciones_limpias(df_rep["ID_Animal"]):
                    st.error("ID vacío o ya existente.")
                else:
                    ocupadas_b = len(hembras_activas_df[hembras_activas_df["Banda"].astype(str).str.strip() == banda_n]) if not hembras_activas_df.empty else 0
                    if "Hembra" in tipo_n and banda_n != "Sin Asignar" and ocupadas_b >= CUPO_BANDA:
                        st.error(f"🚫 La {banda_n} ya alcanzó el cupo máximo de {CUPO_BANDA} cerdas. Selecciona otra banda o 'Sin Asignar'.")
                    else:
                        nueva_r = pd.DataFrame([{
                            "ID_Animal": id_clean,
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
                        st.success(f"{id_clean} agregado correctamente.")
                        st.rerun()

        with col_del:
            st.subheader("🗑️ Eliminar Reproductor (Solo Sin Historial)")
            lista_del = opciones_limpias(df_rep["ID_Animal"]) if not df_rep.empty else []
            if lista_del:
                id_del = st.selectbox("Seleccionar Animal", lista_del)
                tiene_hist = (
                    (not df_repro.empty and id_del in opciones_limpias(df_repro["ID_Cerda"])) or
                    (not df_lotes.empty and id_del in opciones_limpias(df_lotes["Madre_Origen"]))
                )
                if tiene_hist:
                    st.warning(f"🔒 **{id_del} tiene historial reproductivo/partos.** No se permite borrarlo definitivamente para proteger la trazabilidad. Registra su salida como **Mortalidad** o **Venta de Descarte**.")
                if st.button("🗑️ Eliminar Registro", disabled=tiene_hist):
                    df_rep = df_rep[df_rep["ID_Animal"].astype(str).str.strip() != id_del]
                    guardar_hoja("Reproductores", df_rep)
                    st.success(f"{id_del} eliminado.")
                    st.rerun()

        st.divider()
        st.subheader("Tabla de Reproductores")
        st.dataframe(df_rep, use_container_width=True, hide_index=True)