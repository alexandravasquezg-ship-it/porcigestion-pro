import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime, timedelta
from openpyxl.worksheet.datavalidation import DataValidation

# --- CONFIGURACIÓN DE PÁGINA Y ESTÉTICA ---
st.set_page_config(
    page_title="PorciGestión Pro | Inteligencia Porcina",
    page_icon="🐷",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    div[data-testid="metric-container"] {
        background-color: #FFFFFF;
        border-left: 5px solid #1E8449;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
    }
    h1, h2, h3 { color: #1B2631; font-family: 'Segoe UI', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #1B2631; color: #FFFFFF; }
    section[data-testid="stSidebar"] * { color: #ECF0F1 !important; }
    .stButton>button {
        background-color: #1E8449;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.2rem;
        font-weight: 600;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #145A32;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    .paso-card {
        background-color: #FFFFFF;
        padding: 16px 20px;
        border-radius: 10px;
        border-left: 4px solid #2E86C1;
        margin-bottom: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.04);
    }
</style>
""", unsafe_allow_html=True)

# --- SISTEMA DE SEGURIDAD Y ROLES ---
def verificar_acceso():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        st.session_state.rol = None

    if not st.session_state.autenticado:
        col1, col2, col3 = st.columns([1, 1.5, 1])
        with col2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown("""
            <div style="background-color: white; padding: 30px; border-radius: 15px; box-shadow: 0 8px 20px rgba(0,0,0,0.1); text-align: center;">
                <h1 style="color: #1E8449;">🐷 PorciGestión Pro</h1>
                <p style="color: #566573;">Sistema Inteligente de Administración y Analítica Porcina</p>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            usuario = st.selectbox("Selecciona tu perfil de ingreso:", ["Administrador (Dueño/a)", "Operario de Granja (Trabajador)"])
            clave = st.text_input("Contraseña de acceso:", type="password", placeholder="Ingresa la clave...")

            try:
                clave_admin = st.secrets.get("CLAVE_ADMIN", "admin2026")
                clave_oper = st.secrets.get("CLAVE_OPERARIO", "granja2026")
            except Exception:
                clave_admin = "admin2026"
                clave_oper = "granja2026"

            if st.button("Ingresar al Sistema", width="stretch"):
                if usuario == "Administrador (Dueño/a)" and clave == clave_admin:
                    st.session_state.autenticado = True
                    st.session_state.rol = "admin"
                    st.rerun()
                elif usuario == "Operario de Granja (Trabajador)" and clave == clave_oper:
                    st.session_state.autenticado = True
                    st.session_state.rol = "operario"
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta. Intenta nuevamente.")
        st.stop()

verificar_acceso()

# --- CONEXIÓN HÍBRIDA: GOOGLE SHEETS (NUBE) O EXCEL LOCAL ---
ARCHIVO_EXCEL = "granja_porcina.xlsx"
USAR_GSHEETS = False

try:
    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        from streamlit_gsheets import GSheetsConnection
        conn = st.connection("gsheets", type=GSheetsConnection)
        USAR_GSHEETS = True
except Exception:
    USAR_GSHEETS = False

OPCIONES_ESTADO_REPRO = ["Vacía", "Inseminada", "Preñada", "Lactante", "Recelador Activo", "Descartada"]
OPCIONES_TIPO_REPRO = ["Hembra Madre", "Verraco Recelador (Estímulo Celo)"]
OPCIONES_RAZAS = ["Landrace", "Yorkshire", "Large White", "Duroc", "Pietrain", "F1 (Landrace x Yorkshire)", "Hampshire", "Otra"]
OPCIONES_ETAPA_LOTE = ["Lactante", "Precebo", "Levante", "Engorde"]
OPCIONES_CONCENTRADO = ["Preiniciador", "Iniciador", "Levante", "Engorde", "Gestación", "Lactancia"]

def crear_estructura_base():
    cerdas_init = []
    for i in range(1, 15):
        cerdas_init.append({
            "ID_Animal": f"M-{i:02d}",
            "Tipo": "Hembra Madre",
            "Linea_Genetica": "F1 (Landrace x Yorkshire)",
            "Padre_ID": "Inseminación",
            "Madre_ID": "Externa",
            "Partos": 0,
            "Abortos": 0,
            "Repeticiones": 0,
            "Estado": "Vacía",
            "Racion_Diaria_Kg": 2.2,
            "Costo_Acumulado_Ciclo": 0.0,
            "Observaciones": "Inventario inicial"
        })
    cerdas_init.append({
        "ID_Animal": "V-01",
        "Tipo": "Verraco Recelador (Estímulo Celo)",
        "Linea_Genetica": "Duroc",
        "Padre_ID": "Externo",
        "Madre_ID": "M-01",
        "Partos": 0,
        "Abortos": 0,
        "Repeticiones": 0,
        "Estado": "Recelador Activo",
        "Racion_Diaria_Kg": 2.0,
        "Costo_Acumulado_Ciclo": 0.0,
        "Observaciones": "Solo para detección y estímulo de celo (No monta)"
    })
    return {
        "Reproductores": pd.DataFrame(cerdas_init),
        "Reproduccion": pd.DataFrame(columns=[
            "ID_Evento", "ID_Cerda", "Raza_Cerda", "Fecha_Servicio", "Tipo_Servicio",
            "Macho_O_Semen", "Uso_Recelador", "Costo_Servicio", "Fecha_Eco_28d",
            "Estado_Prenez", "Fecha_Prob_Parto", "Fecha_Real_Parto", "Nacidos_Vivos",
            "Mortinatos", "Momias", "Nombre_Lote_Asignado", "Fecha_Prob_Destete",
            "Destetados_Reales", "Peso_Destete_Total_Kg", "Costo_Por_Lechon"
        ]),
        "Lotes_Engorde": pd.DataFrame(columns=[
            "ID_Lote", "Madre_Origen", "Raza_Cruce", "Corral", "Etapa", "Fecha_Inicio",
            "Cant_Inicial", "Cant_Actual", "Peso_Prom_Inicial_Kg", "Peso_Prom_Actual_Kg",
            "Kg_Alimento_Consumido", "Costo_Inicial_Lechones", "Costo_Alimento",
            "Costo_Sanidad", "Estado"
        ]),
        "Bodega": pd.DataFrame([
            {"ID_Compra": "C-001", "Fecha": datetime.now().strftime("%Y-%m-%d"), "Tipo_Concentrado": "Gestación",
             "Bultos": 10, "Peso_Bulto_Kg": 40, "Precio_Bulto": 82000, "Costo_Kg": 2050.0, "Saldo_Kg": 400.0},
            {"ID_Compra": "C-002", "Fecha": datetime.now().strftime("%Y-%m-%d"), "Tipo_Concentrado": "Engorde",
             "Bultos": 15, "Peso_Bulto_Kg": 40, "Precio_Bulto": 86000, "Costo_Kg": 2150.0, "Saldo_Kg": 600.0}
        ]),
        "Consumo_Alimento": pd.DataFrame(columns=[
            "Fecha", "Destino_Tipo", "ID_Destino", "Tipo_Concentrado", "Kilos_Servidos", "Costo_Total"
        ]),
        "Sanidad_Eventos": pd.DataFrame(columns=[
            "Fecha", "ID_Destino", "Categoria", "Detalle_Sintoma", "Medicamento",
            "Costo", "Dias_Retiro", "Fecha_Fin_Retiro", "Estado_Caso"
        ]),
        "Ventas": pd.DataFrame(columns=[
            "Fecha", "ID_Origen", "Modalidad", "Cantidad_Animales", "Peso_Pie_Kg",
            "Peso_Carne_Canal_Kg", "Costo_Sacrificio", "Ingreso_Bruto", "Costo_Produccion_Asignado", "Utilidad_Neta"
        ])
    }

def asegurar_columnas(df, df_modelo):
    """Garantiza que las hojas existentes en Google Sheets incorporen columnas nuevas sin perder datos."""
    for col in df_modelo.columns:
        if col not in df.columns:
            df[col] = "" if df_modelo[col].dtype == object else 0
    return df

def aplicar_validaciones_excel(writer):
    wb = writer.book
    ws_rep = wb["Reproductores"]
    dv_tipo = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_TIPO_REPRO)}"', allow_blank=True)
    dv_estado = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_ESTADO_REPRO)}"', allow_blank=True)
    ws_rep.add_data_validation(dv_tipo)
    ws_rep.add_data_validation(dv_estado)
    dv_tipo.add("B2:B100")
    dv_estado.add("I2:I100")

def cargar_datos():
    base = crear_estructura_base()
    if USAR_GSHEETS:
        if "db_nube" not in st.session_state:
            datos = {}
            for hoja, df_def in base.items():
                try:
                    df_leido = conn.read(worksheet=hoja, ttl=300)
                    # Si la hoja ni siquiera tiene encabezados, se inicializa; si tiene encabezados (aunque tenga 0 filas), se respeta
                    if df_leido is None or len(df_leido.columns) == 0:
                        conn.update(worksheet=hoja, data=df_def)
                        datos[hoja] = df_def.copy()
                    else:
                        df_limpio = df_leido.dropna(how="all")
                        datos[hoja] = asegurar_columnas(df_limpio, df_def)
                except Exception:
                    datos[hoja] = df_def.copy()
            st.session_state["db_nube"] = datos
        return {k: v.copy() for k, v in st.session_state["db_nube"].items()}
    else:
        if not os.path.exists(ARCHIVO_EXCEL):
            with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
                for hoja, df in base.items():
                    df.to_excel(writer, sheet_name=hoja, index=False)
                aplicar_validaciones_excel(writer)
        xls = pd.ExcelFile(ARCHIVO_EXCEL)
        res = {}
        for sheet in base.keys():
            if sheet in xls.sheet_names:
                res[sheet] = asegurar_columnas(pd.read_excel(xls, sheet), base[sheet])
            else:
                res[sheet] = base[sheet]
        return res

def guardar_datos(dfs, hoja_modificada=None):
    if USAR_GSHEETS:
        anterior = st.session_state.get("db_nube", {})
        if hoja_modificada and hoja_modificada in dfs:
            conn.update(worksheet=hoja_modificada, data=dfs[hoja_modificada])
        else:
            for hoja, df in dfs.items():
                df_ant = anterior.get(hoja)
                if df_ant is None or not df.equals(df_ant):
                    conn.update(worksheet=hoja, data=df)
        st.session_state["db_nube"] = {k: v.copy() for k, v in dfs.items()}
    else:
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
            aplicar_validaciones_excel(writer)

db = cargar_datos()
hoy = datetime.now().date()

# --- FUNCIÓN DE INTELIGENCIA DE CRUCES GENÉTICOS ---
def analizar_cruces_razas(rep_df):
    """Evalúa el historial de cruces (Raza Cerda x Raza Semen) para detectar éxitos o fallas repetitivas."""
    alertas_malas = []
    alertas_buenas = []
    if rep_df.empty or "Raza_Cerda" not in rep_df.columns or "Macho_O_Semen" not in rep_df.columns:
        return alertas_malas, alertas_buenas

    df_val = rep_df[(rep_df["Raza_Cerda"].astype(str).str.strip() != "") & (rep_df["Macho_O_Semen"].astype(str).str.strip() != "")].copy()
    if df_val.empty:
        return alertas_malas, alertas_buenas

    df_val["Cruce"] = df_val["Raza_Cerda"].astype(str) + " × " + df_val["Macho_O_Semen"].astype(str)
    for cruce, grupo in df_val.groupby("Cruce"):
        total_serv = len(grupo)
        fallas = len(grupo[grupo["Estado_Prenez"].isin(["Repitió Celo", "Aborto"])])
        partos = grupo[grupo["Nacidos_Vivos"] > 0]
        prom_vivos = partos["Nacidos_Vivos"].mean() if not partos.empty else None

        if total_serv >= 2 and (fallas / total_serv) >= 0.5:
            alertas_malas.append(
                f"🧬 **Cruce de Riesgo ({cruce}):** De {total_serv} inseminaciones registradas, **{fallas} fallaron** (repetición de celo o aborto). Evita repetir esta combinación genética."
            )
        elif prom_vivos is not None and prom_vivos < 8.5 and len(partos) >= 1:
            alertas_malas.append(
                f"📉 **Baja Camada en Cruce ({cruce}):** Promedia apenas **{prom_vivos:.1f} nacidos vivos** por parto. Evalúa cambiar la raza del semen."
            )

        if prom_vivos is not None and prom_vivos >= 11.0 and fallas == 0:
            alertas_buenas.append(
                f"🏆 **Cruce Estrella ({cruce}):** Excelente compatibilidad genética con un promedio de **{prom_vivos:.1f} nacidos vivos** por camada y 0 fallas."
            )
    return alertas_malas, alertas_buenas

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("## 🐷 PorciGestión Pro")
    modo_txt = "☁️ Nube (Google Sheets)" if USAR_GSHEETS else "💻 Local (Excel)"
    st.caption(f"Conexión: {modo_txt}")
    st.caption(f"Perfil: **{'Administrador' if st.session_state.rol == 'admin' else 'Operario'}**")
    if USAR_GSHEETS:
        if st.button("🔄 Sincronizar con Google Sheets", width="stretch"):
            st.cache_data.clear()
            if "db_nube" in st.session_state:
                del st.session_state["db_nube"]
            st.rerun()
    st.divider()

    menu_opciones = [
        "🚨 Tablero de Alertas",
        "🧬 Reproducción e Inseminación",
        "🐖 Lotes, Madres y Unión",
        "🌾 Bodega y Alimentación",
        "💉 Sanidad y Mortalidad",
        "⚙️ Gestionar Cerdas y Tablas"
    ]
    if st.session_state.rol == "admin":
        menu_opciones.insert(5, "💰 Ventas (Pie / Canal)")
        menu_opciones.insert(6, "📊 Analítica e Inteligencia")

    seleccion = st.radio("Navegación:", menu_opciones)
    st.divider()
    if st.button("Cerrar Sesión", width="stretch"):
        st.session_state.autenticado = False
        st.rerun()

# --- KPIs SUPERIORES GLOBALES ---
df_madres = db["Reproductores"][db["Reproductores"]["Tipo"] == "Hembra Madre"] if not db["Reproductores"].empty else pd.DataFrame()
total_madres = len(df_madres)
gestantes = len(df_madres[df_madres["Estado"].isin(["Preñada", "Gestante"])]) if not df_madres.empty else 0
cerdos_engorde = int(db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["Cant_Actual"].sum()) if not db["Lotes_Engorde"].empty else 0
kilos_bodega = float(db["Bodega"]["Saldo_Kg"].sum()) if not db["Bodega"].empty else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("🐖 Cerdas Madres", f"{total_madres} hembras", f"{gestantes} preñadas")
k2.metric("🥩 Cerdos en Lotes", f"{cerdos_engorde} animales", "Activos en granja")
k3.metric("🌾 Inventario Bodega", f"{kilos_bodega:,.0f} kg", f"{kilos_bodega/40:.1f} bultos (40kg)")
k4.metric("📅 Fecha Operativa", hoy.strftime("%d/%m/%Y"), "Verraco: Solo recelador")
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 1. TABLERO DE ALERTAS INTELIGENTES
# ==========================================
if seleccion == "🚨 Tablero de Alertas":
    st.header("🚨 Centro de Alertas Operativas, Genéticas y Sanitarias")
    col1, col2, col3 = st.columns(3)

    rep = db["Reproduccion"]
    san = db["Sanidad_Eventos"]
    bod = db["Bodega"]
    lotes = db["Lotes_Engorde"]
    madres = db["Reproductores"]

    with col1:
        st.subheader("🤰 Maternidad y Cruces")
        alertas_rep = 0
        if not rep.empty:
            for _, row in rep[rep["Estado_Prenez"].isin(["Positiva", "Preñada"])].iterrows():
                if str(row["Fecha_Prob_Parto"]).strip() != "":
                    f_parto = pd.to_datetime(row["Fecha_Prob_Parto"]).date()
                    dias_faltan = (f_parto - hoy).days
                    if 0 <= dias_faltan <= 7:
                        st.error(f"🔴 **PARTO PRÓXIMO:** Cerda **{row['ID_Cerda']}** pare en **{dias_faltan} días** ({f_parto}). Preparar paridera.")
                        alertas_rep += 1
                    elif dias_faltan < 0:
                        st.warning(f"🟠 **PARTO ATRASADO:** Cerda **{row['ID_Cerda']}** tenía fecha {f_parto}. Registrar parto.")
                        alertas_rep += 1

            for _, row in rep[rep["Estado_Prenez"] == "Pendiente Eco"].iterrows():
                if str(row["Fecha_Eco_28d"]).strip() != "":
                    f_eco = pd.to_datetime(row["Fecha_Eco_28d"]).date()
                    if hoy >= f_eco:
                        st.info(f"🔵 **CONFIRMAR PREÑEZ:** Revisar celo/ecografía de **{row['ID_Cerda']}** (Inseminada el {row['Fecha_Servicio']}).")
                        alertas_rep += 1

        # Alertas de cruces por raza
        malas_razas, buenas_razas = analizar_cruces_razas(rep)
        for al_m in malas_razas:
            st.warning(al_m)
            alertas_rep += 1
        for al_b in buenas_razas:
            st.success(al_b)
            alertas_rep += 1

        for _, m in madres.iterrows():
            if m["Abortos"] >= 2 or m["Repeticiones"] >= 2:
                st.warning(f"⚠️ **EVALUAR DESCARTE:** **{m['ID_Animal']}** ({m['Linea_Genetica']}) acumula {m['Abortos']} abortos y {m['Repeticiones']} repeticiones.")
                alertas_rep += 1

        if alertas_rep == 0:
            st.success("✅ Sin alertas reproductivas urgentes hoy.")

    with col2:
        st.subheader("💉 Sanidad y Retiros")
        alertas_san = 0
        for _, s in san.iterrows():
            if pd.notna(s["Fecha_Fin_Retiro"]) and str(s["Fecha_Fin_Retiro"]).strip() != "":
                f_ret = pd.to_datetime(s["Fecha_Fin_Retiro"]).date()
                if hoy <= f_ret:
                    dias_r = (f_ret - hoy).days
                    st.error(f"🚫 **TIEMPO DE RETIRO:** **{s['ID_Destino']}** bloqueado por **{dias_r} días** más (Hasta {f_ret}).")
                    alertas_san += 1
            if s["Estado_Caso"] == "En Tratamiento":
                st.warning(f"🩺 **EN TRATAMIENTO:** **{s['ID_Destino']}** - {s['Detalle_Sintoma']} ({s['Medicamento']}).")
                alertas_san += 1
        if alertas_san == 0:
            st.success("✅ Todos los animales libres de retiro sanitario.")

    with col3:
        st.subheader("📦 Bodega y Lotes")
        for _, b in bod.iterrows():
            if b["Saldo_Kg"] < 120:
                st.error(f"📉 **STOCK BAJO:** Quedan **{b['Saldo_Kg']} kg** ({b['Saldo_Kg']/40:.1f} bultos) de **{b['Tipo_Concentrado']}**.")

        if not lotes.empty:
            for _, l in lotes[lotes["Estado"] == "Activo"].iterrows():
                if str(l["Fecha_Inicio"]).strip() != "":
                    f_ini = pd.to_datetime(l["Fecha_Inicio"]).date()
                    dias_lote = (hoy - f_ini).days
                    madre_txt = f" (Madre: {l['Madre_Origen']})" if str(l.get("Madre_Origen", "")).strip() != "" else ""
                    if l["Etapa"] == "Lactante" and dias_lote >= 28:
                        st.warning(f"🍼 **DESTETE LISTO:** **{l['ID_Lote']}**{madre_txt} cumplió {dias_lote} días. Destetar en Reproducción.")
                    elif l["Etapa"] == "Precebo" and dias_lote >= 42:
                        st.info(f"🔄 **CAMBIO DE ETAPA:** **{l['ID_Lote']}** lleva {dias_lote} días. Pasar a Levante.")
                    elif l["Etapa"] == "Levante" and dias_lote >= 45:
                        st.info(f"🔄 **CAMBIO DE ETAPA:** **{l['ID_Lote']}** listo para pasar a Engorde.")

# ==========================================
# 2. REPRODUCCIÓN E INSEMINACIÓN ARTIFICIAL
# ==========================================
elif seleccion == "🧬 Reproducción e Inseminación":
    st.header("🧬 Ciclo Reproductivo por Inseminación Artificial")
    st.info("ℹ️ **Flujo operativo:** El verraco se utiliza únicamente como **recelador** para estimular y detectar celo. Toda reproducción se realiza mediante **inseminación artificial** registrando la raza de la marrana y la raza del semen.")

    tab1, tab2, tab3, tab4 = st.tabs([
        "1️⃣ Registrar Inseminación",
        "2️⃣ Confirmar Preñez o Repetición",
        "3️⃣ Registrar Parto (Crea Lote)",
        "4️⃣ Registrar Destete (Pasa a Precebo)"
    ])

    hembras_df = db["Reproductores"][db["Reproductores"]["Tipo"] == "Hembra Madre"] if not db["Reproductores"].empty else pd.DataFrame()

    # --- PASO 1: INSEMINAR ---
    with tab1:
        st.subheader("Paso 1: Registrar Servicio de Inseminación")
        if hembras_df.empty:
            st.warning("No hay cerdas madres registradas. Agrégalas primero en '⚙️ Gestionar Cerdas y Tablas'.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                opciones_cerdas = [f"{r['ID_Animal']} ({r['Linea_Genetica']} - {r['Estado']})" for _, r in hembras_df.iterrows()]
                sel_cerda_str = st.selectbox("1. Selecciona la Marrana a inseminar:", opciones_cerdas)
                id_cerda = sel_cerda_str.split(" (")[0]
                cerda_info = hembras_df[hembras_df["ID_Animal"] == id_cerda].iloc[0]
                raza_cerda = str(cerda_info["Linea_Genetica"])

                f_serv = st.date_input("2. Fecha de la Inseminación:", hoy)
                raza_semen = st.selectbox("3. Raza del Semen utilizado (Pajilla):", OPCIONES_RAZAS)

            with c2:
                recelador = st.checkbox("¿Se usó verraco recelador para detectar celo?", value=True)
                costo_serv = st.number_input("4. Costo de las pajillas / inseminación ($):", min_value=0, value=45000, step=5000)
                f_eco_calc = f_serv + timedelta(days=28)
                f_parto_calc = f_serv + timedelta(days=114)
                st.markdown(f"""
                <div class="paso-card">
                    <b>📅 Proyección Automática para {id_cerda}:</b><br>
                    • Raza Marrana: <b>{raza_cerda}</b> × Semen: <b>{raza_semen}</b><br>
                    • Confirmación de preñez (28 días): <b>{f_eco_calc.strftime('%d/%m/%Y')}</b><br>
                    • Fecha probable de parto (114 días): <b>{f_parto_calc.strftime('%d/%m/%Y')}</b>
                </div>
                """, unsafe_allow_html=True)

            # Advertencia genética preventiva si cruzan razas iguales o con historial negativo
            if raza_cerda.strip().lower() == raza_semen.strip().lower():
                st.warning(f"⚠️ **Aviso Genético:** Estás cruzando **{raza_cerda} × {raza_semen}** (misma raza pura). Si has notado problemas de vigor o camadas pequeñas con este cruce puro, considera usar semen de otra raza (ej. Duroc o Pietrain) para lograr vigor híbrido.")

            if not db["Reproduccion"].empty:
                prev_cruce = db["Reproduccion"][
                    (db["Reproduccion"]["Raza_Cerda"] == raza_cerda) &
                    (db["Reproduccion"]["Macho_O_Semen"] == raza_semen) &
                    (db["Reproduccion"]["Estado_Prenez"].isin(["Repitió Celo", "Aborto"]))
                ]
                if len(prev_cruce) >= 2:
                    st.error(f"🚨 **Alerta de Historial:** El cruce **{raza_cerda} × {raza_semen}** ya registra **{len(prev_cruce)} fallas** previas en tu granja.")

            if st.button("✅ Guardar Inseminación"):
                nuevo_ev = {
                    "ID_Evento": f"INS-{len(db['Reproduccion'])+1:03d}",
                    "ID_Cerda": id_cerda,
                    "Raza_Cerda": raza_cerda,
                    "Fecha_Servicio": str(f_serv),
                    "Tipo_Servicio": "Inseminación Artificial",
                    "Macho_O_Semen": raza_semen,
                    "Uso_Recelador": "Sí (V-01)" if recelador else "No",
                    "Costo_Servicio": costo_serv,
                    "Fecha_Eco_28d": str(f_eco_calc),
                    "Estado_Prenez": "Pendiente Eco",
                    "Fecha_Prob_Parto": str(f_parto_calc),
                    "Fecha_Real_Parto": "",
                    "Nacidos_Vivos": 0,
                    "Mortinatos": 0,
                    "Momias": 0,
                    "Nombre_Lote_Asignado": "",
                    "Fecha_Prob_Destete": str(f_parto_calc + timedelta(days=28)),
                    "Destetados_Reales": 0,
                    "Peso_Destete_Total_Kg": 0,
                    "Costo_Por_Lechon": 0
                }
                db["Reproduccion"] = pd.concat([db["Reproduccion"], pd.DataFrame([nuevo_ev])], ignore_index=True)
                idx = db["Reproductores"][db["Reproductores"]["ID_Animal"] == id_cerda].index[0]
                db["Reproductores"].at[idx, "Costo_Acumulado_Ciclo"] += costo_serv
                db["Reproductores"].at[idx, "Estado"] = "Inseminada"
                guardar_datos(db)
                st.success(f"Inseminación registrada para {id_cerda}. Ahora pasa al Paso 2 cuando confirmes si quedó preñada.")
                st.rerun()

    # --- PASO 2: CONFIRMAR PREÑEZ ---
    with tab2:
        st.subheader("Paso 2: Confirmar si la Cerda Quedó Preñada o Repitió Celo")
        pendientes = db["Reproduccion"][db["Reproduccion"]["Estado_Prenez"] == "Pendiente Eco"] if not db["Reproduccion"].empty else pd.DataFrame()
        if pendientes.empty:
            st.info("No hay cerdas inseminadas pendientes por confirmar preñez en este momento.")
        else:
            opciones_pend = [
                f"{r['ID_Evento']} | Cerda: {r['ID_Cerda']} ({r['Raza_Cerda']} × {r['Macho_O_Semen']}) | Inseminada: {r['Fecha_Servicio']}"
                for _, r in pendientes.iterrows()
            ]
            sel_p = st.selectbox("Selecciona la cerda inseminada a evaluar:", opciones_pend)
            id_ev_p = sel_p.split(" | ")[0]
            idx_ev = db["Reproduccion"][db["Reproduccion"]["ID_Evento"] == id_ev_p].index[0]
            cerda_ev = db["Reproduccion"].at[idx_ev, "ID_Cerda"]
            idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == cerda_ev].index[0]

            resultado = st.radio("¿Cuál fue el resultado?", [
                "✅ Preñada Confirmada (Activar alerta de parto)",
                "❌ Repitió Celo (No quedó preñada)",
                "⚠️ Abortó"
            ])

            if st.button("Guardar Diagnóstico"):
                if "Preñada" in resultado:
                    db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = "Preñada"
                    db["Reproductores"].at[idx_m, "Estado"] = "Preñada"
                    guardar_datos(db)
                    st.success(f"¡Confirmado! {cerda_ev} queda registrada como PREÑADA. Fecha probable de parto: {db['Reproduccion'].at[idx_ev, 'Fecha_Prob_Parto']}.")
                else:
                    motivo = "Repitió Celo" if "Repitió" in resultado else "Aborto"
                    db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = motivo
                    db["Reproductores"].at[idx_m, "Estado"] = "Vacía"
                    if motivo == "Repitió Celo":
                        db["Reproductores"].at[idx_m, "Repeticiones"] += 1
                    else:
                        db["Reproductores"].at[idx_m, "Abortos"] += 1
                    guardar_datos(db)
                    st.warning(f"Registrado como '{motivo}'. El costo del semen y alimento acumulado se cargará a su próximo parto exitoso.")
                st.rerun()

    # --- PASO 3: REGISTRAR PARTO Y CREAR LOTE ---
    with tab3:
        st.subheader("Paso 3: Registrar Parto y Asignar Nombre al Lote de Lechones")
        prenadas = db["Reproduccion"][db["Reproduccion"]["Estado_Prenez"].isin(["Preñada", "Positiva"])] if not db["Reproduccion"].empty else pd.DataFrame()
        if prenadas.empty:
            st.info("No hay cerdas marcadas como 'Preñada' listas para parir. Confírmalas primero en el Paso 2.")
        else:
            opciones_pre = [
                f"{r['ID_Evento']} | Madre: {r['ID_Cerda']} | Cruce: {r['Raza_Cerda']} × {r['Macho_O_Semen']} | Parto Est: {r['Fecha_Prob_Parto']}"
                for _, r in prenadas.iterrows()
            ]
            sel_pre = st.selectbox("Selecciona la Cerda que parió:", opciones_pre)
            id_ev_pre = sel_pre.split(" | ")[0]
            idx_ev = db["Reproduccion"][db["Reproduccion"]["ID_Evento"] == id_ev_pre].index[0]
            cerda_ev = db["Reproduccion"].at[idx_ev, "ID_Cerda"]
            cruce_txt = f"{db['Reproduccion'].at[idx_ev, 'Raza_Cerda']} × {db['Reproduccion'].at[idx_ev, 'Macho_O_Semen']}"
            idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == cerda_ev].index[0]

            cp1, cp2 = st.columns(2)
            with cp1:
                f_parto_r = st.date_input("Fecha Real del Parto:", hoy)
                vivos = st.number_input("Lechones Nacidos Vivos:", min_value=1, value=11)
                muertos = st.number_input("Mortinatos (Nacidos muertos):", min_value=0, value=0)
                momias = st.number_input("Momias:", min_value=0, value=0)
            with cp2:
                sugerencia_lote = f"LOTE-{len(db['Lotes_Engorde'])+1:02d}-{cerda_ev}"
                nombre_lote = st.text_input("Nombre del Lote para estos cerditos:", value=sugerencia_lote)
                corral_paridera = st.text_input("Corral / Paridera:", value=f"Paridera-{cerda_ev}")
                peso_nac_prom = st.number_input("Peso promedio al nacer (kg):", min_value=0.5, value=1.4, step=0.1)

            if st.button("🐖 Registrar Parto y Crear Lote Lactante"):
                db["Reproduccion"].at[idx_ev, "Fecha_Real_Parto"] = str(f_parto_r)
                db["Reproduccion"].at[idx_ev, "Nacidos_Vivos"] = vivos
                db["Reproduccion"].at[idx_ev, "Mortinatos"] = muertos
                db["Reproduccion"].at[idx_ev, "Momias"] = momias
                db["Reproduccion"].at[idx_ev, "Nombre_Lote_Asignado"] = nombre_lote
                db["Reproduccion"].at[idx_ev, "Fecha_Prob_Destete"] = str(f_parto_r + timedelta(days=28))
                db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = "En Lactancia"

                costo_ciclo_madre = float(db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"])
                nuevo_lote = {
                    "ID_Lote": nombre_lote,
                    "Madre_Origen": cerda_ev,
                    "Raza_Cruce": cruce_txt,
                    "Corral": corral_paridera,
                    "Etapa": "Lactante",
                    "Fecha_Inicio": str(f_parto_r),
                    "Cant_Inicial": vivos,
                    "Cant_Actual": vivos,
                    "Peso_Prom_Inicial_Kg": peso_nac_prom,
                    "Peso_Prom_Actual_Kg": peso_nac_prom,
                    "Kg_Alimento_Consumido": 0.0,
                    "Costo_Inicial_Lechones": costo_ciclo_madre,
                    "Costo_Alimento": 0.0,
                    "Costo_Sanidad": 0.0,
                    "Estado": "Activo"
                }
                db["Lotes_Engorde"] = pd.concat([db["Lotes_Engorde"], pd.DataFrame([nuevo_lote])], ignore_index=True)

                db["Reproductores"].at[idx_m, "Estado"] = "Lactante"
                db["Reproductores"].at[idx_m, "Partos"] += 1
                db["Reproductores"].at[idx_m, "Racion_Diaria_Kg"] = 5.5
                db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] = 0.0
                guardar_datos(db)
                st.success(f"¡Parto registrado! Se creó automáticamente el lote **{nombre_lote}** (Madre: {cerda_ev}) con {vivos} lechones.")
                st.rerun()

    # --- PASO 4: DESTETE ---
    with tab4:
        st.subheader("Paso 4: Destetar Camada (28 días) y Pasar Lote a Precebo")
        lactantes = db["Reproduccion"][db["Reproduccion"]["Estado_Prenez"] == "En Lactancia"] if not db["Reproduccion"].empty else pd.DataFrame()
        if lactantes.empty:
            st.info("No hay cerdas en lactancia pendientes de destete.")
        else:
            opciones_lac = [
                f"{r['ID_Evento']} | Madre: {r['ID_Cerda']} | Lote: {r['Nombre_Lote_Asignado']} | Parto: {r['Fecha_Real_Parto']}"
                for _, r in lactantes.iterrows()
            ]
            sel_lac = st.selectbox("Selecciona la camada a destetar:", opciones_lac)
            id_ev_lac = sel_lac.split(" | ")[0]
            idx_ev = db["Reproduccion"][db["Reproduccion"]["ID_Evento"] == id_ev_lac].index[0]
            cerda_ev = db["Reproduccion"].at[idx_ev, "ID_Cerda"]
            lote_asoc = str(db["Reproduccion"].at[idx_ev, "Nombre_Lote_Asignado"])
            idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == cerda_ev].index[0]

            destetados = st.number_input("Lechones Destetados Vivos:", min_value=1, value=int(db["Reproduccion"].at[idx_ev, "Nacidos_Vivos"]))
            peso_tot = st.number_input("Peso Total de toda la camada al destete (kg):", min_value=5.0, value=float(destetados * 7.2))
            corral_dest = st.text_input("Corral de Precebo destino:", value="Corral-Precebo-1")

            if st.button("✅ Finalizar Destete y Pasar Lote a Precebo"):
                costo_lactancia_madre = float(db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"])
                db["Reproduccion"].at[idx_ev, "Destetados_Reales"] = destetados
                db["Reproduccion"].at[idx_ev, "Peso_Destete_Total_Kg"] = peso_tot
                db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = "Destetada"

                # Actualizar el lote ya creado en el parto
                if lote_asoc in db["Lotes_Engorde"]["ID_Lote"].values:
                    idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == lote_asoc].index[0]
                    costo_previo = float(db["Lotes_Engorde"].at[idx_l, "Costo_Inicial_Lechones"])
                    costo_total_dest = costo_previo + costo_lactancia_madre
                    db["Lotes_Engorde"].at[idx_l, "Etapa"] = "Precebo"
                    db["Lotes_Engorde"].at[idx_l, "Corral"] = corral_dest
                    db["Lotes_Engorde"].at[idx_l, "Fecha_Inicio"] = str(hoy)
                    db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] = destetados
                    db["Lotes_Engorde"].at[idx_l, "Peso_Prom_Actual_Kg"] = round(peso_tot / destetados, 2)
                    db["Lotes_Engorde"].at[idx_l, "Costo_Inicial_Lechones"] = costo_total_dest
                    db["Reproduccion"].at[idx_ev, "Costo_Por_Lechon"] = costo_total_dest / destetados

                db["Reproductores"].at[idx_m, "Estado"] = "Vacía"
                db["Reproductores"].at[idx_m, "Racion_Diaria_Kg"] = 2.2
                db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] = 0.0
                guardar_datos(db)
                st.success(f"Destete completado. El lote **{lote_asoc}** pasó a Precebo y la madre **{cerda_ev}** quedó Vacía lista para nuevo celo.")
                st.rerun()

    st.divider()
    st.subheader("📋 Historial Reproductivo Completo")
    st.dataframe(db["Reproduccion"], width="stretch")

# ==========================================
# 3. LOTES, MADRES Y UNIÓN DE LOTES
# ==========================================
elif seleccion == "🐖 Lotes, Madres y Unión":
    st.header("🐖 Control de Lotes, Trazabilidad de Madres y Unión de Grupos")

    col_top1, col_top2 = st.columns(2)

    # --- CREAR LOTE MANUAL CON MADRE ---
    with col_top1:
        with st.expander("➕ Crear Lote Manual (Asignando Madre y Nombre)", expanded=False):
            nombre_m = st.text_input("Nombre del Lote (ej. LOTE-01-M01):", f"LOTE-{len(db['Lotes_Engorde'])+1:02d}")
            lista_madres = ["Varias / Externa"] + (hembras_df["ID_Animal"].tolist() if not (hembras_df := db["Reproductores"][db["Reproductores"]["Tipo"] == "Hembra Madre"]).empty else [])
            madre_m = st.selectbox("Madre de este lote:", lista_madres)
            cruce_m = st.text_input("Raza / Cruce del lote:", "F1 × Duroc")
            corral_m = st.text_input("Corral:", "Corral-1")
            etapa_m = st.selectbox("Etapa Actual:", OPCIONES_ETAPA_LOTE)
            cant_m = st.number_input("Cantidad de Cerdos:", min_value=1, value=10)
            peso_m = st.number_input("Peso Promedio Hoy (kg):", min_value=1.0, value=15.0)
            costo_est_m = st.number_input("Costo acumulado hasta hoy ($):", min_value=0, value=1200000, step=50000)

            if st.button("Crear Lote"):
                if nombre_m in db["Lotes_Engorde"]["ID_Lote"].values:
                    st.error("Ya existe un lote con ese nombre. Elige otro identificador.")
                else:
                    nl = {
                        "ID_Lote": nombre_m, "Madre_Origen": madre_m, "Raza_Cruce": cruce_m,
                        "Corral": corral_m, "Etapa": etapa_m, "Fecha_Inicio": str(hoy),
                        "Cant_Inicial": cant_m, "Cant_Actual": cant_m,
                        "Peso_Prom_Inicial_Kg": peso_m, "Peso_Prom_Actual_Kg": peso_m,
                        "Kg_Alimento_Consumido": 0.0, "Costo_Inicial_Lechones": costo_est_m,
                        "Costo_Alimento": 0.0, "Costo_Sanidad": 0.0, "Estado": "Activo"
                    }
                    db["Lotes_Engorde"] = pd.concat([db["Lotes_Engorde"], pd.DataFrame([nl])], ignore_index=True)
                    guardar_datos(db)
                    st.success(f"Lote **{nombre_m}** (Madre: {madre_m}) creado.")
                    st.rerun()

    # --- UNIR / FUSIONAR LOTES ---
    with col_top2:
        with st.expander("🔗 Unir / Fusionar Dos o Más Lotes en Uno", expanded=False):
            lotes_activos_lista = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["ID_Lote"].tolist() if not db["Lotes_Engorde"].empty else []
            if len(lotes_activos_lista) < 2:
                st.info("Necesitas al menos 2 lotes activos para poder unirlos.")
            else:
                lotes_a_unir = st.multiselect("Selecciona los lotes que vas a juntar en el mismo corral:", lotes_activos_lista)
                nuevo_nombre_union = st.text_input("Nombre del nuevo lote unificado:", value="LOTE-UNIDO-01")
                nuevo_corral_union = st.text_input("Corral destino de la unión:", value="Corral-Engorde-1")
                nueva_etapa_union = st.selectbox("Etapa del grupo unido:", OPCIONES_ETAPA_LOTE, index=2)

                if st.button("🔗 Ejecutar Unión de Lotes"):
                    if len(lotes_a_unir) < 2:
                        st.error("Selecciona mínimo 2 lotes para unir.")
                    elif nuevo_nombre_union in lotes_activos_lista and nuevo_nombre_union not in lotes_a_unir:
                        st.error("Ese nombre ya lo usa otro lote activo diferente.")
                    else:
                        sub = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"].isin(lotes_a_unir)]
                        tot_ini = int(sub["Cant_Inicial"].sum())
                        tot_act = int(sub["Cant_Actual"].sum())
                        peso_tot_act = (sub["Cant_Actual"] * sub["Peso_Prom_Actual_Kg"]).sum()
                        peso_tot_ini = (sub["Cant_Inicial"] * sub["Peso_Prom_Inicial_Kg"]).sum()

                        madres_unidas = " + ".join(sub["Madre_Origen"].astype(str).unique())
                        cruces_unidos = " / ".join(sub["Raza_Cruce"].astype(str).unique())
                        obs_union = f"Unión de: {', '.join(lotes_a_unir)}"

                        # Marcar los lotes anteriores como Fusionados
                        for idx_u in sub.index:
                            db["Lotes_Engorde"].at[idx_u, "Estado"] = f"Unido en {nuevo_nombre_union}"
                            db["Lotes_Engorde"].at[idx_u, "Cant_Actual"] = 0

                        lote_fusionado = {
                            "ID_Lote": nuevo_nombre_union,
                            "Madre_Origen": f"{madres_unidas} ({obs_union})",
                            "Raza_Cruce": cruces_unidos,
                            "Corral": nuevo_corral_union,
                            "Etapa": nueva_etapa_union,
                            "Fecha_Inicio": str(hoy),
                            "Cant_Inicial": tot_ini,
                            "Cant_Actual": tot_act,
                            "Peso_Prom_Inicial_Kg": round(peso_tot_ini / tot_ini, 2) if tot_ini > 0 else 0,
                            "Peso_Prom_Actual_Kg": round(peso_tot_act / tot_act, 2) if tot_act > 0 else 0,
                            "Kg_Alimento_Consumido": float(sub["Kg_Alimento_Consumido"].sum()),
                            "Costo_Inicial_Lechones": float(sub["Costo_Inicial_Lechones"].sum()),
                            "Costo_Alimento": float(sub["Costo_Alimento"].sum()),
                            "Costo_Sanidad": float(sub["Costo_Sanidad"].sum()),
                            "Estado": "Activo"
                        }
                        db["Lotes_Engorde"] = pd.concat([db["Lotes_Engorde"], pd.DataFrame([lote_fusionado])], ignore_index=True)
                        guardar_datos(db)
                        st.success(f"¡Lotes unidos! Se creó **{nuevo_nombre_union}** con {tot_act} cerdos y todos sus costos acumulados sumados.")
                        st.rerun()

    lotes_act = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"] if not db["Lotes_Engorde"].empty else pd.DataFrame()
    if not lotes_act.empty:
        st.subheader("Actualizar Pesaje o Etapa de un Lote")
        col_p1, col_p2, col_p3 = st.columns(3)
        lote_sel = col_p1.selectbox("Lote a actualizar:", lotes_act["ID_Lote"].tolist())
        nuevo_peso = col_p2.number_input("Nuevo Peso Promedio (kg):", min_value=0.5, value=50.0)
        nueva_etapa = col_p3.selectbox("Etapa:", OPCIONES_ETAPA_LOTE)
        if st.button("Actualizar Lote"):
            idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == lote_sel].index[0]
            db["Lotes_Engorde"].at[idx_l, "Peso_Prom_Actual_Kg"] = nuevo_peso
            db["Lotes_Engorde"].at[idx_l, "Etapa"] = nueva_etapa
            guardar_datos(db)
            st.success("Pesaje y etapa actualizados.")
            st.rerun()

        st.subheader("📊 Indicadores en Tiempo Real por Lote y Madre")
        tabla_kpi = []
        for _, r in lotes_act.iterrows():
            costo_total_lote = r["Costo_Inicial_Lechones"] + r["Costo_Alimento"] + r["Costo_Sanidad"]
            costo_por_cerdo_vivo = costo_total_lote / r["Cant_Actual"] if r["Cant_Actual"] > 0 else 0
            kilos_totales_vivos = r["Cant_Actual"] * r["Peso_Prom_Actual_Kg"]
            costo_por_kilo_pie = costo_total_lote / kilos_totales_vivos if kilos_totales_vivos > 0 else 0

            ganancia_peso_lote = (r["Peso_Prom_Actual_Kg"] - r["Peso_Prom_Inicial_Kg"]) * r["Cant_Actual"]
            conv_alim = r["Kg_Alimento_Consumido"] / ganancia_peso_lote if ganancia_peso_lote > 0 else 0

            fila = {
                "Lote": r["ID_Lote"],
                "Madre / Origen": r.get("Madre_Origen", ""),
                "Cruce": r.get("Raza_Cruce", ""),
                "Corral": r["Corral"],
                "Etapa": r["Etapa"],
                "Vivos / Inicial": f"{r['Cant_Actual']} / {r['Cant_Inicial']}",
                "Peso Prom (kg)": r["Peso_Prom_Actual_Kg"],
                "Conv. Alimenticia": round(conv_alim, 2)
            }
            if st.session_state.rol == "admin":
                fila["Costo Total Lote ($)"] = f"${costo_total_lote:,.0f}"
                fila["Costo x Cerdo ($)"] = f"${costo_por_cerdo_vivo:,.0f}"
                fila["Punto Equilibrio ($/kg)"] = f"${costo_por_kilo_pie:,.0f}"
            tabla_kpi.append(fila)
        st.dataframe(pd.DataFrame(tabla_kpi), width="stretch")
    else:
        st.info("No hay lotes activos. Se crean automáticamente cuando registras un parto o puedes crear uno manual arriba.")

# ==========================================
# 4. BODEGA Y ALIMENTACIÓN
# ==========================================
elif seleccion == "🌾 Bodega y Alimentación":
    st.header("🌾 Control de Bodega Central y Consumo por Grupo")
    b1, b2 = st.columns(2)

    with b1:
        st.subheader("1. Ingresar Compra de Bultos a Bodega")
        tipo_conc = st.selectbox("Tipo de Concentrado", OPCIONES_CONCENTRADO)
        bultos_c = st.number_input("Cantidad de Bultos Comprados", min_value=1, value=10)
        peso_b = st.number_input("Peso por Bulto (Kg)", min_value=1, value=40)
        precio_b = st.number_input("Precio por Bulto ($)", min_value=1000, value=85000, step=1000)

        if st.button("Registrar Entrada a Bodega"):
            kilos_tot = bultos_c * peso_b
            costo_kg = precio_b / peso_b
            nueva_c = {
                "ID_Compra": f"C-{len(db['Bodega'])+1:03d}", "Fecha": str(hoy),
                "Tipo_Concentrado": tipo_conc, "Bultos": bultos_c, "Peso_Bulto_Kg": peso_b,
                "Precio_Bulto": precio_b, "Costo_Kg": costo_kg, "Saldo_Kg": kilos_tot
            }
            db["Bodega"] = pd.concat([db["Bodega"], pd.DataFrame([nueva_c])], ignore_index=True)
            guardar_datos(db)
            st.success(f"Ingresados {kilos_tot} kg de {tipo_conc} (${costo_kg:,.0f}/kg).")
            st.rerun()

    with b2:
        st.subheader("2. Registrar Consumo (Bodega ➡️ Corral o Cerda)")
        dest_tipo = st.radio("¿A quién se le dio el alimento?", ["Lote de Engorde", "Cerda Madre / Verraco"])
        if dest_tipo == "Lote de Engorde":
            opciones_dest = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["ID_Lote"].tolist() if not db["Lotes_Engorde"].empty else []
        else:
            opciones_dest = db["Reproductores"]["ID_Animal"].tolist() if not db["Reproductores"].empty else []

        if opciones_dest:
            id_dest = st.selectbox("Seleccionar Destino", opciones_dest)
            conc_disponibles = db["Bodega"][db["Bodega"]["Saldo_Kg"] > 0]["Tipo_Concentrado"].unique().tolist()
            if conc_disponibles:
                conc_usar = st.selectbox("Concentrado sacado de bodega", conc_disponibles)
                unidad = st.radio("Unidad de medida", ["Kilos (kg)", "Bultos (40 kg)"])
                cant_ing = st.number_input("Cantidad entregada", min_value=0.5, value=40.0 if unidad == "Kilos (kg)" else 1.0, step=0.5)
                kilos_reales = cant_ing if unidad == "Kilos (kg)" else cant_ing * 40.0

                if st.button("Descontar de Bodega y Cargar Costo"):
                    idx_bod = db["Bodega"][(db["Bodega"]["Tipo_Concentrado"] == conc_usar) & (db["Bodega"]["Saldo_Kg"] > 0)].index
                    saldo_total_tipo = db["Bodega"].loc[idx_bod, "Saldo_Kg"].sum()

                    if kilos_reales > saldo_total_tipo:
                        st.error(f"No hay suficiente {conc_usar} en bodega. Saldo disponible: {saldo_total_tipo} kg.")
                    else:
                        costo_kilo_prom = db["Bodega"].loc[idx_bod, "Costo_Kg"].mean()
                        costo_consumo = kilos_reales * costo_kilo_prom

                        por_descontar = kilos_reales
                        for i in idx_bod:
                            disp = db["Bodega"].at[i, "Saldo_Kg"]
                            if disp >= por_descontar:
                                db["Bodega"].at[i, "Saldo_Kg"] -= por_descontar
                                break
                            else:
                                por_descontar -= disp
                                db["Bodega"].at[i, "Saldo_Kg"] = 0

                        if dest_tipo == "Lote de Engorde":
                            idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == id_dest].index[0]
                            db["Lotes_Engorde"].at[idx_l, "Kg_Alimento_Consumido"] += kilos_reales
                            db["Lotes_Engorde"].at[idx_l, "Costo_Alimento"] += costo_consumo
                        else:
                            idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == id_dest].index[0]
                            db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] += costo_consumo

                        reg_cons = {
                            "Fecha": str(hoy), "Destino_Tipo": dest_tipo, "ID_Destino": id_dest,
                            "Tipo_Concentrado": conc_usar, "Kilos_Servidos": kilos_reales, "Costo_Total": costo_consumo
                        }
                        db["Consumo_Alimento"] = pd.concat([db["Consumo_Alimento"], pd.DataFrame([reg_cons])], ignore_index=True)
                        guardar_datos(db)
                        st.success(f"Descontados {kilos_reales} kg (${costo_consumo:,.0f}) y cargados a {id_dest}.")
                        st.rerun()
            else:
                st.warning("Bodega vacía. Ingresa compras primero.")
        else:
            st.info("No hay lotes activos aún.")

    st.subheader("Inventario Actual en Bodega")
    st.dataframe(db["Bodega"], width="stretch")

# ==========================================
# 5. SANIDAD Y MORTALIDAD
# ==========================================
elif seleccion == "💉 Sanidad y Mortalidad":
    st.header("💉 Bitácora Sanitaria, Vacunas y Bajas (Mortalidad)")
    s1, s2 = st.columns(2)

    lotes_act_ids = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["ID_Lote"].tolist() if not db["Lotes_Engorde"].empty else []
    repro_ids = db["Reproductores"]["ID_Animal"].tolist() if not db["Reproductores"].empty else []
    destinos_todos = lotes_act_ids + repro_ids

    with s1:
        st.subheader("Registrar Evento Sanitario o Mortalidad")
        if destinos_todos:
            dest_s = st.selectbox("Animal o Lote afectado", destinos_todos)
            cat_s = st.selectbox("Tipo de Registro", ["Vacunación Preventiva", "Tratamiento / Enfermedad", "Mortalidad (Baja)", "Observación General"])
            detalle_s = st.text_input("Síntoma, Vacuna o Causa de Muerte", "Ej. Vacuna Circovirus / Diarrea / Aplastamiento")
            med_s = st.text_input("Medicamento aplicado (si aplica)", "Ninguno")
            costo_s = st.number_input("Costo del tratamiento/vacuna ($)", min_value=0, value=0, step=1000)
            retiro_s = st.number_input("Días de retiro antes de sacrificio", min_value=0, value=0)
            muertos_s = st.number_input("Cantidad de animales muertos (solo si es Mortalidad)", min_value=0, value=1 if cat_s == "Mortalidad (Baja)" else 0)

            if st.button("Guardar Registro Sanitario"):
                f_fin_ret = str(hoy + timedelta(days=retiro_s)) if retiro_s > 0 else ""
                estado_c = "En Tratamiento" if cat_s == "Tratamiento / Enfermedad" else "Cerrado"

                nuevo_s = {
                    "Fecha": str(hoy), "ID_Destino": dest_s, "Categoria": cat_s,
                    "Detalle_Sintoma": detalle_s, "Medicamento": med_s, "Costo": costo_s,
                    "Dias_Retiro": retiro_s, "Fecha_Fin_Retiro": f_fin_ret, "Estado_Caso": estado_c
                }
                db["Sanidad_Eventos"] = pd.concat([db["Sanidad_Eventos"], pd.DataFrame([nuevo_s])], ignore_index=True)

                if not db["Lotes_Engorde"].empty and dest_s in db["Lotes_Engorde"]["ID_Lote"].values:
                    idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == dest_s].index[0]
                    db["Lotes_Engorde"].at[idx_l, "Costo_Sanidad"] += costo_s
                    if cat_s == "Mortalidad (Baja)" and muertos_s > 0:
                        db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] = max(0, db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] - muertos_s)
                        st.warning(f"⚠️ Se restaron {muertos_s} cerdos de {dest_s}. El costo se redistribuyó entre los sobrevivientes.")
                elif not db["Reproductores"].empty and dest_s in db["Reproductores"]["ID_Animal"].values:
                    idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == dest_s].index[0]
                    db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] += costo_s

                guardar_datos(db)
                st.success("Evento sanitario registrado.")
                st.rerun()
        else:
            st.info("Registra primero cerdas o lotes para asociar eventos sanitarios.")

    with s2:
        st.subheader("Historial Sanitario y Observaciones")
        st.dataframe(db["Sanidad_Eventos"], width="stretch")

# ==========================================
# 6. VENTAS (PIE / CANAL) - SOLO ADMIN
# ==========================================
elif seleccion == "💰 Ventas (Pie / Canal)":
    st.header("💰 Calculadora de Venta (En Pie vs. Beneficiado) y Balance")
    lotes_v = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"] if not db["Lotes_Engorde"].empty else pd.DataFrame()
    if not lotes_v.empty:
        v1, v2 = st.columns(2)
        with v1:
            st.subheader("Simular y Registrar Venta")
            lote_venta = st.selectbox("Lote a vender", lotes_v["ID_Lote"].tolist())
            row_lv = lotes_v[lotes_v["ID_Lote"] == lote_venta].iloc[0]

            bloqueado = False
            for _, s in db["Sanidad_Eventos"][db["Sanidad_Eventos"]["ID_Destino"] == lote_venta].iterrows():
                if pd.notna(s["Fecha_Fin_Retiro"]) and str(s["Fecha_Fin_Retiro"]).strip() != "":
                    if hoy <= pd.to_datetime(s["Fecha_Fin_Retiro"]).date():
                        st.error(f"🚫 VENTA BLOQUEADA: Este lote tiene tiempo de retiro vigente hasta {s['Fecha_Fin_Retiro']}.")
                        bloqueado = True

            max_c = max(1, int(row_lv["Cant_Actual"]))
            cant_v = st.number_input("Cantidad de cerdos a vender", min_value=1, max_value=max_c, value=max_c)
            modalidad = st.radio("Modalidad de Venta", ["En Pie (Cerdo entero en báscula)", "Beneficiado (Carne en canal por kilos)"])
            peso_pie_tot = st.number_input("Peso Total en Pie (kg)", min_value=5.0, value=float(cant_v * row_lv["Peso_Prom_Actual_Kg"]))

            costo_tot_lote = row_lv["Costo_Inicial_Lechones"] + row_lv["Costo_Alimento"] + row_lv["Costo_Sanidad"]
            costo_proporcional = (costo_tot_lote / row_lv["Cant_Actual"]) * cant_v if row_lv["Cant_Actual"] > 0 else 0

            if modalidad == "En Pie (Cerdo entero en báscula)":
                precio_kg_pie = st.number_input("Precio de venta por Kg en Pie ($)", min_value=1000, value=8500, step=100)
                ingreso_bruto = peso_pie_tot * precio_kg_pie
                costo_sacrificio = 0
                peso_canal = 0
            else:
                rendimiento = st.slider("Rendimiento en Canal (%)", 70, 82, 78)
                peso_canal = peso_pie_tot * (rendimiento / 100.0)
                costo_sacrificio = st.number_input("Costo total de sacrificio / planta / transporte ($)", min_value=0, value=cant_v * 65000, step=5000)
                precio_kg_carne = st.number_input("Precio de venta por Kg de Carne en Canal ($)", min_value=1000, value=14500, step=200)
                ingreso_bruto = peso_canal * precio_kg_carne

            utilidad_est = ingreso_bruto - costo_proporcional - costo_sacrificio

            st.info(f"""
**Resumen Financiero de esta Operación:**
- Costo de producción acumulado: **${costo_proporcional:,.0f}**
- Ingreso bruto esperado: **${ingreso_bruto:,.0f}**
- **Utilidad Neta Real: ${utilidad_est:,.0f}**
            """)

            if st.button("Confirmar y Registrar Venta", disabled=bloqueado):
                nueva_v = {
                    "Fecha": str(hoy), "ID_Origen": lote_venta, "Modalidad": modalidad,
                    "Cantidad_Animales": cant_v, "Peso_Pie_Kg": peso_pie_tot,
                    "Peso_Carne_Canal_Kg": peso_canal, "Costo_Sacrificio": costo_sacrificio,
                    "Ingreso_Bruto": ingreso_bruto, "Costo_Produccion_Asignado": costo_proporcional,
                    "Utilidad_Neta": utilidad_est
                }
                db["Ventas"] = pd.concat([db["Ventas"], pd.DataFrame([nueva_v])], ignore_index=True)

                idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == lote_venta].index[0]
                restantes = row_lv["Cant_Actual"] - cant_v
                db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] = restantes
                factor = restantes / row_lv["Cant_Actual"] if row_lv["Cant_Actual"] > 0 else 0
                db["Lotes_Engorde"].at[idx_l, "Costo_Inicial_Lechones"] *= factor
                db["Lotes_Engorde"].at[idx_l, "Costo_Alimento"] *= factor
                db["Lotes_Engorde"].at[idx_l, "Costo_Sanidad"] *= factor
                if restantes == 0:
                    db["Lotes_Engorde"].at[idx_l, "Estado"] = "Vendido"

                guardar_datos(db)
                st.success("¡Venta registrada con éxito!")
                st.rerun()

        with v2:
            st.subheader("Historial de Ventas y Utilidades")
            st.dataframe(db["Ventas"], width="stretch")
            if not db["Ventas"].empty:
                st.metric("Ingresos Totales Acumulados", f"${db['Ventas']['Ingreso_Bruto'].sum():,.0f}")
                st.metric("Utilidad Neta Total Acumulada", f"${db['Ventas']['Utilidad_Neta'].sum():,.0f}")
    else:
        st.info("Crea o desteta un lote primero para simular ventas.")

# ==========================================
# 7. ANALÍTICA E INTELIGENCIA DE DATOS
# ==========================================
elif seleccion == "📊 Analítica e Inteligencia":
    st.header("📊 Motor de Analítica Zootécnica y Cruces Genéticos")
    st.caption("Analiza automáticamente qué cruces de razas (Marrana × Semen) te dan más lechones, cuáles fallan, y cómo rinden tus lotes.")

    hallazgos_encontrados = 0
    lotes_df = db["Lotes_Engorde"].copy()
    rep_df = db["Reproduccion"].copy()
    madres_df = db["Reproductores"].copy()
    san_df = db["Sanidad_Eventos"].copy()
    ventas_df = db["Ventas"].copy()

    # 1. Alertas de cruces entre razas
    malas_razas, buenas_razas = analizar_cruces_razas(rep_df)
    for m_r in malas_razas:
        st.error(m_r)
        hallazgos_encontrados += 1
    for b_r in buenas_razas:
        st.success(b_r)
        hallazgos_encontrados += 1

    # 2. Eficiencia individual por cerda
    if not madres_df.empty:
        for _, m in madres_df[madres_df["Tipo"] == "Hembra Madre"].iterrows():
            fallas = m["Abortos"] + m["Repeticiones"]
            if fallas >= 2:
                st.warning(f"🧬 **Eficiencia Reproductiva ({m['ID_Animal']} - {m['Linea_Genetica']}):** Acumula **{fallas} fallas** ({m['Repeticiones']} repeticiones, {m['Abortos']} abortos) y **${m['Costo_Acumulado_Ciclo']:,.0f}** en costos acumulados. **Recomendación:** Evaluar descarte.")
                hallazgos_encontrados += 1

    # 3. Conversión alimenticia por lote
    if not lotes_df.empty:
        lotes_df["Ganancia_Total"] = (lotes_df["Peso_Prom_Actual_Kg"] - lotes_df["Peso_Prom_Inicial_Kg"]) * lotes_df["Cant_Actual"]
        lotes_df["CA"] = lotes_df.apply(lambda r: r["Kg_Alimento_Consumido"] / r["Ganancia_Total"] if r["Ganancia_Total"] > 0 else 0, axis=1)
        lotes_validos = lotes_df[lotes_df["CA"] > 0]
        if len(lotes_validos) >= 2:
            ca_prom_granja = lotes_validos["CA"].mean()
            ineficientes = lotes_validos[lotes_validos["CA"] > ca_prom_granja * 1.15]
            for _, l_malo in ineficientes.iterrows():
                st.warning(f"📉 **Alerta de Conversión ({l_malo['ID_Lote']}):** Conversión de **{l_malo['CA']:.2f}** vs promedio de granja ({ca_prom_granja:.2f}).")
                hallazgos_encontrados += 1

    # 4. Patrones sanitarios
    if not san_df.empty:
        bajas = san_df[san_df["Categoria"] == "Mortalidad (Baja)"]
        if len(bajas) >= 2:
            causa_top = bajas["Detalle_Sintoma"].value_counts().idxmax()
            cant_causa = bajas["Detalle_Sintoma"].value_counts().max()
            st.warning(f"🩺 **Patrón Sanitario:** La causa principal de mortalidad es **'{causa_top}'** ({cant_causa} casos).")
            hallazgos_encontrados += 1

    # 5. Rentabilidad comercial
    if not ventas_df.empty:
        ventas_df["Mes"] = pd.to_datetime(ventas_df["Fecha"]).dt.strftime("%Y-%m")
        mejor_mod = ventas_df.groupby("Modalidad")["Utilidad_Neta"].mean()
        if len(mejor_mod) > 1:
            top_mod = mejor_mod.idxmax()
            diff = mejor_mod.max() - mejor_mod.min()
            st.success(f"💰 **Estrategia Comercial:** Vender en **{top_mod}** genera en promedio **${diff:,.0f} más de utilidad neta** por operación.")
            hallazgos_encontrados += 1

    if hallazgos_encontrados == 0:
        st.info("🌱 **Motor de Inteligencia Activo:** Registra inseminaciones, partos, pesajes y ventas para que el sistema compare tus cruces de razas y rentabilidad.")

    st.divider()
    g1, g2 = st.columns(2)
    with g1:
        st.subheader("1. Rendimiento por Cruce (Raza Cerda × Semen)")
        if not rep_df.empty and rep_df["Nacidos_Vivos"].sum() > 0:
            rep_df["Cruce"] = rep_df["Raza_Cerda"].astype(str) + " × " + rep_df["Macho_O_Semen"].astype(str)
            res_cruce = rep_df[rep_df["Nacidos_Vivos"] > 0].groupby("Cruce")[["Nacidos_Vivos", "Destetados_Reales"]].mean().reset_index()
            fig_cruce = px.bar(res_cruce, x="Cruce", y=["Nacidos_Vivos", "Destetados_Reales"], barmode="group", color_discrete_sequence=["#2E86C1", "#1E8449"])
            st.plotly_chart(fig_cruce, width="stretch")
        else:
            st.info("Registra partos para ver qué cruce de razas produce más lechones.")

    with g2:
        st.subheader("2. Distribución de Costos de Alimento ($)")
        if not db["Consumo_Alimento"].empty:
            fig_cons = px.pie(db["Consumo_Alimento"], values="Costo_Total", names="Tipo_Concentrado", hole=0.45, color_discrete_sequence=px.colors.sequential.Greens_r)
            st.plotly_chart(fig_cons, width="stretch")
        else:
            st.info("Registra consumos de alimento para visualizar costos.")

# ==========================================
# 8. GESTIONAR CERDAS, VERRACO Y TABLAS
# ==========================================
elif seleccion == "⚙️ Gestionar Cerdas y Tablas":
    st.header("⚙️ Agregar, Editar o Eliminar Cerdas, Verracos y Lotes")

    col_add, col_del = st.columns(2)

    # --- FORMULARIO PARA AGREGAR NUEVA CERDA O VERRACO ---
    with col_add:
        st.subheader("➕ Registrar Nueva Cerda o Verraco")
        with st.form("form_nuevo_animal"):
            nuevo_id = st.text_input("Código del Animal (ej. M-15 o V-02):", placeholder="M-15")
            nuevo_tipo = st.selectbox("Tipo de Animal:", OPCIONES_TIPO_REPRO)
            nueva_raza = st.selectbox("Raza / Genética de la Marrana o Verraco:", OPCIONES_RAZAS)
            nuevo_estado = st.selectbox("Estado Inicial:", OPCIONES_ESTADO_REPRO)
            nueva_racion = st.number_input("Ración Diaria (kg):", min_value=0.5, max_value=10.0, value=2.2, step=0.1)
            nueva_obs = st.text_input("Observaciones:", value="Ingreso nuevo")
            btn_crear_animal = st.form_submit_button("💾 Guardar Nuevo Animal")

            if btn_crear_animal:
                id_limpio = nuevo_id.strip().upper()
                if not id_limpio:
                    st.error("Escribe un código válido para el animal (ej. M-15).")
                elif not db["Reproductores"].empty and id_limpio in db["Reproductores"]["ID_Animal"].values:
                    st.error(f"El código **{id_limpio}** ya existe en tu inventario.")
                else:
                    fila_nueva = {
                        "ID_Animal": id_limpio,
                        "Tipo": nuevo_tipo,
                        "Linea_Genetica": nueva_raza,
                        "Padre_ID": "Inseminación",
                        "Madre_ID": "Externa",
                        "Partos": 0,
                        "Abortos": 0,
                        "Repeticiones": 0,
                        "Estado": nuevo_estado,
                        "Racion_Diaria_Kg": nueva_racion,
                        "Costo_Acumulado_Ciclo": 0.0,
                        "Observaciones": nueva_obs
                    }
                    db["Reproductores"] = pd.concat([db["Reproductores"], pd.DataFrame([fila_nueva])], ignore_index=True)
                    guardar_datos(db)
                    st.success(f"¡Animal **{id_limpio}** ({nueva_raza}) registrado y guardado en Google Sheets!")
                    st.rerun()

    # --- ELIMINAR CERDAS O VERRACOS ---
    with col_del:
        st.subheader("🗑️ Eliminar Cerdas o Registros")
        if db["Reproductores"].empty:
            st.info("No hay animales registrados actualmente.")
        else:
            lista_animales = db["Reproductores"]["ID_Animal"].tolist()
            seleccion_borrar = st.multiselect(
                "Selecciona una o varias cerdas/animales que deseas eliminar:",
                lista_animales
            )
            if st.button("🗑️ Eliminar Seleccionados Definitivamente"):
                if not seleccion_borrar:
                    st.warning("Selecciona al menos un animal de la lista.")
                else:
                    db["Reproductores"] = db["Reproductores"][~db["Reproductores"]["ID_Animal"].isin(seleccion_borrar)].reset_index(drop=True)
                    guardar_datos(db)
                    st.success(f"Se eliminaron: {', '.join(seleccion_borrar)} de la app y de Google Sheets.")
                    st.rerun()

    st.divider()
    st.subheader("📋 Tabla Editable de Cerdas y Verracos")
    st.caption("Puedes modificar directamente la **Raza (Linea_Genetica)**, el **Estado** o cualquier celda y luego presionar el botón de guardar.")
    df_repro_editado = st.data_editor(
        db["Reproductores"],
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=OPCIONES_TIPO_REPRO, required=True),
            "Linea_Genetica": st.column_config.SelectboxColumn("Raza / Genética", options=OPCIONES_RAZAS, required=True),
            "Estado": st.column_config.SelectboxColumn("Estado Actual", options=OPCIONES_ESTADO_REPRO, required=True),
            "Racion_Diaria_Kg": st.column_config.NumberColumn("Ración Diaria (kg)", min_value=0.0, max_value=12.0, step=0.1),
        }
    )
    if st.button("💾 Guardar Cambios de la Tabla de Reproductores"):
        db["Reproductores"] = df_repro_editado.dropna(subset=["ID_Animal"]).reset_index(drop=True)
        guardar_datos(db)
        st.success("¡Cambios guardados en Google Sheets!")
        st.rerun()

    st.divider()
    st.subheader("🐖 Tabla Editable de Lotes (Eliminar o Corregir Lotes)")
    df_lotes_editado = st.data_editor(
        db["Lotes_Engorde"],
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Etapa": st.column_config.SelectboxColumn("Etapa", options=OPCIONES_ETAPA_LOTE, required=True),
        }
    )
    if st.button("💾 Guardar Cambios en Lotes"):
        db["Lotes_Engorde"] = df_lotes_editado.dropna(subset=["ID_Lote"]).reset_index(drop=True)
        guardar_datos(db)
        st.success("¡Lotes actualizados en Google Sheets!")
        st.rerun()