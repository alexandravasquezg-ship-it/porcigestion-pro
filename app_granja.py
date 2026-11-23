import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

            if st.button("Ingresar al Sistema", use_container_width=True):
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

OPCIONES_ESTADO_REPRO = ["Vacía", "Servida", "Gestante", "Lactante", "Levante", "Descartada"]
OPCIONES_TIPO_REPRO = ["Hembra Madre", "Verraco (En crecimiento)", "Verraco Adulto"]
OPCIONES_ETAPA_LOTE = ["Lactante", "Precebo", "Levante", "Engorde"]
OPCIONES_CONCENTRADO = ["Preiniciador", "Iniciador", "Levante", "Engorde", "Gestación", "Lactancia"]

def crear_estructura_base():
    cerdas_init = []
    for i in range(1, 15):
        cerdas_init.append({
            "ID_Animal": f"M-{i:02d}", "Tipo": "Hembra Madre", "Linea_Genetica": "Landrace x Yorkshire",
            "Padre_ID": "Externo", "Madre_ID": "Externa", "Partos": 1, "Abortos": 0,
            "Repeticiones": 0, "Estado": "Vacía", "Racion_Diaria_Kg": 2.2,
            "Costo_Acumulado_Ciclo": 0.0, "Observaciones": "Inventario inicial"
        })
    cerdas_init.append({
        "ID_Animal": "V-01", "Tipo": "Verraco (En crecimiento)", "Linea_Genetica": "Duroc / Pietrain",
        "Padre_ID": "Externo", "Madre_ID": "M-01", "Partos": 0, "Abortos": 0,
        "Repeticiones": 0, "Estado": "Levante", "Racion_Diaria_Kg": 1.8,
        "Costo_Acumulado_Ciclo": 0.0, "Observaciones": "Futuro reproductor - No cruzar con M-01"
    })
    return {
        "Reproductores": pd.DataFrame(cerdas_init),
        "Reproduccion": pd.DataFrame(columns=[
            "ID_Evento", "ID_Cerda", "Fecha_Servicio", "Tipo_Servicio", "Macho_O_Semen",
            "Costo_Servicio", "Fecha_Eco_28d", "Estado_Prenez", "Fecha_Prob_Parto",
            "Fecha_Real_Parto", "Nacidos_Vivos", "Mortinatos", "Momias",
            "Fecha_Prob_Destete", "Destetados_Reales", "Peso_Destete_Total_Kg", "Costo_Por_Lechon"
        ]),
        "Lotes_Engorde": pd.DataFrame(columns=[
            "ID_Lote", "Corral", "Etapa", "Fecha_Inicio", "Cant_Inicial", "Cant_Actual",
            "Peso_Prom_Inicial_Kg", "Peso_Prom_Actual_Kg", "Kg_Alimento_Consumido",
            "Costo_Inicial_Lechones", "Costo_Alimento", "Costo_Sanidad", "Estado"
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

def aplicar_validaciones_excel(writer):
    wb = writer.book
    ws_rep = wb["Reproductores"]
    dv_tipo = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_TIPO_REPRO)}"', allow_blank=True)
    dv_estado = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_ESTADO_REPRO)}"', allow_blank=True)
    ws_rep.add_data_validation(dv_tipo)
    ws_rep.add_data_validation(dv_estado)
    dv_tipo.add("B2:B100")
    dv_estado.add("I2:I100")

    ws_lot = wb["Lotes_Engorde"]
    dv_etapa = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_ETAPA_LOTE)}"', allow_blank=True)
    ws_lot.add_data_validation(dv_etapa)
    dv_etapa.add("C2:C100")

    ws_bod = wb["Bodega"]
    dv_conc = DataValidation(type="list", formula1=f'"{",".join(OPCIONES_CONCENTRADO)}"', allow_blank=True)
    ws_bod.add_data_validation(dv_conc)
    dv_conc.add("C2:C100")

def cargar_datos():
    base = crear_estructura_base()
    if USAR_GSHEETS:
        datos = {}
        for hoja, df_def in base.items():
            try:
                df_leido = conn.read(worksheet=hoja, ttl=0)
                if df_leido is None or df_leido.empty:
                    conn.update(worksheet=hoja, data=df_def)
                    datos[hoja] = df_def
                else:
                    datos[hoja] = df_leido.dropna(how="all")
            except Exception:
                datos[hoja] = df_def
        return datos
    else:
        if not os.path.exists(ARCHIVO_EXCEL):
            with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
                for hoja, df in base.items():
                    df.to_excel(writer, sheet_name=hoja, index=False)
                aplicar_validaciones_excel(writer)
        xls = pd.ExcelFile(ARCHIVO_EXCEL)
        return {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}

def guardar_datos(dfs):
    if USAR_GSHEETS:
        for hoja, df in dfs.items():
            conn.update(worksheet=hoja, data=df)
    else:
        with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
            aplicar_validaciones_excel(writer)

db = cargar_datos()
hoy = datetime.now().date()

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("## 🐷 PorciGestión Pro")
    modo_txt = "☁️ Nube (Google Sheets)" if USAR_GSHEETS else "💻 Local (Excel)"
    st.caption(f"Conexión: {modo_txt}")
    st.caption(f"Perfil: **{'Administrador' if st.session_state.rol == 'admin' else 'Operario'}**")
    st.divider()

    menu_opciones = [
        "🚨 Tablero de Alertas",
        "🧬 Reproducción y Genética",
        "🌾 Bodega y Alimentación",
        "🐖 Lotes y Crecimiento",
        "💉 Sanidad y Mortalidad",
        "⚙️ Editar Tablas y Cerdas"
    ]
    if st.session_state.rol == "admin":
        menu_opciones.insert(5, "💰 Ventas (Pie / Canal)")
        menu_opciones.insert(6, "📊 Analítica e Inteligencia")

    seleccion = st.radio("Navegación:", menu_opciones)
    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# --- KPIs SUPERIORES GLOBALES ---
total_madres = len(db["Reproductores"][db["Reproductores"]["Tipo"] == "Hembra Madre"])
gestantes = len(db["Reproductores"][db["Reproductores"]["Estado"] == "Gestante"])
cerdos_engorde = int(db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["Cant_Actual"].sum()) if not db["Lotes_Engorde"].empty else 0
kilos_bodega = float(db["Bodega"]["Saldo_Kg"].sum()) if not db["Bodega"].empty else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("🐖 Cerdas Madres Activas", f"{total_madres} hembras", f"{gestantes} gestantes")
k2.metric("🥩 Cerdos en Crecimiento", f"{cerdos_engorde} animales", "En corrales activos")
k3.metric("🌾 Inventario Bodega", f"{kilos_bodega:,.0f} kg", f"{kilos_bodega/40:.1f} bultos (40kg)")
k4.metric("📅 Fecha Operativa", hoy.strftime("%d/%m/%Y"), "Sistema Sincronizado")
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 1. TABLERO DE ALERTAS INTELIGENTES
# ==========================================
if seleccion == "🚨 Tablero de Alertas":
    st.header("🚨 Centro de Alertas Operativas y Sanitarias")
    col1, col2, col3 = st.columns(3)
    
    rep = db["Reproduccion"]
    san = db["Sanidad_Eventos"]
    bod = db["Bodega"]
    lotes = db["Lotes_Engorde"]
    madres = db["Reproductores"]

    with col1:
        st.subheader("🤰 Maternidad y Descarte")
        alertas_rep = 0
        for _, row in rep[rep["Estado_Prenez"] == "Positiva"].iterrows():
            f_parto = pd.to_datetime(row["Fecha_Prob_Parto"]).date()
            dias_faltan = (f_parto - hoy).days
            if 0 <= dias_faltan <= 7:
                st.error(f"🔴 **PARTO INMINENTE:** Cerda **{row['ID_Cerda']}** pare en **{dias_faltan} días** ({f_parto}). Preparar jaula.")
                alertas_rep += 1
            elif dias_faltan < 0:
                st.warning(f"🟠 **PARTO ATRASADO:** Cerda **{row['ID_Cerda']}** tenía fecha {f_parto}. Registrar parto.")
                alertas_rep += 1

        for _, row in rep[rep["Estado_Prenez"] == "Pendiente Eco"].iterrows():
            f_eco = pd.to_datetime(row["Fecha_Eco_28d"]).date()
            if hoy >= f_eco:
                st.info(f"🔵 **ECOGRAFÍA / CELO:** Revisar preñez de Cerda **{row['ID_Cerda']}** (Servicio: {row['Fecha_Servicio']}).")
                alertas_rep += 1

        for _, m in madres.iterrows():
            if m["Abortos"] >= 2 or m["Repeticiones"] >= 2:
                st.warning(f"⚠️ **SUGERENCIA DESCARTE:** **{m['ID_Animal']}** acumula {m['Abortos']} abortos y {m['Repeticiones']} repeticiones.")
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
                st.warning(f"🩺 **ENFERMO ACTIVO:** **{s['ID_Destino']}** - {s['Detalle_Sintoma']} ({s['Medicamento']}).")
                alertas_san += 1
        if alertas_san == 0:
            st.success("✅ Todos los animales libres de retiro sanitario.")

    with col3:
        st.subheader("📦 Bodega y Transiciones")
        for _, b in bod.iterrows():
            if b["Saldo_Kg"] < 120:
                st.error(f"📉 **STOCK BAJO:** Quedan **{b['Saldo_Kg']} kg** ({b['Saldo_Kg']/40:.1f} bultos) de **{b['Tipo_Concentrado']}**.")
        
        for _, l in lotes[lotes["Estado"] == "Activo"].iterrows():
            f_ini = pd.to_datetime(l["Fecha_Inicio"]).date()
            dias_lote = (hoy - f_ini).days
            if l["Etapa"] == "Lactante" and dias_lote >= 28:
                st.warning(f"🍼 **DESTETE:** **{l['ID_Lote']}** cumplió {dias_lote} días. Pasar a Precebo.")
            elif l["Etapa"] == "Precebo" and dias_lote >= 42:
                st.info(f"🔄 **CAMBIO DE ETAPA:** **{l['ID_Lote']}** lleva {dias_lote} días. Pasar a Levante.")
            elif l["Etapa"] == "Levante" and dias_lote >= 45:
                st.info(f"🔄 **CAMBIO DE ETAPA:** **{l['ID_Lote']}** listo para pasar a Engorde.")

# ==========================================
# 2. REPRODUCCIÓN Y GENÉTICA
# ==========================================
elif seleccion == "🧬 Reproducción y Genética":
    st.header("🧬 Control de Cerdas Madres, Verraco e Inseminación")
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("1. Registrar Servicio (Monta / Inseminación)")
        hembras_disp = db["Reproductores"][db["Reproductores"]["Tipo"] == "Hembra Madre"]["ID_Animal"].tolist()
        id_cerda = st.selectbox("Seleccionar Cerda Madre", hembras_disp)
        f_serv = st.date_input("Fecha del Servicio", hoy)
        tipo_serv = st.selectbox("Tipo de Servicio", ["Inseminación Artificial (Pajilla)", "Monta Verraco Propio (V-01)", "Monta Verraco Alquilado"])
        macho_id = st.text_input("ID del Macho o Línea Genética del Semen", value="V-01" if "Propio" in tipo_serv else "Duroc-Genética")
        costo_serv = st.number_input("Costo del Servicio / Semen ($)", min_value=0, value=45000, step=5000)

        cerda_info = db["Reproductores"][db["Reproductores"]["ID_Animal"] == id_cerda].iloc[0]
        alerta_consang = False
        if macho_id == "V-01":
            verraco_info = db["Reproductores"][db["Reproductores"]["ID_Animal"] == "V-01"].iloc[0]
            if id_cerda == verraco_info["Madre_ID"] or cerda_info["Madre_ID"] == verraco_info["Madre_ID"]:
                st.error("🚨 ¡ALERTA DE CONSANGUINIDAD! V-01 es hijo o hermano de esta cerda. No se recomienda cruzar.")
                alerta_consang = True

        if st.button("Guardar Servicio", disabled=alerta_consang):
            nuevo_ev = {
                "ID_Evento": f"EV-{len(db['Reproduccion'])+1:03d}",
                "ID_Cerda": id_cerda,
                "Fecha_Servicio": str(f_serv),
                "Tipo_Servicio": tipo_serv,
                "Macho_O_Semen": macho_id,
                "Costo_Servicio": costo_serv,
                "Fecha_Eco_28d": str(f_serv + timedelta(days=28)),
                "Estado_Prenez": "Pendiente Eco",
                "Fecha_Prob_Parto": str(f_serv + timedelta(days=114)),
                "Fecha_Real_Parto": "", "Nacidos_Vivos": 0, "Mortinatos": 0, "Momias": 0,
                "Fecha_Prob_Destete": str(f_serv + timedelta(days=114+28)),
                "Destetados_Reales": 0, "Peso_Destete_Total_Kg": 0, "Costo_Por_Lechon": 0
            }
            db["Reproduccion"] = pd.concat([db["Reproduccion"], pd.DataFrame([nuevo_ev])], ignore_index=True)
            idx = db["Reproductores"][db["Reproductores"]["ID_Animal"] == id_cerda].index[0]
            db["Reproductores"].at[idx, "Costo_Acumulado_Ciclo"] += costo_serv
            db["Reproductores"].at[idx, "Estado"] = "Servida"
            guardar_datos(db)
            st.success(f"Servicio registrado. Parto probable: {f_serv + timedelta(days=114)}")
            st.rerun()

    with c2:
        st.subheader("2. Actualizar Preñez, Parto o Destete")
        if not db["Reproduccion"].empty:
            ev_sel = st.selectbox("Seleccionar Evento Reproductivo", db["Reproduccion"]["ID_Evento"] + " - " + db["Reproduccion"]["ID_Cerda"])
            id_ev = ev_sel.split(" - ")[0]
            idx_ev = db["Reproduccion"][db["Reproduccion"]["ID_Evento"] == id_ev].index[0]
            cerda_ev = db["Reproduccion"].at[idx_ev, "ID_Cerda"]
            idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == cerda_ev].index[0]

            accion = st.radio("Acción a registrar:", ["Confirmar Preñez (Positiva)", "Repitió Celo / Abortó (Pérdida)", "Registrar Parto", "Registrar Destete y Crear Lote"])
            
            if accion == "Confirmar Preñez (Positiva)":
                if st.button("Confirmar Preñez"):
                    db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = "Positiva"
                    db["Reproductores"].at[idx_m, "Estado"] = "Gestante"
                    guardar_datos(db)
                    st.success("Preñez confirmada.")
                    st.rerun()

            elif accion == "Repitió Celo / Abortó (Pérdida)":
                motivo = st.selectbox("Motivo", ["Repitió Celo", "Aborto"])
                if st.button("Registrar Falla Reproductiva"):
                    db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = motivo
                    db["Reproductores"].at[idx_m, "Estado"] = "Vacía"
                    if motivo == "Repitió Celo":
                        db["Reproductores"].at[idx_m, "Repeticiones"] += 1
                    else:
                        db["Reproductores"].at[idx_m, "Abortos"] += 1
                    guardar_datos(db)
                    st.warning("Falla registrada. El costo acumulado se cargará a su próxima camada exitosa.")
                    st.rerun()

            elif accion == "Registrar Parto":
                f_parto_r = st.date_input("Fecha Real de Parto", hoy)
                vivos = st.number_input("Nacidos Vivos", min_value=0, value=10)
                muertos = st.number_input("Mortinatos (Nacidos muertos)", min_value=0, value=0)
                momias = st.number_input("Momias", min_value=0, value=0)
                if st.button("Guardar Parto"):
                    db["Reproduccion"].at[idx_ev, "Fecha_Real_Parto"] = str(f_parto_r)
                    db["Reproduccion"].at[idx_ev, "Nacidos_Vivos"] = vivos
                    db["Reproduccion"].at[idx_ev, "Mortinatos"] = muertos
                    db["Reproduccion"].at[idx_ev, "Momias"] = momias
                    db["Reproduccion"].at[idx_ev, "Fecha_Prob_Destete"] = str(f_parto_r + timedelta(days=28))
                    db["Reproductores"].at[idx_m, "Estado"] = "Lactante"
                    db["Reproductores"].at[idx_m, "Partos"] += 1
                    db["Reproductores"].at[idx_m, "Racion_Diaria_Kg"] = 5.5
                    guardar_datos(db)
                    st.success("Parto registrado. Ración ajustada a 5.5 kg/día.")
                    st.rerun()

            elif accion == "Registrar Destete y Crear Lote":
                destetados = st.number_input("Lechones Destetados Reales", min_value=1, value=10)
                peso_tot = st.number_input("Peso Total de la Camada al Destete (kg)", min_value=10.0, value=75.0)
                corral_dest = st.text_input("Corral de Precebo destino", value="Corral-Precebo-1")
                if st.button("Finalizar Destete y Pasar a Engorde"):
                    costo_ciclo_madre = db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"]
                    costo_por_lechon = costo_ciclo_madre / destetados if destetados > 0 else 0
                    
                    db["Reproduccion"].at[idx_ev, "Destetados_Reales"] = destetados
                    db["Reproduccion"].at[idx_ev, "Peso_Destete_Total_Kg"] = peso_tot
                    db["Reproduccion"].at[idx_ev, "Costo_Por_Lechon"] = costo_por_lechon
                    db["Reproduccion"].at[idx_ev, "Estado_Prenez"] = "Destetada"

                    nuevo_lote = {
                        "ID_Lote": f"LOTE-{len(db['Lotes_Engorde'])+1:03d}",
                        "Corral": corral_dest, "Etapa": "Precebo", "Fecha_Inicio": str(hoy),
                        "Cant_Inicial": destetados, "Cant_Actual": destetados,
                        "Peso_Prom_Inicial_Kg": round(peso_tot / destetados, 2),
                        "Peso_Prom_Actual_Kg": round(peso_tot / destetados, 2),
                        "Kg_Alimento_Consumido": 0.0,
                        "Costo_Inicial_Lechones": costo_ciclo_madre,
                        "Costo_Alimento": 0.0, "Costo_Sanidad": 0.0, "Estado": "Activo"
                    }
                    db["Lotes_Engorde"] = pd.concat([db["Lotes_Engorde"], pd.DataFrame([nuevo_lote])], ignore_index=True)

                    db["Reproductores"].at[idx_m, "Estado"] = "Vacía"
                    db["Reproductores"].at[idx_m, "Racion_Diaria_Kg"] = 2.2
                    db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] = 0.0
                    guardar_datos(db)
                    st.success(f"Destete exitoso. Cada lechón arranca costando ${costo_por_lechon:,.0f}.")
                    st.rerun()

    st.subheader("Inventario Actual de Reproductores")
    st.dataframe(db["Reproductores"], use_container_width=True)

# ==========================================
# 3. BODEGA Y ALIMENTACIÓN
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
            opciones_dest = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["ID_Lote"].tolist()
        else:
            opciones_dest = db["Reproductores"]["ID_Animal"].tolist()

        if opciones_dest:
            id_dest = st.selectbox("Seleccionar Destino", opciones_dest)
            conc_disponibles = db["Bodega"][db["Bodega"]["Saldo_Kg"] > 0]["Tipo_Concentrado"].unique().tolist()
            if conc_disponibles:
                conc_usar = st.selectbox("Concentrado sacado de bodega", conc_disponibles)
                unidad = st.radio("Unidad de medida", ["Kilos (kg)", "Bultos (40 kg)"])
                cant_ing = st.number_input("Cantidad entregada", min_value=0.5, value=40.0 if unidad=="Kilos (kg)" else 1.0, step=0.5)
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
    st.dataframe(db["Bodega"], use_container_width=True)

# ==========================================
# 4. LOTES DE ENGORDE Y CRECIMIENTO
# ==========================================
elif seleccion == "🐖 Lotes y Crecimiento":
    st.header("🐖 Control de Lotes, Pesajes y Conversión Alimenticia")
    
    with st.expander("➕ Crear Lote Manual (Para cerdos o camadas que ya tienes actualmente en la granja)"):
        c_l1, c_l2, c_l3, c_l4 = st.columns(4)
        corral_m = c_l1.text_input("Corral", "Corral-1")
        etapa_m = c_l2.selectbox("Etapa Actual", OPCIONES_ETAPA_LOTE)
        cant_m = c_l3.number_input("Cantidad de Cerdos", min_value=1, value=12)
        peso_m = c_l4.number_input("Peso Promedio Hoy (kg)", min_value=1.0, value=35.0)
        costo_est_m = st.number_input("Costo estimado invertido hasta hoy en este grupo ($)", min_value=0, value=2500000, step=100000)
        if st.button("Crear Lote Existente"):
            nl = {
                "ID_Lote": f"LOTE-{len(db['Lotes_Engorde'])+1:03d}", "Corral": corral_m,
                "Etapa": etapa_m, "Fecha_Inicio": str(hoy), "Cant_Inicial": cant_m, "Cant_Actual": cant_m,
                "Peso_Prom_Inicial_Kg": peso_m, "Peso_Prom_Actual_Kg": peso_m,
                "Kg_Alimento_Consumido": 0.0, "Costo_Inicial_Lechones": costo_est_m,
                "Costo_Alimento": 0.0, "Costo_Sanidad": 0.0, "Estado": "Activo"
            }
            db["Lotes_Engorde"] = pd.concat([db["Lotes_Engorde"], pd.DataFrame([nl])], ignore_index=True)
            guardar_datos(db)
            st.success("Lote agregado al sistema.")
            st.rerun()

    lotes_act = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]
    if not lotes_act.empty:
        st.subheader("Actualizar Pesaje o Etapa de un Lote")
        col_p1, col_p2, col_p3 = st.columns(3)
        lote_sel = col_p1.selectbox("Lote a actualizar", lotes_act["ID_Lote"].tolist())
        nuevo_peso = col_p2.number_input("Nuevo Peso Promedio (kg)", min_value=1.0, value=50.0)
        nueva_etapa = col_p3.selectbox("Etapa", OPCIONES_ETAPA_LOTE)
        if st.button("Actualizar Lote"):
            idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == lote_sel].index[0]
            db["Lotes_Engorde"].at[idx_l, "Peso_Prom_Actual_Kg"] = nuevo_peso
            db["Lotes_Engorde"].at[idx_l, "Etapa"] = nueva_etapa
            guardar_datos(db)
            st.success("Pesaje y etapa actualizados.")
            st.rerun()

        st.subheader("📊 Indicadores en Tiempo Real por Lote (Absorbiendo Mortalidad)")
        tabla_kpi = []
        for _, r in lotes_act.iterrows():
            costo_total_lote = r["Costo_Inicial_Lechones"] + r["Costo_Alimento"] + r["Costo_Sanidad"]
            costo_por_cerdo_vivo = costo_total_lote / r["Cant_Actual"] if r["Cant_Actual"] > 0 else 0
            kilos_totales_vivos = r["Cant_Actual"] * r["Peso_Prom_Actual_Kg"]
            costo_por_kilo_pie = costo_total_lote / kilos_totales_vivos if kilos_totales_vivos > 0 else 0
            
            ganancia_peso_lote = (r["Peso_Prom_Actual_Kg"] - r["Peso_Prom_Inicial_Kg"]) * r["Cant_Actual"]
            conv_alim = r["Kg_Alimento_Consumido"] / ganancia_peso_lote if ganancia_peso_lote > 0 else 0

            fila = {
                "Lote": r["ID_Lote"], "Corral": r["Corral"], "Etapa": r["Etapa"],
                "Vivos / Inicial": f"{r['Cant_Actual']} / {r['Cant_Inicial']}",
                "Peso Prom (kg)": r["Peso_Prom_Actual_Kg"],
                "Conv. Alimenticia": round(conv_alim, 2)
            }
            if st.session_state.rol == "admin":
                fila["Costo Total Lote ($)"] = f"${costo_total_lote:,.0f}"
                fila["Costo x Cerdo Vivo ($)"] = f"${costo_por_cerdo_vivo:,.0f}"
                fila["Punto Equilibrio Pie ($/kg)"] = f"${costo_por_kilo_pie:,.0f}"
            tabla_kpi.append(fila)
        st.dataframe(pd.DataFrame(tabla_kpi), use_container_width=True)

# ==========================================
# 5. SANIDAD Y MORTALIDAD
# ==========================================
elif seleccion == "💉 Sanidad y Mortalidad":
    st.header("💉 Bitácora Sanitaria, Vacunas y Bajas (Mortalidad)")
    s1, s2 = st.columns(2)
    
    destinos_todos = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]["ID_Lote"].tolist() + db["Reproductores"]["ID_Animal"].tolist()
    
    with s1:
        st.subheader("Registrar Evento Sanitario o Mortalidad")
        dest_s = st.selectbox("Animal o Lote afectado", destinos_todos)
        cat_s = st.selectbox("Tipo de Registro", ["Vacunación Preventiva", "Tratamiento / Enfermedad", "Mortalidad (Baja)", "Observación General"])
        detalle_s = st.text_input("Síntoma, Vacuna o Causa de Muerte", "Ej. Vacuna Circovirus / Diarrea / Aplastamiento")
        med_s = st.text_input("Medicamento aplicado (si aplica)", "Ninguno")
        costo_s = st.number_input("Costo del tratamiento/vacuna ($)", min_value=0, value=0, step=1000)
        retiro_s = st.number_input("Días de retiro antes de sacrificio", min_value=0, value=0)
        muertos_s = st.number_input("Cantidad de animales muertos (solo si es Mortalidad)", min_value=0, value=1 if cat_s=="Mortalidad (Baja)" else 0)

        if st.button("Guardar Registro Sanitario"):
            f_fin_ret = str(hoy + timedelta(days=retiro_s)) if retiro_s > 0 else ""
            estado_c = "En Tratamiento" if cat_s == "Tratamiento / Enfermedad" else "Cerrado"
            
            nuevo_s = {
                "Fecha": str(hoy), "ID_Destino": dest_s, "Categoria": cat_s,
                "Detalle_Sintoma": detalle_s, "Medicamento": med_s, "Costo": costo_s,
                "Dias_Retiro": retiro_s, "Fecha_Fin_Retiro": f_fin_ret, "Estado_Caso": estado_c
            }
            db["Sanidad_Eventos"] = pd.concat([db["Sanidad_Eventos"], pd.DataFrame([nuevo_s])], ignore_index=True)

            if dest_s in db["Lotes_Engorde"]["ID_Lote"].values:
                idx_l = db["Lotes_Engorde"][db["Lotes_Engorde"]["ID_Lote"] == dest_s].index[0]
                db["Lotes_Engorde"].at[idx_l, "Costo_Sanidad"] += costo_s
                if cat_s == "Mortalidad (Baja)" and muertos_s > 0:
                    db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] = max(0, db["Lotes_Engorde"].at[idx_l, "Cant_Actual"] - muertos_s)
                    st.warning(f"⚠️ Se restaron {muertos_s} cerdos de {dest_s}. El costo se redistribuyó entre los sobrevivientes.")
            elif dest_s in db["Reproductores"]["ID_Animal"].values:
                idx_m = db["Reproductores"][db["Reproductores"]["ID_Animal"] == dest_s].index[0]
                db["Reproductores"].at[idx_m, "Costo_Acumulado_Ciclo"] += costo_s

            guardar_datos(db)
            st.success("Evento sanitario registrado.")
            st.rerun()

    with s2:
        st.subheader("Historial Sanitario y Observaciones")
        st.dataframe(db["Sanidad_Eventos"], use_container_width=True)

# ==========================================
# 6. VENTAS (PIE / CANAL) - SOLO ADMIN
# ==========================================
elif seleccion == "💰 Ventas (Pie / Canal)":
    st.header("💰 Calculadora de Venta (En Pie vs. Beneficiado) y Balance")
    lotes_v = db["Lotes_Engorde"][db["Lotes_Engorde"]["Estado"] == "Activo"]
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

            cant_v = st.number_input("Cantidad de cerdos a vender", min_value=1, max_value=int(row_lv["Cant_Actual"]), value=int(row_lv["Cant_Actual"]))
            modalidad = st.radio("Modalidad de Venta", ["En Pie (Cerdo entero en báscula)", "Beneficiado (Carne en canal por kilos)"])
            peso_pie_tot = st.number_input("Peso Total en Pie (kg)", min_value=10.0, value=float(cant_v * row_lv["Peso_Prom_Actual_Kg"]))
            
            costo_tot_lote = row_lv["Costo_Inicial_Lechones"] + row_lv["Costo_Alimento"] + row_lv["Costo_Sanidad"]
            costo_proporcional = (costo_tot_lote / row_lv["Cant_Actual"]) * cant_v

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
            st.dataframe(db["Ventas"], use_container_width=True)
            if not db["Ventas"].empty:
                st.metric("Ingresos Totales Acumulados", f"${db['Ventas']['Ingreso_Bruto'].sum():,.0f}")
                st.metric("Utilidad Neta Total Acumulada", f"${db['Ventas']['Utilidad_Neta'].sum():,.0f}")
    else:
        st.info("Crea o desteta un lote primero para simular ventas.")

# ==========================================
# 7. ANALÍTICA E INTELIGENCIA DE DATOS - MOTOR DINÁMICO
# ==========================================
elif seleccion == "📊 Analítica e Inteligencia":
    st.header("📊 Motor de Analítica Zootécnica y Recomendaciones Automáticas")
    st.caption("El sistema analiza correlaciones reales entre tus cerdas, genética, corrales, mortalidad y ventas para detectar oportunidades de ahorro.")
    
    hallazgos_encontrados = 0
    lotes_df = db["Lotes_Engorde"].copy()
    rep_df = db["Reproduccion"].copy()
    madres_df = db["Reproductores"].copy()
    san_df = db["Sanidad_Eventos"].copy()
    ventas_df = db["Ventas"].copy()

    # --- 1. ANÁLISIS REPRODUCTIVO Y DESCARTE DE CERDAS ---
    for _, m in madres_df[madres_df["Tipo"] == "Hembra Madre"].iterrows():
        fallas = m["Abortos"] + m["Repeticiones"]
        if fallas >= 2:
            st.error(f"🧬 **Eficiencia Reproductiva ({m['ID_Animal']}):** Acumula **{fallas} fallas** ({m['Repeticiones']} repeticiones, {m['Abortos']} abortos) y **${m['Costo_Acumulado_Ciclo']:,.0f}** en costos acumulados del ciclo actual. **Recomendación:** Evaluar descarte y reemplazo genético.")
            hallazgos_encontrados += 1

    # --- 2. ANÁLISIS DE GENÉTICA / PADRES ---
    if not rep_df.empty and rep_df["Destetados_Reales"].sum() > 0:
        gen_perf = rep_df[rep_df["Destetados_Reales"] > 0].groupby("Macho_O_Semen").agg(
            Promedio_Destetados=("Destetados_Reales", "mean"),
            Peso_Camada_Prom=("Peso_Destete_Total_Kg", "mean"),
            Camadas=("ID_Evento", "count")
        ).reset_index()
        if len(gen_perf) >= 2:
            mejor_gen = gen_perf.sort_values(by="Peso_Camada_Prom", ascending=False).iloc[0]
            st.success(f"🏆 **Línea Genética Superior:** Las camadas de **{mejor_gen['Macho_O_Semen']}** lideran el rendimiento con un peso promedio de **{mejor_gen['Peso_Camada_Prom']:.1f} kg** al destete ({mejor_gen['Promedio_Destetados']:.1f} lechones/camada).")
            hallazgos_encontrados += 1

    # --- 3. ANÁLISIS DINÁMICO DE DENSIDAD Y CONVERSIÓN ALIMENTICIA ---
    if not lotes_df.empty:
        lotes_df["Ganancia_Total"] = (lotes_df["Peso_Prom_Actual_Kg"] - lotes_df["Peso_Prom_Inicial_Kg"]) * lotes_df["Cant_Actual"]
        lotes_df["CA"] = lotes_df.apply(lambda r: r["Kg_Alimento_Consumido"] / r["Ganancia_Total"] if r["Ganancia_Total"] > 0 else 0, axis=1)
        lotes_validos = lotes_df[lotes_df["CA"] > 0]
        
        if len(lotes_validos) >= 2:
            mediana_tamano = lotes_validos["Cant_Actual"].median()
            ca_prom_granja = lotes_validos["CA"].mean()
            
            # Detectar lotes con conversión 15% peor que el promedio de la granja
            ineficientes = lotes_validos[lotes_validos["CA"] > ca_prom_granja * 1.15]
            for _, l_malo in ineficientes.iterrows():
                extra_txt = " (posible competencia por comedero al superar el tamaño habitual de grupo)" if l_malo["Cant_Actual"] > mediana_tamano else " (revisar desperdicio en tolva o estado sanitario)"
                st.warning(f"📉 **Alerta de Conversión ({l_malo['ID_Lote']} - {l_malo['Corral']}):** Tiene una conversión de **{l_malo['CA']:.2f}**, un **{((l_malo['CA']/ca_prom_granja)-1)*100:.0f}% peor** que el promedio de tu granja ({ca_prom_granja:.2f}){extra_txt}.")
                hallazgos_encontrados += 1

    # --- 4. ANÁLISIS DE PATRONES DE MORTALIDAD Y SANIDAD ---
    if not san_df.empty:
        bajas = san_df[san_df["Categoria"] == "Mortalidad (Baja)"]
        if len(bajas) >= 2:
            causa_top = bajas["Detalle_Sintoma"].value_counts().idxmax()
            cant_causa = bajas["Detalle_Sintoma"].value_counts().max()
            st.warning(f"🩺 **Patrón Sanitario Detectado:** La causa principal de mortalidad registrada es **'{causa_top}'** ({cant_causa} eventos). Revisa protocolos preventivos en esa etapa.")
            hallazgos_encontrados += 1

    # --- 5. ANÁLISIS FINANCIERO (PIE VS. BENEFICIADO Y ESTACIONALIDAD) ---
    if not ventas_df.empty:
        ventas_df["Mes"] = pd.to_datetime(ventas_df["Fecha"]).dt.strftime("%Y-%m")
        mejor_mod = ventas_df.groupby("Modalidad")["Utilidad_Neta"].mean()
        if len(mejor_mod) > 1:
            top_mod = mejor_mod.idxmax()
            diff = mejor_mod.max() - mejor_mod.min()
            st.success(f"💰 **Estrategia Comercial:** Vender bajo la modalidad **{top_mod}** te está dejando en promedio **${diff:,.0f} más de utilidad neta** por operación frente a la otra modalidad.")
            hallazgos_encontrados += 1

    if hallazgos_encontrados == 0:
        st.info("🌱 **Motor de Inteligencia en Aprendizaje:** A medida que registres pesajes, servicios, eventos sanitarios y ventas, esta sección generará automáticamente diagnósticos personalizados sobre tus cerdas, corrales y rentabilidad.")

    st.divider()

    # --- GRÁFICOS INTERACTIVOS ---
    g1, g2 = st.columns(2)
    with g1:
        st.subheader("1. Distribución de Costos de Alimento ($)")
        if not db["Consumo_Alimento"].empty:
            fig_cons = px.pie(
                db["Consumo_Alimento"], values="Costo_Total", names="Tipo_Concentrado",
                hole=0.45, color_discrete_sequence=px.colors.sequential.Greens_r
            )
            st.plotly_chart(fig_cons, use_container_width=True)
        else:
            st.info("Registra consumos de alimento para visualizar la distribución de costos.")

    with g2:
        st.subheader("2. Rendimiento por Cerda Madre (Nacidos vs Destetados)")
        if not rep_df.empty and rep_df["Nacidos_Vivos"].sum() > 0:
            fig_rep = px.bar(
                rep_df, x="ID_Cerda", y=["Nacidos_Vivos", "Destetados_Reales"],
                barmode="group", color_discrete_sequence=["#2E86C1", "#1E8449"]
            )
            st.plotly_chart(fig_rep, use_container_width=True)
        else:
            st.info("Cuando registres partos y destetes verás aquí el ranking gráfico de tus cerdas.")

    g3, g4 = st.columns(2)
    with g3:
        st.subheader("3. Tamaño de Grupo vs. Conversión Alimenticia")
        if not lotes_df.empty and lotes_df["CA"].sum() > 0:
            fig_ca = px.scatter(
                lotes_df[lotes_df["CA"] > 0], x="Cant_Actual", y="CA", size="Peso_Prom_Actual_Kg",
                color="Etapa", hover_name="ID_Lote",
                labels={"Cant_Actual": "Cerdos en Corral", "CA": "Conversión Alimenticia (Menor = Mejor)"}
            )
            st.plotly_chart(fig_ca, use_container_width=True)
        else:
            st.info("Actualiza el peso de tus lotes después de alimentarlos para ver la curva de eficiencia.")

    with g4:
        st.subheader("4. Evolución Mensual de Ingresos vs. Utilidad Neta")
        if not ventas_df.empty:
            resumen_mes = ventas_df.groupby("Mes")[["Ingreso_Bruto", "Utilidad_Neta"]].sum().reset_index()
            fig_fin = px.bar(
                resumen_mes, x="Mes", y=["Ingreso_Bruto", "Utilidad_Neta"],
                barmode="group", color_discrete_sequence=["#5D6D7E", "#1E8449"]
            )
            st.plotly_chart(fig_fin, use_container_width=True)
        else:
            st.info("Registra tus primeras ventas para analizar los mejores meses del año.")

# ==========================================
# 8. EDITOR DIRECTO CON LISTAS DESPLEGABLES
# ==========================================
elif seleccion == "⚙️ Editar Tablas y Cerdas":
    st.header("⚙️ Editor Directo de Cerdas y Lotes (Con Listas Desplegables)")
    st.write("Haz doble clic sobre cualquier celda (como **Estado**, **Tipo** o **Etapa**) para desplegar las opciones válidas y guarda los cambios.")

    st.subheader("1. Editar Cerdas Madres y Verracos")
    df_repro_editado = st.data_editor(
        db["Reproductores"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=OPCIONES_TIPO_REPRO, required=True),
            "Estado": st.column_config.SelectboxColumn("Estado Actual", options=OPCIONES_ESTADO_REPRO, required=True),
            "Racion_Diaria_Kg": st.column_config.NumberColumn("Ración Diaria (kg)", min_value=0.0, max_value=12.0, step=0.1),
        }
    )
    if st.button("💾 Guardar Cambios en Reproductores"):
        db["Reproductores"] = df_repro_editado
        guardar_datos(db)
        st.success("¡Cambios guardados con éxito!")
        st.rerun()

    st.divider()
    st.subheader("2. Editar Lotes de Cerdos")
    df_lotes_editado = st.data_editor(
        db["Lotes_Engorde"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Etapa": st.column_config.SelectboxColumn("Etapa", options=OPCIONES_ETAPA_LOTE, required=True),
            "Estado": st.column_config.SelectboxColumn("Estado", options=["Activo", "Vendido"], required=True),
        }
    )
    if st.button("💾 Guardar Cambios en Lotes"):
        db["Lotes_Engorde"] = df_lotes_editado
        guardar_datos(db)
        st.success("¡Lotes actualizados correctamente!")
        st.rerun()