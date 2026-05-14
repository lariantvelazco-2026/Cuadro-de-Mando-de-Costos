import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from fpdf import FPDF
from datetime import datetime
import io
import warnings
warnings.filterwarnings('ignore')

# ============================================
# CONFIGURACIÓN DE PÁGINA
# ============================================
st.set_page_config(
    page_title="Cuadro de Mando Control de Costos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# ESTILOS CSS
# ============================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    .main-header {
        font-size: 1.8rem;
        font-weight: bold;
        background: linear-gradient(135deg, #8B0000 0%, #DC143C 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(139, 0, 0, 0.3);
    }
    
    .sub-header {
        color: #8B0000;
        font-size: 1.4rem;
        font-weight: bold;
        border-bottom: 3px solid #8B0000;
        padding-bottom: 0.5rem;
        margin: 1.5rem 0 1rem 0;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1.2rem;
        border-radius: 12px;
        border-left: 5px solid #8B0000;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    
    .alert-danger {
        background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
        border-left: 5px solid #dc3545;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .alert-warning {
        background: linear-gradient(135deg, #fff3cd 0%, #ffeeba 100%);
        border-left: 5px solid #ffc107;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .alert-success {
        background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
        border-left: 5px solid #28a745;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .alert-info {
        background: linear-gradient(135deg, #cce5ff 0%, #b8daff 100%);
        border-left: 5px solid #007bff;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .upload-area {
        background: #f8f9fa;
        border: 3px dashed #8B0000;
        border-radius: 15px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
    }
    
    .footer {
        text-align: center;
        color: #6c757d;
        padding: 2rem;
        font-size: 0.85rem;
        border-top: 2px solid #dee2e6;
        margin-top: 3rem;
        background: #f8f9fa;
    }
    
    .sidebar-logo {
        text-align: center;
        padding: 1rem;
        background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
        border-radius: 10px;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# FUNCIONES DE PROCESAMIENTO
# ============================================

@st.cache_data
def cargar_excel(uploaded_file):
    """Carga el archivo Excel"""
    try:
        xl = pd.ExcelFile(uploaded_file)
        hojas = xl.sheet_names
        datos = {'hojas': hojas, 'dataframes': {}}
        
        for hoja in hojas:
            df = pd.read_excel(uploaded_file, sheet_name=hoja, header=0)
            df.columns = df.columns.astype(str).str.strip()
            datos['dataframes'][hoja] = df
        
        return datos, None
    except Exception as e:
        return None, str(e)

def procesar_matriz(df):
    """Procesa la hoja MATRIZ"""
    resultados = {}
    
    # Buscar columna TIPO
    col_tipo = None
    for col in df.columns:
        if 'TIPO' in col.upper():
            col_tipo = col
            break
    
    # Separar INGRESOS y COSTOS
    if col_tipo:
        df_ingresos = df[df[col_tipo].astype(str).str.upper().str.contains('INGRESO', na=False)].copy()
        df_costos = df[df[col_tipo].astype(str).str.upper().str.contains('COSTO', na=False)].copy()
    else:
        df_ingresos = df.copy()
        df_costos = pd.DataFrame()
    
    resultados['df_ingresos'] = df_ingresos
    resultados['df_costos'] = df_costos
    resultados['total_registros'] = len(df)
    resultados['registros_ingresos'] = len(df_ingresos)
    resultados['registros_costos'] = len(df_costos)
    
    # Buscar columna STATUS
    col_status = None
    for col in df.columns:
        if 'STATUS' in col.upper():
            col_status = col
            break
    
    if col_status:
        resultados['status_counts'] = df[col_status].value_counts().to_dict()
        resultados['en_proceso'] = df_ingresos[df_ingresos[col_status].astype(str).str.contains('Proceso', case=False, na=False)].copy()
        resultados['no_iniciadas'] = df_ingresos[df_ingresos[col_status].astype(str).str.contains('No Iniciada', case=False, na=False)].copy()
        resultados['completadas'] = df_ingresos[df_ingresos[col_status].astype(str).str.contains('Completada', case=False, na=False)].copy()
    
    # Buscar columna AWP/AREA
    col_awp = None
    for col in df.columns:
        if 'AREA' in col.upper() and '(*)' in col:
            col_awp = col
            break
    
    if col_awp:
        resultados['col_awp'] = col_awp
        resultados['awp_list'] = df_ingresos[col_awp].dropna().unique().tolist()
        resultados['resumen_awp'] = calcular_resumen_awp(df_ingresos, col_awp)
    
    # Buscar columna DISCIPLINA
    col_disciplina = None
    for col in df.columns:
        if 'DISCIPLINA' in col.upper():
            col_disciplina = col
            break
    
    if col_disciplina:
        resultados['disciplinas'] = df_ingresos[col_disciplina].dropna().unique().tolist()
        resultados['resumen_disciplina'] = calcular_resumen_disciplina(df_ingresos, col_disciplina)
    
    # Buscar columna FAMILIAS
    col_familias = None
    for col in df.columns:
        if 'FAMILIA' in col.upper():
            col_familias = col
            break
    
    if col_familias:
        resultados['familias'] = df_ingresos[col_familias].dropna().unique().tolist()
    
    # Calcular totales
    resultados['totales'] = calcular_totales(df_ingresos)
    
    # Calcular variaciones por status
    if col_status:
        resultados['variaciones_status'] = calcular_variaciones_status(df_ingresos, col_status)
    
    return resultados

def calcular_resumen_awp(df, col_awp):
    """Resumen por AWP"""
    resumen = []
    
    # Buscar columnas de totales
    col_ingreso = buscar_columna(df, ['TOTAL ACUMULADO INGRESOS(+)', 'TOTAL ACUMULADO INGRESOS', 'INGRESOS'])
    col_costo = buscar_columna(df, ['TOTAL ACUMULADO COSTOS (-)', 'TOTAL ACUMULADO COSTOS', 'COSTOS'])
    col_resultado = buscar_columna(df, ['RESULTADO'])
    
    for awp in df[col_awp].dropna().unique():
        df_awp = df[df[col_awp] == awp]
        
        ingreso = pd.to_numeric(df_awp[col_ingreso], errors='coerce').sum() if col_ingreso else 0
        costo = pd.to_numeric(df_awp[col_costo], errors='coerce').sum() if col_costo else 0
        resultado = pd.to_numeric(df_awp[col_resultado], errors='coerce').sum() if col_resultado else ingreso - costo
        
        resumen.append({
            'AWP': str(awp)[:50],
            'INGRESO_TOTAL': ingreso,
            'COSTO_TOTAL': costo,
            'RESULTADO': resultado,
            'REGISTROS': len(df_awp)
        })
    
    return pd.DataFrame(resumen)

def calcular_resumen_disciplina(df, col_disc):
    """Resumen por Disciplina"""
    resumen = []
    
    col_ingreso = buscar_columna(df, ['TOTAL ACUMULADO INGRESOS(+)', 'TOTAL ACUMULADO INGRESOS'])
    col_costo = buscar_columna(df, ['TOTAL ACUMULADO COSTOS (-)', 'TOTAL ACUMULADO COSTOS'])
    col_resultado = buscar_columna(df, ['RESULTADO'])
    col_cv_pos = buscar_columna(df, ['SV POSITIVAS', 'CV POSITIVAS'])
    col_cv_neg = buscar_columna(df, ['SV NEGATIVAS', 'CV NEGATIVAS'])
    
    for disc in df[col_disc].dropna().unique():
        df_disc = df[df[col_disc] == disc]
        
        resumen.append({
            'DISCIPLINA': str(disc),
            'INGRESO_TOTAL': pd.to_numeric(df_disc[col_ingreso], errors='coerce').sum() if col_ingreso else 0,
            'COSTO_TOTAL': pd.to_numeric(df_disc[col_costo], errors='coerce').sum() if col_costo else 0,
            'RESULTADO': pd.to_numeric(df_disc[col_resultado], errors='coerce').sum() if col_resultado else 0,
            'CV_POSITIVAS': pd.to_numeric(df_disc[col_cv_pos], errors='coerce').sum() if col_cv_pos else 0,
            'CV_NEGATIVAS': pd.to_numeric(df_disc[col_cv_neg], errors='coerce').sum() if col_cv_neg else 0,
            'REGISTROS': len(df_disc)
        })
    
    return pd.DataFrame(resumen)

def calcular_totales(df):
    """Calcula totales generales"""
    totales = {}
    
    cols_buscar = {
        'PO': ['P.O.'],
        'POM': ['P.O.M.'],
        'INGRESOS_ACUM': ['TOTAL ACUMULADO INGRESOS(+)', 'TOTAL ACUMULADO INGRESOS'],
        'COSTOS_ACUM': ['TOTAL ACUMULADO COSTOS (-)', 'TOTAL ACUMULADO COSTOS'],
        'RESULTADO': ['RESULTADO'],
        'AVANCE_TRABAJOS_EP': ['AVANCE TRABAJOS EP INGRESOS', 'AVANCE TRABAJOS EP'],
        'DESFASADO': ['DESFASADO OONN + OOEE', 'DESFASADO'],
    }
    
    for key, opciones in cols_buscar.items():
        col = buscar_columna(df, opciones)
        if col:
            totales[key] = pd.to_numeric(df[col], errors='coerce').sum()
        else:
            totales[key] = 0
    
    # Margen
    if totales.get('INGRESOS_ACUM', 0) != 0:
        totales['MARGEN_PCT'] = (totales.get('RESULTADO', 0) / totales.get('INGRESOS_ACUM', 1)) * 100
    else:
        totales['MARGEN_PCT'] = 0
    
    return totales

def calcular_variaciones_status(df, col_status):
    """Variaciones por status"""
    resultados = []
    
    col_pom = buscar_columna(df, ['P.O.M.'])
    col_avance = buscar_columna(df, ['AVANCE TRABAJOS EP INGRESOS', 'AVANCE TRABAJOS EP'])
    col_costos = buscar_columna(df, ['TOTAL ACUMULADO COSTOS (-)', 'TOTAL ACUMULADO COSTOS'])
    col_cv_pos = buscar_columna(df, ['SV POSITIVAS', 'CV POSITIVAS'])
    col_cv_neg = buscar_columna(df, ['SV NEGATIVAS', 'CV NEGATIVAS'])
    
    for status in df[col_status].dropna().unique():
        df_st = df[df[col_status] == status]
        
        pom = pd.to_numeric(df_st[col_pom], errors='coerce').sum() if col_pom else 0
        avance = pd.to_numeric(df_st[col_avance], errors='coerce').sum() if col_avance else 0
        costos = pd.to_numeric(df_st[col_costos], errors='coerce').sum() if col_costos else 0
        cv_pos = pd.to_numeric(df_st[col_cv_pos], errors='coerce').sum() if col_cv_pos else 0
        cv_neg = pd.to_numeric(df_st[col_cv_neg], errors='coerce').sum() if col_cv_neg else 0
        
        resultados.append({
            'STATUS': str(status),
            'P.O.M': pom,
            'AVANCE_TRABAJOS_EP': avance,
            'COSTOS_TOTALES': costos,
            'CV_TOTAL': avance - costos,
            'CV_POSITIVAS': cv_pos,
            'CV_NEGATIVAS': cv_neg,
            'REGISTROS': len(df_st)
        })
    
    return pd.DataFrame(resultados)

def buscar_columna(df, opciones):
    """Busca una columna por lista de opciones"""
    for opcion in opciones:
        for col in df.columns:
            if opcion.upper() in col.upper():
                return col
    return None

def formato_moneda(valor, decimales=0):
    """Formatea valores monetarios"""
    if pd.isna(valor) or valor == 0:
        return "$ 0"
    
    if abs(valor) >= 1e9:
        return f"$ {valor/1e9:,.{decimales}f} MM"
    elif abs(valor) >= 1e6:
        return f"$ {valor/1e6:,.{decimales}f} M"
    elif abs(valor) >= 1e3:
        return f"$ {valor/1e3:,.{decimales}f} K"
    else:
        return f"$ {valor:,.{decimales}f}"

# ============================================
# CLASE PDF
# ============================================

class PDFInforme(FPDF):
    def __init__(self, nombre_proyecto, fecha_corte, mandante=""):
        super().__init__()
        self.nombre_proyecto = nombre_proyecto
        self.fecha_corte = fecha_corte
        self.mandante = mandante
        self.set_auto_page_break(auto=True, margin=15)
        
    def header(self):
        self.set_fill_color(139, 0, 0)
        self.rect(0, 0, 210, 20, 'F')
        self.set_xy(10, 5)
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(255, 255, 255)
        self.cell(50, 10, 'SALFAGESTION', 0, 0, 'L')
        self.set_font('Helvetica', '', 8)
        self.set_xy(100, 5)
        self.cell(100, 5, f'CONTROL DE COSTOS {self.nombre_proyecto} al {self.fecha_corte}', 0, 2, 'R')
        self.cell(100, 5, 'GESTION CONTRACTUAL', 0, 0, 'R')
        self.ln(15)
        self.set_text_color(0, 0, 0)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'GESTION CONTRACTUAL', 0, 0, 'L')
        self.cell(0, 10, f'{self.page_no()}', 0, 0, 'R')
        
    def portada(self):
        self.add_page()
        self.set_fill_color(128, 128, 128)
        self.rect(20, 60, 80, 25, 'F')
        self.set_xy(25, 67)
        self.set_font('Helvetica', 'B', 18)
        self.set_text_color(255, 255, 255)
        self.cell(70, 10, 'SALFAGESTION', 0, 0, 'L')
        
        self.set_fill_color(139, 0, 0)
        self.rect(20, 120, 170, 30, 'F')
        self.set_xy(25, 125)
        self.set_font('Helvetica', 'B', 11)
        self.cell(160, 8, f'INFORME MENSUAL CONTROL DE COSTOS {self.nombre_proyecto} al {self.fecha_corte}', 0, 2, 'L')
        self.cell(160, 8, 'GESTION CONTRACTUAL', 0, 0, 'L')
        self.set_text_color(0, 0, 0)
        
    def indice(self):
        self.add_page()
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(139, 0, 0)
        self.cell(0, 10, 'Indice', 0, 1, 'L')
        self.ln(5)
        self.set_font('Helvetica', '', 10)
        self.set_text_color(0, 0, 0)
        
        indices = [
            ('1. Datos Contractuales', 3),
            ('2. Resumen Ejecutivo', 6),
            ('3. Mayores Variaciones Negativas de Costo (CV)', 16),
            ('4. Variaciones de Ingresos (SV)', 20),
            ('5. Conclusiones', 23),
        ]
        
        for titulo, pag in indices:
            puntos = '.' * (70 - len(titulo))
            self.cell(0, 6, f'{titulo}{puntos}{pag}', 0, 1, 'L')
    
    def seccion_titulo(self, numero, titulo):
        self.add_page()
        self.ln(30)
        self.set_font('Helvetica', 'B', 16)
        self.cell(0, 10, f'{numero}. {titulo}', 0, 1, 'C')
        
    def subtitulo(self, texto):
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(139, 0, 0)
        self.cell(0, 8, texto, 0, 1, 'L')
        self.ln(3)
        self.set_text_color(0, 0, 0)
        
    def texto_normal(self, texto):
        self.set_font('Helvetica', '', 10)
        self.multi_cell(0, 5, texto)
        self.ln(2)
        
    def tabla_simple(self, df, max_rows=15):
        if df.empty:
            return
            
        df_show = df.head(max_rows)
        n_cols = min(len(df_show.columns), 6)
        col_width = 190 / n_cols
        
        self.set_font('Helvetica', 'B', 7)
        self.set_fill_color(139, 0, 0)
        self.set_text_color(255, 255, 255)
        
        for i, col in enumerate(df_show.columns[:n_cols]):
            self.cell(col_width, 6, str(col)[:18], 1, 0, 'C', True)
        self.ln()
        
        self.set_font('Helvetica', '', 6)
        self.set_text_color(0, 0, 0)
        
        for _, row in df_show.iterrows():
            for i, val in enumerate(row.values[:n_cols]):
                if pd.isna(val):
                    text = ''
                elif isinstance(val, float):
                    text = f'{val/1e6:.1f}M' if abs(val) >= 1e6 else f'{val:,.0f}'
                else:
                    text = str(val)[:15]
                self.cell(col_width, 5, text, 1, 0, 'C')
            self.ln()
        self.ln(5)
        
    def generar_completo(self, datos, datos_contrato):
        self.portada()
        self.indice()
        
        # 1. Datos Contractuales
        self.seccion_titulo('1', 'Datos Contractuales')
        self.subtitulo('Datos Contractuales')
        info = f"""
Proyecto: {datos_contrato.get('proyecto', 'N/A')}
Mandante: {datos_contrato.get('mandante', 'N/A')}
Plazo: {datos_contrato.get('plazo', 'N/A')} Dias
Fecha de Inicio: {datos_contrato.get('fecha_inicio', 'N/A')}
Anticipo: {datos_contrato.get('anticipo', 'No presenta')}
        """
        self.texto_normal(info)
        
        # 2. Resumen Ejecutivo
        self.seccion_titulo('2', 'Resumen Ejecutivo')
        if 'resumen_awp' in datos and not datos['resumen_awp'].empty:
            self.subtitulo('Resultado Economico por AWP')
            self.tabla_simple(datos['resumen_awp'])
        
        # 3. Variaciones CV
        self.seccion_titulo('3', 'Mayores Variaciones Negativas de Costo (CV)')
        if 'variaciones_status' in datos and not datos['variaciones_status'].empty:
            self.subtitulo('Variaciones por Status')
            self.tabla_simple(datos['variaciones_status'])
        
        # 4. Variaciones SV
        self.seccion_titulo('4', 'Variaciones de Ingresos (SV)')
        self.texto_normal("Analisis de variaciones de ingreso segun Estado de Pago.")
        
        # 5. Conclusiones
        self.seccion_titulo('5', 'Conclusiones')
        conclusiones = f"""
* Analisis de variaciones de costo completado segun matriz de conversion.
* Se identificaron las principales desviaciones por AWP y STATUS.
* Total de registros procesados: {datos.get('total_registros', 0):,}
* Registros de Ingresos: {datos.get('registros_ingresos', 0):,}
* Registros de Costos: {datos.get('registros_costos', 0):,}
* Informe generado automaticamente el {datetime.now().strftime('%d/%m/%Y %H:%M')}.
        """
        self.texto_normal(conclusiones)
        
        return self

def generar_pdf(datos, nombre_proyecto, fecha_corte, datos_contrato):
    pdf = PDFInforme(nombre_proyecto, fecha_corte)
    pdf.generar_completo(datos, datos_contrato)
    return pdf.output(dest='S').encode('latin-1')

# ============================================
# SESSION STATE
# ============================================
if 'configurado' not in st.session_state:
    st.session_state.configurado = False

# ============================================
# PANTALLA INICIAL
# ============================================
if not st.session_state.configurado:
    
    st.markdown("""
    <div class="main-header">
        📊 CUADRO DE MANDO<br>
        <span style="font-size: 1.3rem;">CONTROL DE COSTOS</span><br>
        <span style="font-size: 0.9rem; font-weight: normal;">SALFAGESTION - GESTIÓN CONTRACTUAL</span>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown('<div class="sub-header">⚙️ Configuración del Proyecto</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📋 Información del Proyecto")
        nombre_proyecto = st.text_input("Nombre del Proyecto *", value="TENIENTE CC - 126")
        mandante = st.text_input("Mandante", value="CODELCO")
        num_contrato = st.text_input("N° Contrato", value="4600027812")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fecha_inicio = st.date_input("Fecha Inicio", value=datetime(2024, 4, 15))
        with col_f2:
            fecha_corte = st.date_input("Fecha de Corte *", value=datetime(2026, 2, 28))
        
        plazo_dias = st.number_input("Plazo (días)", value=686, min_value=1)
    
    with col2:
        st.markdown("#### 📁 Cargar Matriz de Costos")
        st.markdown("""
        <div class="upload-area">
            <h3>📤 Matriz de Conversión Excel</h3>
            <p>Arrastre o seleccione el archivo .xlsx</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Seleccionar archivo", type=['xlsx', 'xls'])
        
        if uploaded_file:
            st.markdown(f"""
            <div class="alert-success">
                ✅ <strong>Archivo:</strong> {uploaded_file.name}<br>
                📦 <strong>Tamaño:</strong> {uploaded_file.size/1024/1024:.2f} MB
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 INICIAR CUADRO DE MANDO", use_container_width=True, type="primary"):
            if not nombre_proyecto:
                st.error("❌ Ingrese el nombre del proyecto")
            elif not uploaded_file:
                st.error("❌ Cargue el archivo Excel")
            else:
                with st.spinner("⏳ Procesando archivo Excel..."):
                    datos, error = cargar_excel(uploaded_file)
                    
                    if error:
                        st.error(f"❌ Error: {error}")
                    else:
                        # Buscar hoja MATRIZ
                        hoja_matriz = None
                        for h in datos['hojas']:
                            if 'MATRIZ' in h.upper():
                                hoja_matriz = h
                                break
                        
                        if not hoja_matriz:
                            hoja_matriz = datos['hojas'][0]
                        
                        datos_proc = procesar_matriz(datos['dataframes'][hoja_matriz])
                        
                        # Guardar en session state
                        st.session_state.nombre_proyecto = nombre_proyecto
                        st.session_state.mandante = mandante
                        st.session_state.num_contrato = num_contrato
                        st.session_state.fecha_inicio = fecha_inicio
                        st.session_state.fecha_corte = fecha_corte
                        st.session_state.plazo_dias = plazo_dias
                        st.session_state.datos = datos
                        st.session_state.datos_proc = datos_proc
                        st.session_state.hoja_matriz = hoja_matriz
                        st.session_state.configurado = True
                        
                        st.success("✅ Datos cargados correctamente")
                        st.rerun()

# ============================================
# DASHBOARD PRINCIPAL
# ============================================
else:
    nombre_proyecto = st.session_state.nombre_proyecto
    fecha_corte = st.session_state.fecha_corte.strftime("%d/%m/%Y")
    datos = st.session_state.datos
    datos_proc = st.session_state.datos_proc
    
    # SIDEBAR
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-logo">
            <h2 style="color: white; margin: 0;">🏗️ SALFAGESTION</h2>
            <p style="color: #bdc3c7; margin: 0;">Control de Costos</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 📋 Proyecto")
        st.success(f"**{nombre_proyecto}**")
        st.info(f"📅 **{fecha_corte}**\n\n🏢 {st.session_state.mandante}")
        
        st.markdown("---")
        st.markdown("### 🧭 Navegación")
        
        pagina = st.radio("", [
            "🏠 Resumen Ejecutivo",
            "📊 Datos Contractuales",
            "🔄 Actividades en Proceso",
            "✅ Actividades Terminadas",
            "📉 Variaciones CV",
            "📈 Variaciones SV",
            "🏗️ Análisis por AWP",
            "👷 Análisis por Disciplina",
            "📋 Generar PDF",
            "📑 Explorar Excel",
            "⚙️ Configuración"
        ])
        
        st.markdown("---")
        st.metric("📊 Registros", f"{datos_proc.get('total_registros', 0):,}")
        
        if st.button("🔄 Nuevo Proyecto", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    
    # HEADER
    st.markdown(f"""
    <div class="main-header">
        📊 CUADRO DE MANDO CONTROL DE COSTOS<br>
        <span style="font-size: 1.3rem;">{nombre_proyecto} al {fecha_corte}</span><br>
        <span style="font-size: 0.9rem; font-weight: normal;">GESTIÓN CONTRACTUAL</span>
    </div>
    """, unsafe_allow_html=True)
    
    # CONTENIDO
    if pagina == "🏠 Resumen Ejecutivo":
        st.markdown('<div class="sub-header">📊 Resumen Ejecutivo</div>', unsafe_allow_html=True)
        
        totales = datos_proc.get('totales', {})
        
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("💰 Ingresos", formato_moneda(totales.get('INGRESOS_ACUM', 0)))
        with c2:
            st.metric("📉 Costos", formato_moneda(totales.get('COSTOS_ACUM', 0)))
        with c3:
            res = totales.get('RESULTADO', 0)
            st.metric("📊 Resultado", formato_moneda(res), "Ganancia" if res >= 0 else "Pérdida", delta_color="normal" if res >= 0 else "inverse")
        with c4:
            st.metric("📈 Margen", f"{totales.get('MARGEN_PCT', 0):.1f}%")
        with c5:
            st.metric("📋 Registros", f"{datos_proc.get('total_registros', 0):,}")
        
        st.markdown("---")
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### 📊 Resultado por AWP")
            if 'resumen_awp' in datos_proc and not datos_proc['resumen_awp'].empty:
                df = datos_proc['resumen_awp']
                fig = px.bar(df, x='AWP', y='RESULTADO', color='RESULTADO',
                            color_continuous_scale=['red', 'yellow', 'green'])
                fig.update_layout(height=400, xaxis_tickangle=-45, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
        
        with c2:
            st.markdown("#### 📈 Distribución por Status")
            if 'status_counts' in datos_proc:
                df_st = pd.DataFrame(list(datos_proc['status_counts'].items()), columns=['Status', 'Cantidad'])
                fig = px.pie(df_st, values='Cantidad', names='Status')
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 📋 Resumen por AWP")
        if 'resumen_awp' in datos_proc and not datos_proc['resumen_awp'].empty:
            st.dataframe(datos_proc['resumen_awp'], use_container_width=True, hide_index=True)
    
    elif pagina == "📊 Datos Contractuales":
        st.markdown('<div class="sub-header">📊 Datos Contractuales</div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"""
            | Campo | Valor |
            |-------|-------|
            | **Proyecto** | {st.session_state.nombre_proyecto} |
            | **Mandante** | {st.session_state.mandante} |
            | **N° Contrato** | {st.session_state.num_contrato} |
            | **Plazo** | {st.session_state.plazo_dias} días |
            | **Fecha Inicio** | {st.session_state.fecha_inicio.strftime('%d/%m/%Y')} |
            | **Fecha Corte** | {fecha_corte} |
            """)
        
        with c2:
            st.markdown("#### 📑 Hojas del Excel")
            for h in datos.get('hojas', []):
                df_h = datos['dataframes'].get(h, pd.DataFrame())
                st.write(f"📄 **{h}**: {len(df_h):,} filas, {len(df_h.columns)} cols")
    
    elif pagina == "🔄 Actividades en Proceso":
        st.markdown('<div class="sub-header">🔄 Actividades en Proceso</div>', unsafe_allow_html=True)
        
        if 'en_proceso' in datos_proc and not datos_proc['en_proceso'].empty:
            df = datos_proc['en_proceso']
            st.metric("Total en Proceso", len(df))
            
            cols = [c for c in ['AREA (*)', 'ID PROGRAMA', 'ID DESCRIPCIÓN', 'STATUS', 'RESULTADO', 'MARGEN %', 'AVANCE REAL'] if c in df.columns]
            if cols:
                st.dataframe(df[cols].head(100), use_container_width=True, hide_index=True)
        else:
            st.info("No se encontraron actividades en proceso")
    
    elif pagina == "✅ Actividades Terminadas":
        st.markdown('<div class="sub-header">✅ Actividades Terminadas</div>', unsafe_allow_html=True)
        
        if 'completadas' in datos_proc and not datos_proc['completadas'].empty:
            df = datos_proc['completadas']
            st.metric("Total Completadas", len(df))
            
            if 'STATUS' in df.columns:
                resumen = df['STATUS'].value_counts().reset_index()
                resumen.columns = ['Status', 'Cantidad']
                fig = px.bar(resumen, x='Status', y='Cantidad')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay actividades completadas en el período")
    
    elif pagina == "📉 Variaciones CV":
        st.markdown('<div class="sub-header">📉 Variaciones de Costo (CV)</div>', unsafe_allow_html=True)
        
        if 'variaciones_status' in datos_proc and not datos_proc['variaciones_status'].empty:
            df = datos_proc['variaciones_status']
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            fig = go.Figure()
            fig.add_trace(go.Bar(name='CV Positivas', x=df['STATUS'], y=df['CV_POSITIVAS'], marker_color='green'))
            fig.add_trace(go.Bar(name='CV Negativas', x=df['STATUS'], y=df['CV_NEGATIVAS'], marker_color='red'))
            fig.update_layout(barmode='relative', xaxis_tickangle=-45, height=500)
            st.plotly_chart(fig, use_container_width=True)
    
    elif pagina == "📈 Variaciones SV":
        st.markdown('<div class="sub-header">📈 Variaciones de Ingreso (SV)</div>', unsafe_allow_html=True)
        st.info("Las variaciones de ingreso se calculan desde la misma matriz con TIPO = INGRESOS")
        
        if 'variaciones_status' in datos_proc:
            st.dataframe(datos_proc['variaciones_status'], use_container_width=True, hide_index=True)
    
    elif pagina == "🏗️ Análisis por AWP":
        st.markdown('<div class="sub-header">🏗️ Análisis por AWP</div>', unsafe_allow_html=True)
        
        if 'resumen_awp' in datos_proc and not datos_proc['resumen_awp'].empty:
            df = datos_proc['resumen_awp']
            
            awp_sel = st.selectbox("Seleccionar AWP:", df['AWP'].tolist())
            fila = df[df['AWP'] == awp_sel].iloc[0]
            
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("💰 Ingreso", formato_moneda(fila['INGRESO_TOTAL']))
            with c2:
                st.metric("📉 Costo", formato_moneda(fila['COSTO_TOTAL']))
            with c3:
                st.metric("📊 Resultado", formato_moneda(fila['RESULTADO']))
            
            st.markdown("---")
            st.dataframe(df, use_container_width=True, hide_index=True)
    
    elif pagina == "👷 Análisis por Disciplina":
        st.markdown('<div class="sub-header">👷 Análisis por Disciplina</div>', unsafe_allow_html=True)
        
        if 'resumen_disciplina' in datos_proc and not datos_proc['resumen_disciplina'].empty:
            st.dataframe(datos_proc['resumen_disciplina'], use_container_width=True, hide_index=True)
        else:
            st.info("No se encontró la columna DISCIPLINA")
    
    elif pagina == "📋 Generar PDF":
        st.markdown('<div class="sub-header">📋 Generar Informe PDF</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="alert-info">
            📄 <strong>Generador de Informes</strong><br>
            Genere un informe PDF con el formato del informe mensual de control de costos.
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📄 GENERAR PDF", use_container_width=True, type="primary"):
            with st.spinner("Generando PDF..."):
                datos_contrato = {
                    'proyecto': st.session_state.nombre_proyecto,
                    'mandante': st.session_state.mandante,
                    'plazo': st.session_state.plazo_dias,
                    'fecha_inicio': st.session_state.fecha_inicio.strftime('%d/%m/%Y'),
                    'anticipo': 'No presenta'
                }
                
                pdf_bytes = generar_pdf(datos_proc, nombre_proyecto, fecha_corte, datos_contrato)
                
                st.download_button(
                    "📥 DESCARGAR PDF",
                    data=pdf_bytes,
                    file_name=f"Informe_{nombre_proyecto.replace(' ', '_')}_{fecha_corte.replace('/', '-')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success("✅ PDF generado")
        
        st.markdown("---")
        st.markdown("#### 🔗 Link del Dashboard")
        st.code("https://[tu-usuario]-cuadro-mando.streamlit.app", language="text")
    
    elif pagina == "📑 Explorar Excel":
        st.markdown('<div class="sub-header">📑 Explorador Excel</div>', unsafe_allow_html=True)
        
        hoja = st.selectbox("Hoja:", datos.get('hojas', []))
        
        if hoja:
            df = datos['dataframes'].get(hoja, pd.DataFrame())
            
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Filas", f"{len(df):,}")
            with c2:
                st.metric("Columnas", len(df.columns))
            with c3:
                st.metric("Celdas", f"{len(df) * len(df.columns):,}")
            
            cols = st.multiselect("Columnas:", df.columns.tolist(), default=df.columns.tolist()[:8])
            filas = st.slider("Filas:", 10, min(500, len(df)), 50)
            
            if cols:
                st.dataframe(df[cols].head(filas), use_container_width=True, hide_index=True)
    
    elif pagina == "⚙️ Configuración":
        st.markdown('<div class="sub-header">⚙️ Configuración</div>', unsafe_allow_html=True)
        
        st.json({
            "Proyecto": st.session_state.nombre_proyecto,
            "Mandante": st.session_state.mandante,
            "Contrato": st.session_state.num_contrato,
            "Fecha Corte": fecha_corte,
            "Hoja Principal": st.session_state.hoja_matriz,
            "Total Registros": datos_proc.get('total_registros', 0),
            "Registros Ingresos": datos_proc.get('registros_ingresos', 0),
            "Registros Costos": datos_proc.get('registros_costos', 0),
        })
        
        if st.button("🗑️ Reiniciar"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    
    # FOOTER
    st.markdown("---")
    st.markdown(f"""
    <div class="footer">
        📊 <strong>Cuadro de Mando Control de Costos</strong><br>
        SALFAGESTION | {nombre_proyecto} | {st.session_state.mandante}<br>
        Datos al {fecha_corte} | Contrato N°{st.session_state.num_contrato}
    </div>
    """, unsafe_allow_html=True)