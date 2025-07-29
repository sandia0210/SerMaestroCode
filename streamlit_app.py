import streamlit as st
import pandas as pd
import os
import io
import json
from datetime import datetime
import traceback

# Importar tu clase principal
from main import GoogleDriveTopicModelling

# Configuración de la página
st.set_page_config(
    page_title="Procesador de Documentos - Diplomados",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para estilo
st.markdown("""
<style>
    .main {
        background-color: white;
    }
    
    .stButton > button {
        background-color: #28a745;
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: bold;
        border-radius: 0.5rem;
        cursor: pointer;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton > button:hover {
        background-color: #218838;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .stProgress > div > div > div > div {
        background-color: #28a745;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Inicializa las variables de estado de la sesión"""
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'result_df' not in st.session_state:
        st.session_state.result_df = pd.DataFrame()
    if 'processing_status' not in st.session_state:
        st.session_state.processing_status = None
    if 'topic_model' not in st.session_state:
        st.session_state.topic_model = None

def setup_google_credentials():
    """Configura las credenciales de Google Drive desde los secrets de Streamlit"""
    try:
        # Obtener las credenciales desde st.secrets
        if "google_credentials" in st.secrets:
            credentials_info = st.secrets["google_credentials"]
            
            # Crear el archivo credentials.json temporalmente
            credentials_dict = {
                "type": credentials_info.get("type"),
                "project_id": credentials_info.get("project_id"),
                "private_key_id": credentials_info.get("private_key_id"),
                "private_key": credentials_info.get("private_key").replace('\\n', '\n'),
                "client_email": credentials_info.get("client_email"),
                "client_id": credentials_info.get("client_id"),
                "auth_uri": credentials_info.get("auth_uri"),
                "token_uri": credentials_info.get("token_uri"),
                "auth_provider_x509_cert_url": credentials_info.get("auth_provider_x509_cert_url"),
                "client_x509_cert_url": credentials_info.get("client_x509_cert_url"),
                "universe_domain": credentials_info.get("universe_domain", "googleapis.com")
            }
            
            # Si es service account, usar esas credenciales directamente
            if credentials_dict.get("type") == "service_account":
                return "service_account", credentials_dict
            else:
                # Si es OAuth, crear el archivo credentials.json
                with open('credentials.json', 'w') as f:
                    json.dump(credentials_dict, f)
                return "oauth", 'credentials.json'
        else:
            st.error("No se encontraron credenciales de Google en los secrets de Streamlit")
            return None, None
            
    except Exception as e:
        st.error(f"Error al configurar credenciales: {str(e)}")
        return None, None

def authenticate_drive():
    """Función para autenticar con Google Drive"""
    try:
        if st.session_state.topic_model is None:
            st.session_state.topic_model = GoogleDriveTopicModelling(language='spanish')
        
        # Configurar credenciales
        cred_type, cred_data = setup_google_credentials()
        
        if cred_type == "service_account":
            # Usar autenticación de service account
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            
            credentials = service_account.Credentials.from_service_account_info(
                cred_data, scopes=st.session_state.topic_model.SCOPES
            )
            st.session_state.topic_model.service = build('drive', 'v3', credentials=credentials)
            st.success("🔐 Autenticación con Service Account exitosa!")
        elif cred_type == "oauth":
            # Usar autenticación OAuth (requiere interacción del usuario)
            st.session_state.topic_model.authenticate_google_drive()
        else:
            st.error("❌ No se pudieron configurar las credenciales")
            return False
        
        return True
    except Exception as e:
        st.error(f"Error de autenticación: {str(e)}")
        return False

def process_diplomados():
    """Función principal para procesar los diplomados"""
    try:
        # Mostrar mensaje de inicio
        status_container = st.empty()
        progress_bar = st.progress(0)
        
        status_container.info("🔄 Iniciando procesamiento...")
        progress_bar.progress(10)
        
        # Autenticar si no está autenticado
        if st.session_state.topic_model is None or st.session_state.topic_model.service is None:
            status_container.info("🔐 Autenticando con Google Drive...")
            if not authenticate_drive():
                st.error("❌ Error en la autenticación. Verifica tus credenciales.")
                return False
        
        progress_bar.progress(20)
        
        # ID de la carpeta padre (configurable)
        parent_folder_id = st.session_state.get('parent_folder_id', "1-_W-Esk4lzkztPSeZpqO4Gq3ao1P9XKo")
        
        status_container.info("📁 Buscando diplomados...")
        progress_bar.progress(30)
        
        # Procesar todos los diplomados
        status_container.info("⚙️ Procesando documentos...")
        progress_bar.progress(50)
        
        result_df = st.session_state.topic_model.process_all_diplomados(
            parent_folder_id, 
            top_keywords=5
        )
        
        progress_bar.progress(80)
        
        if not result_df.empty:
            st.session_state.result_df = result_df
            st.session_state.processing_complete = True
            st.session_state.processing_status = "success"
            
            status_container.success("✅ Procesamiento completado exitosamente!")
            progress_bar.progress(100)
            
            return True
        else:
            st.session_state.processing_status = "no_data"
            status_container.warning("⚠️ No se encontraron documentos para procesar.")
            return False
            
    except Exception as e:
        st.session_state.processing_status = "error"
        error_msg = f"❌ Error durante el procesamiento: {str(e)}"
        st.error(error_msg)
        
        # Mostrar traceback en expandible para debugging
        with st.expander("Ver detalles del error"):
            st.code(traceback.format_exc())
        
        return False

def main():
    """Función principal de la aplicación Streamlit"""
    
    # Inicializar estado de sesión
    initialize_session_state()
    
    # Título principal
    st.title("📚 Procesador de Documentos - Diplomados")
    st.markdown("---")
    
    # Descripción
    st.markdown("""
    ### Bienvenido al Procesador Automático de Documentos
    
    Esta aplicación procesa automáticamente los documentos de sistematización de todos los diplomados 
    y extrae las palabras clave más relevantes de cada proyecto.
    
    **¿Qué hace la aplicación?**
    - 🔍 Busca automáticamente todas las carpetas de diplomados
    - 📄 Localiza los archivos de sistematización en cada grupo
    - 🔑 Extrae palabras clave usando técnicas de procesamiento de lenguaje natural
    - 📊 Genera un reporte completo en formato Excel
    """)
    
    # Sidebar para configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        # Campo para ID de carpeta padre
        parent_folder_id = st.text_input(
            "ID de Carpeta Padre:", 
            value="1-_W-Esk4lzkztPSeZpqO4Gq3ao1P9XKo",
            help="ID de la carpeta de Google Drive que contiene los diplomados"
        )
        st.session_state.parent_folder_id = parent_folder_id
        
        st.markdown("---")
        
        # Información del estado
        st.header("📊 Estado del Sistema")
        
        if st.session_state.topic_model is not None and hasattr(st.session_state.topic_model, 'service') and st.session_state.topic_model.service is not None:
            st.success("✅ Modelo autenticado")
        else:
            st.info("⏳ Modelo no autenticado")
        
        if st.session_state.processing_complete:
            st.success("✅ Procesamiento completo")
            st.metric("Documentos procesados", len(st.session_state.result_df))
        else:
            st.info("⏳ Sin procesar")
    
    # Sección principal
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### 🚀 Iniciar Procesamiento")
        
        # Botón principal de procesamiento
        if st.button("🔄 PROCESAR DIPLOMADOS", key="process_btn"):
            with st.spinner("Procesando documentos..."):
                success = process_diplomados()
    
    # Mostrar resultados si están disponibles
    if st.session_state.processing_complete and not st.session_state.result_df.empty:
        st.markdown("---")
        st.markdown("## 📊 Resultados del Procesamiento")
        
        # Métricas
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Proyectos", len(st.session_state.result_df))
        
        with col2:
            st.metric("Diplomados", st.session_state.result_df['Diplomado'].nunique())
        
        with col3:
            avg_keywords = st.session_state.result_df[[f'keyword {i+1}' for i in range(5)]].apply(
                lambda x: sum(1 for val in x if val and val.strip()), axis=1
            ).mean()
            st.metric("Promedio Keywords", f"{avg_keywords:.1f}")
        
        with col4:
            st.metric("Archivos Procesados", len(st.session_state.result_df))
        
        # Mostrar DataFrame
        st.markdown("### 📋 Vista Previa de Datos")
        st.dataframe(
            st.session_state.result_df,
            use_container_width=True,
            hide_index=True
        )
        
        # Análisis por diplomado
        st.markdown("### 📈 Distribución por Diplomado")
        diplomado_counts = st.session_state.result_df['Diplomado'].value_counts()
        st.bar_chart(diplomado_counts)
        
        # Botón de descarga
        st.markdown("### 💾 Descargar Resultados")
        
        # Crear archivo Excel en memoria
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            st.session_state.result_df.to_excel(writer, sheet_name='Resultados', index=False)
        
        excel_data = output_buffer.getvalue()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"resultados_diplomados_{timestamp}.xlsx"
        
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Mostrar estadísticas adicionales
        with st.expander("📊 Ver Estadísticas Detalladas"):
            st.markdown("#### Keywords más frecuentes")
            
            # Extraer todas las keywords
            all_keywords = []
            for i in range(5):
                col_name = f'keyword {i+1}'
                keywords = st.session_state.result_df[col_name].dropna()
                keywords = keywords[keywords != '']
                all_keywords.extend(keywords.tolist())
            
            if all_keywords:
                keyword_counts = pd.Series(all_keywords).value_counts().head(10)
                st.bar_chart(keyword_counts)
            else:
                st.info("No se encontraron keywords para mostrar")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>🚀 Procesador de Documentos - Diplomados | Desarrollado con Streamlit</p>
        <p>📧 Para soporte técnico, contacta al administrador del sistema</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
