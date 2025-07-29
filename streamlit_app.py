import streamlit as st
import pandas as pd
import os
import io
import json
from datetime import datetime
import traceback

# Importar librerías necesarias para Google Drive
from googleapiclient.discovery import build
from google.oauth2 import service_account

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

def authenticate_drive():
    """Función para autenticar con OAuth usando secrets.toml"""
    try:
        if st.session_state.topic_model is None:
            st.session_state.topic_model = GoogleDriveTopicModelling(language='spanish')
        
        # Verificar si las credenciales OAuth están en st.secrets
        if "oauth_credentials" not in st.secrets:
            st.error("❌ No se encontraron credenciales OAuth en los secrets")
            return False
        
        # Crear credentials.json temporal desde secrets
        oauth_info = st.secrets["oauth_credentials"]
        
        credentials_content = {
            "installed": {
                "client_id": oauth_info["client_id"],
                "project_id": oauth_info["project_id"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": oauth_info["client_secret"],
                "redirect_uris": ["http://localhost"]
            }
        }
        
        # Crear archivo temporal
        import json
        with open("temp_credentials.json", "w") as f:
            json.dump(credentials_content, f)
        
        # Usar el método OAuth original con el archivo temporal
        st.session_state.topic_model.authenticate_google_drive(
            credentials_file="temp_credentials.json",
            token_file="temp_token.json"
        )
        
        # Limpiar archivos temporales
        import os
        if os.path.exists("temp_credentials.json"):
            os.remove("temp_credentials.json")
        
        st.success("🔐 Autenticación OAuth exitosa!")
        return True
        
    except Exception as e:
        st.error(f"❌ Error de autenticación OAuth: {str(e)}")
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
            if st.session_state.topic_model is None or st.session_state.topic_model.service is None:
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
        
        # Verificar estado de credenciales
        if "google_credentials" in st.secrets:
            st.success("✅ Credenciales configuradas")
        else:
            st.error("❌ Credenciales no configuradas")
        
        if st.session_state.topic_model is not None and hasattr(st.session_state.topic_model, 'service') and st.session_state.topic_model.service is not None:
            st.success("✅ Modelo autenticado")
        else:
            st.info("⏳ Modelo no autenticado")
        
        if st.session_state.processing_complete:
            st.success("✅ Procesamiento completo")
            st.metric("Documentos procesados", len(st.session_state.result_df))
        else:
            st.info("⏳ Sin procesar")
        
        # Botón de test de autenticación
        st.markdown("---")
        if st.button("🔧 Probar Autenticación"):
            with st.spinner("Probando autenticación..."):
                if authenticate_drive():
                    st.success("✅ Autenticación exitosa!")
                else:
                    st.error("❌ Error en autenticación")
    
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
