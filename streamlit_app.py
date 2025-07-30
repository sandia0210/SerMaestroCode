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
    page_title="Repositorio de Proyectos SER MAESTRO",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS súper simple - solo lo esencial
st.markdown("""
<style>
    .main {
        background-color: #4faf34;
    }
    
    /* Header verde simple */
    .header-container {
        background-color: #4faf34;
        color: white;
        padding: 2rem;
        margin: -1rem -1rem 2rem -1rem;
    }
    
    .header-title {
        font-size: 2.5rem;
        font-weight: bold;
        margin: 0;
    }
    
    /* Filtros */
    .filters-section {
        margin-bottom: 2rem;
    }
    
    .filters-title {
        font-size: 1.5rem;
        color: #666;
        margin-bottom: 1rem;
    }
    
    /* Resultados */
    .result-item {
        background-color: #f8f9fa;
        border: 1px solid #ddd;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 5px;
    }
    
    .project-title {
        color: #000000;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    
    .project-keywords {
        color: #000000;
        margin-bottom: 0.5rem;
    }
    
    .project-link {
        color: #000000;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Inicializa las variables de estado de la sesión"""
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'result_df' not in st.session_state:
        st.session_state.result_df = pd.DataFrame()
    if 'topic_model' not in st.session_state:
        st.session_state.topic_model = None
    if 'all_keywords' not in st.session_state:
        st.session_state.all_keywords = []

def authenticate_drive():
    """Función para autenticar con Google Drive usando secrets de Streamlit"""
    try:
        if st.session_state.topic_model is None:
            st.session_state.topic_model = GoogleDriveTopicModelling(language='spanish')
        
        if "google_credentials" not in st.secrets:
            st.error("❌ No se encontraron credenciales de Google en los secrets de Streamlit")
            return False
        
        credentials_info = st.secrets["google_credentials"]
        credentials_dict = {}
        for key, value in credentials_info.items():
            if key == "private_key":
                credentials_dict[key] = value.replace('\\n', '\n')
            else:
                credentials_dict[key] = value
        
        credentials = service_account.Credentials.from_service_account_info(
            credentials_dict, 
            scopes=st.session_state.topic_model.SCOPES
        )
        
        st.session_state.topic_model.service = build('drive', 'v3', credentials=credentials)
        return True
        
    except Exception as e:
        st.error(f"❌ Error de autenticación: {str(e)}")
        return False

def process_diplomados():
    """Función principal para procesar los diplomados"""
    try:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("🔐 Autenticando con Google Drive...")
        progress_bar.progress(20)
        
        if not authenticate_drive():
            st.error("❌ Error en la autenticación")
            return False
        
        status_text.text("📁 Buscando diplomados...")
        progress_bar.progress(40)
        
        parent_folder_id = "1-_W-Esk4lzkztPSeZpqO4Gq3ao1P9XKo"
        
        status_text.text("⚙️ Procesando documentos...")
        progress_bar.progress(70)
        
        result_df = st.session_state.topic_model.process_all_diplomados(
            parent_folder_id, 
            top_keywords=5
        )
        
        if not result_df.empty:
            st.session_state.result_df = result_df
            st.session_state.processing_complete = True
            
            # Extraer todas las keywords únicas
            all_keywords = set()
            for i in range(1, 6):  # keyword 1 a keyword 5
                col_name = f'keyword {i}'
                if col_name in result_df.columns:
                    keywords = result_df[col_name].dropna()
                    keywords = keywords[keywords != '']
                    all_keywords.update(keywords.tolist())
            
            st.session_state.all_keywords = sorted(list(all_keywords))
            
            status_text.text("✅ Procesamiento completado!")
            progress_bar.progress(100)
            
            return True
        else:
            st.error("⚠️ No se encontraron documentos para procesar")
            return False
            
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {str(e)}")
        return False

def search_projects(selected_keywords):
    """Busca proyectos que contengan las keywords seleccionadas"""
    if st.session_state.result_df.empty:
        return pd.DataFrame()
    
    if not selected_keywords:
        return pd.DataFrame()
    
    # Filtrar proyectos que contengan alguna de las keywords seleccionadas
    mask = pd.Series([False] * len(st.session_state.result_df))
    
    for i in range(1, 6):  # keyword 1 a keyword 5
        col_name = f'keyword {i}'
        if col_name in st.session_state.result_df.columns:
            mask = mask | st.session_state.result_df[col_name].isin(selected_keywords)
    
    filtered_df = st.session_state.result_df[mask]
    return filtered_df

def main():
    """Función principal de la aplicación Streamlit"""
    
    initialize_session_state()
    
    # Header verde simple
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">Repositorio de Proyectos SER MAESTRO</h1>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar para procesamiento
    with st.sidebar:
        st.image("logo.png", width=200)
     
        st.header("⚙️ Configuración")
        
        if not st.session_state.processing_complete:
            st.info("📋 Primero procesa los documentos")
            if st.button("🔄 Actualizar Base de Datos", use_container_width=True):
                with st.spinner("Procesando documentos..."):
                    process_diplomados()
                    st.rerun()
        else:
            st.success("✅ Base de datos actualizada")
            st.metric("Proyectos encontrados", len(st.session_state.result_df))
            st.metric("Keywords disponibles", len(st.session_state.all_keywords))
            
            if st.button("🔄 Reprocesar Documentos", use_container_width=True):
                st.session_state.processing_complete = False
                st.session_state.result_df = pd.DataFrame()
                st.session_state.all_keywords = []
                st.rerun()
    
    # Contenido principal
    if not st.session_state.processing_complete:
        st.info("📋 Haz clic en 'Actualizar Base de Datos' en el panel lateral para comenzar")
    else:
        # Sección de filtros
        st.markdown('<div class="filters-section">', unsafe_allow_html=True)
        st.markdown('<h2 class="filters-title">🔍 Filtros</h2>', unsafe_allow_html=True)
        
        st.text("Seleccione el eje temático")
        
        # Multiselect para keywords
        selected_keywords = st.multiselect(
            "",
            options=st.session_state.all_keywords,
            placeholder="Selecciona palabras clave...",
            label_visibility="collapsed"
        )
        
        # Botón de búsqueda
        search_clicked = st.button("Buscar Proyectos", type="primary")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Mostrar resultados
        if search_clicked or selected_keywords:
            filtered_projects = search_projects(selected_keywords)
            
            if not filtered_projects.empty:
                st.markdown("### Proyectos encontrados:")
                
                for _, project in filtered_projects.iterrows():
                    # Obtener keywords del proyecto
                    project_keywords = []
                    for i in range(1, 6):
                        col_name = f'keyword {i}'
                        if col_name in project.index and project[col_name] and project[col_name].strip():
                            project_keywords.append(project[col_name])
                    
                    # Crear el item de resultado
                    st.markdown(f"""
                    <div class="result-item">
                        <div class="project-title">📋<strong style='color:#001d57;'> Título del proyecto: </strong>"{project.get('Título del proyecto', 'Sin título')}"</div>
                        <div class="project-keywords">🔑<strong style='color:#001d57;'> Palabras clave:</strong> {', '.join(project_keywords)}</div>
                        <div class="project-link">🔗 <strong style='color:#001d57;'>Proyecto / documento: </strong><a href="{project.get('Enlace de descarga', '#')}" target="_blank">Enlace de descarga</a></div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.success(f"Se encontraron {len(filtered_projects)} proyectos")
            else:
                if selected_keywords:
                    st.warning("No se encontraron proyectos con las palabras clave seleccionadas")
        else:
            st.info("Selecciona palabras clave y haz clic en 'Buscar Proyectos' para ver los resultados")

if __name__ == "__main__":
    main()
