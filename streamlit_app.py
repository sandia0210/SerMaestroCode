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
    page_title="Repositorio de Proyectos - Ser Maestro",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para el nuevo diseño
st.markdown("""
<style>
    /* Fondo general */
    .main {
        background: linear-gradient(135deg, #3e8e2c 0%, #4faf34 100%);
        min-height: 100vh;
    }
    
    /* Header personalizado */
    .custom-header {
        background: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        border-radius: 15px;
        padding: 2rem;
        margin-bottom: 2rem;
        text-align: center;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    .header-title {
        color: white;
        font-size: 3rem;
        font-weight: bold;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .header-subtitle {
        color: rgba(255, 255, 255, 0.9);
        font-size: 1.2rem;
        margin-top: 0.5rem;
    }
    
    /* Sección de filtros */
    .filters-section {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    .filters-title {
        color: #2d6a4f;
        font-size: 1.8rem;
        font-weight: bold;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    /* Botones personalizados */
    .stButton > button {
        background: linear-gradient(135deg, #2d6a4f 0%, #40916c 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: bold;
        border-radius: 50px;
        cursor: pointer;
        transition: all 0.3s ease;
        width: 100%;
        box-shadow: 0 4px 15px rgba(45, 106, 79, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(45, 106, 79, 0.4);
        background: linear-gradient(135deg, #40916c 0%, #52b788 100%);
    }
    
    /* Botón de actualizar base de datos - estilo especial */
    .update-button {
        background: linear-gradient(135deg, #1976d2 0%, #2196f3 100%) !important;
        box-shadow: 0 4px 15px rgba(25, 118, 210, 0.3) !important;
    }
    
    .update-button:hover {
        background: linear-gradient(135deg, #2196f3 0%, #42a5f5 100%) !important;
        box-shadow: 0 6px 20px rgba(25, 118, 210, 0.4) !important;
    }
    
    /* Resultados de proyectos */
    .projects-section {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    .projects-title {
        color: #2d6a4f;
        font-size: 1.8rem;
        font-weight: bold;
        margin-bottom: 1.5rem;
    }
    
    /* Cards de proyectos */
    .project-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-left: 4px solid #52b788;
        transition: all 0.3s ease;
    }
    
    .project-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }
    
    .project-title {
        color: #2d6a4f;
        font-size: 1.3rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .project-keywords {
        color: #666;
        font-size: 0.95rem;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .project-link {
        color: #1976d2;
        text-decoration: none;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .project-link:hover {
        color: #2196f3;
        text-decoration: underline;
    }
    
    /* Multiselect personalizado */
    .stMultiSelect > div > div {
        background-color: white;
        border-radius: 8px;
    }
    
    /* Logo placeholder */
    .logo-placeholder {
        width: 120px;
        height: 60px;
        background: rgba(255, 255, 255, 0.2);
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1rem auto;
        border: 2px dashed rgba(255, 255, 255, 0.4);
        color: rgba(255, 255, 255, 0.8);
        font-size: 0.9rem;
    }
    
    /* Sidebar personalizado */
    .css-1d391kg {
        background: linear-gradient(180deg, #2d6a4f 0%, #40916c 100%);
    }
    
    .css-1d391kg .css-1v0mbdj {
        color: white;
    }
    
    /* Progress bar personalizado */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #52b788, #40916c);
    }
    
    /* Alertas personalizadas */
    .stAlert {
        border-radius: 12px;
        border: none;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    /* Hide streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Inicializa las variables de estado de la sesión"""
    if 'database_loaded' not in st.session_state:
        st.session_state.database_loaded = False
    if 'projects_df' not in st.session_state:
        st.session_state.projects_df = pd.DataFrame()
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

def update_database():
    """Función para actualizar la base de datos de proyectos"""
    try:
        # Crear contenedores para mostrar el progreso
        status_container = st.empty()
        progress_container = st.empty()
        
        # Mostrar estado inicial con spinner más prominente
        with status_container.container():
            st.markdown("""
            <div style='text-align: center; padding: 2rem; background: rgba(33, 150, 243, 0.1); border-radius: 15px; margin: 1rem 0;'>
                <div style='font-size: 3rem; margin-bottom: 1rem;'>🔄</div>
                <h3 style='color: #1976d2; margin: 0;'>Actualizando Base de Datos</h3>
                <p style='color: #666; margin: 0.5rem 0;'>Por favor espera, este proceso puede tomar varios minutos...</p>
            </div>
            """, unsafe_allow_html=True)
        
        with progress_container.container():
            progress_bar = st.progress(0)
            st.markdown("<p style='text-align: center; color: #666;'>Iniciando proceso...</p>", unsafe_allow_html=True)
        
        # Paso 1: Autenticación
        progress_bar.progress(10)
        progress_container.markdown("<p style='text-align: center; color: #666;'>🔐 Autenticando con Google Drive...</p>", unsafe_allow_html=True)
        
        if st.session_state.topic_model is None or st.session_state.topic_model.service is None:
            if not authenticate_drive():
                status_container.error("❌ Error en la autenticación.")
                return False
        
        # Paso 2: Configuración
        progress_bar.progress(20)
        progress_container.markdown("<p style='text-align: center; color: #666;'>📁 Configurando acceso a archivos...</p>", unsafe_allow_html=True)
        
        parent_folder_id = "1-_W-Esk4lzkztPSeZpqO4Gq3ao1P9XKo"
        
        # Paso 3: Búsqueda de diplomados
        progress_bar.progress(30)
        progress_container.markdown("<p style='text-align: center; color: #666;'>🔍 Buscando diplomados...</p>", unsafe_allow_html=True)
        
        # Paso 4: Procesamiento principal
        progress_bar.progress(40)
        progress_container.markdown("<p style='text-align: center; color: #666;'>⚙️ Procesando documentos (esto puede tomar varios minutos)...</p>", unsafe_allow_html=True)
        
        result_df = st.session_state.topic_model.process_all_diplomados(
            parent_folder_id, 
            top_keywords=5
        )
        
        # Paso 5: Procesando resultados
        progress_bar.progress(80)
        progress_container.markdown("<p style='text-align: center; color: #666;'>📊 Organizando resultados...</p>", unsafe_allow_html=True)
        
        if not result_df.empty:
            st.session_state.projects_df = result_df
            st.session_state.database_loaded = True
            
            # Extraer todas las keywords únicas
            all_keywords = set()
            for i in range(1, 6):
                col_name = f'keyword {i}'
                if col_name in result_df.columns:
                    keywords = result_df[col_name].dropna()
                    keywords = keywords[keywords != '']
                    all_keywords.update(keywords.tolist())
            
            st.session_state.all_keywords = sorted(list(all_keywords))
            
            # Finalización exitosa
            progress_bar.progress(100)
            progress_container.markdown("<p style='text-align: center; color: #666;'>✅ Proceso completado</p>", unsafe_allow_html=True)
            
            # Mostrar resultado exitoso
            status_container.markdown("""
            <div style='text-align: center; padding: 2rem; background: rgba(76, 175, 80, 0.1); border-radius: 15px; margin: 1rem 0;'>
                <div style='font-size: 3rem; margin-bottom: 1rem;'>✅</div>
                <h3 style='color: #4caf50; margin: 0;'>¡Base de Datos Actualizada!</h3>
                <p style='color: #666; margin: 0.5rem 0;'>Se procesaron {} proyectos exitosamente</p>
            </div>
            """.format(len(result_df)), unsafe_allow_html=True)
            
            # Limpiar contenedores después de 3 segundos
            import time
            time.sleep(3)
            status_container.empty()
            progress_container.empty()
            
            return True
        else:
            status_container.warning("⚠️ No se encontraron documentos para procesar.")
            progress_container.empty()
            return False
            
    except Exception as e:
        # Mostrar error prominente
        status_container.markdown(f"""
        <div style='text-align: center; padding: 2rem; background: rgba(244, 67, 54, 0.1); border-radius: 15px; margin: 1rem 0;'>
            <div style='font-size: 3rem; margin-bottom: 1rem;'>❌</div>
            <h3 style='color: #f44336; margin: 0;'>Error en la Actualización</h3>
            <p style='color: #666; margin: 0.5rem 0;'>{str(e)}</p>
        </div>
        """, unsafe_allow_html=True)
        progress_container.empty()
        return False

def search_projects(selected_keywords):
    """Busca proyectos que contengan al menos una de las keywords seleccionadas"""
    if st.session_state.projects_df.empty or not selected_keywords:
        return pd.DataFrame()
    
    # Crear máscara para encontrar proyectos que contengan al menos una keyword
    mask = pd.Series([False] * len(st.session_state.projects_df))
    
    for i, row in st.session_state.projects_df.iterrows():
        project_keywords = []
        for j in range(1, 6):
            col_name = f'keyword {j}'
            if col_name in row and pd.notna(row[col_name]) and row[col_name] != '':
                project_keywords.append(row[col_name])
        
        # Verificar si alguna keyword seleccionada está en las keywords del proyecto
        if any(keyword in project_keywords for keyword in selected_keywords):
            mask.iloc[i] = True
    
    return st.session_state.projects_df[mask]

def display_project_card(project):
    """Muestra una tarjeta de proyecto"""
    # Obtener keywords del proyecto
    keywords = []
    for i in range(1, 6):
        col_name = f'keyword {i}'
        if col_name in project and pd.notna(project[col_name]) and project[col_name] != '':
            keywords.append(project[col_name])
    
    keywords_text = ', '.join(keywords) if keywords else 'Sin palabras clave'
    
    st.markdown(f"""
    <div class="project-card">
        <div class="project-title">
            📚 {project['Título del proyecto']}
        </div>
        <div class="project-keywords">
            🔑 <strong>Palabras clave:</strong> {keywords_text}
        </div>
        <a href="{project['Enlace de descarga']}" target="_blank" class="project-link">
            🔗 Proyecto / documento: Enlace de descarga
        </a>
    </div>
    """, unsafe_allow_html=True)

def main():
    """Función principal de la aplicación"""
    initialize_session_state()
    
    # Header con logo placeholder
    st.markdown("""
    <div class="custom-header">
        <div class="logo-placeholder">
            Logo aquí
        </div>
        <h1 class="header-title">Repositorio de Proyectos</h1>
        <p class="header-subtitle">Comunidad de Aprendizaje</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar para configuración
    with st.sidebar:
        st.markdown("""
        <div style='color: white; text-align: center; padding: 1rem;'>
            <h2>⚙️ Configuración</h2>
        </div>
        """, unsafe_allow_html=True)
        
        # Botón de actualizar base de datos
        st.markdown("### 📊 Base de Datos")
        
        # Verificar si hay una actualización en proceso
        if 'updating_database' not in st.session_state:
            st.session_state.updating_database = False
        
        # Botón con estado dinámico
        button_text = "🔄 Actualizando..." if st.session_state.updating_database else "🔄 Actualizar Base de Datos"
        button_disabled = st.session_state.updating_database
        
        if st.button(button_text, key="update_db", disabled=button_disabled, help="Procesa y actualiza todos los proyectos"):
            st.session_state.updating_database = True
            
            # Mostrar mensaje prominente de carga
            st.markdown("""
            <div style='background: rgba(255, 193, 7, 0.1); border-left: 4px solid #ffc107; padding: 1rem; margin: 1rem 0; border-radius: 8px;'>
                <strong>⏳ Proceso en ejecución</strong><br>
                <small>La actualización puede tomar varios minutos. Por favor, no cierres la ventana.</small>
            </div>
            """, unsafe_allow_html=True)
            
            # Ejecutar actualización
            success = update_database()
            
            # Resetear estado
            st.session_state.updating_database = False
            
            if success:
                st.rerun()  # Refrescar la página para mostrar los nuevos datos
        
        st.markdown("---")
        
        # Estado del sistema
        st.markdown("### 📈 Estado del Sistema")
        
        if st.session_state.database_loaded:
            st.success(f"✅ Base de datos cargada")
            st.metric("Proyectos disponibles", len(st.session_state.projects_df))
            st.metric("Keywords únicas", len(st.session_state.all_keywords))
        else:
            st.info("⏳ Base de datos no cargada")
            st.info("💡 Haz clic en 'Actualizar Base de Datos' para cargar los proyectos")
    
    # Sección de filtros
    st.markdown("""
    <div class="filters-section">
        <div class="filters-title">
            🔍 Filtros
        </div>
        <p style="color: #666; margin-bottom: 1.5rem;">Seleccione el eje temático</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Selector de keywords (solo si la base de datos está cargada)
    selected_keywords = []
    if st.session_state.database_loaded and st.session_state.all_keywords:
        selected_keywords = st.multiselect(
            "Selecciona palabras clave:",
            options=st.session_state.all_keywords,
            placeholder="Escribe o selecciona palabras clave...",
            help="Puedes seleccionar múltiples palabras clave. Se mostrarán proyectos que contengan al menos una de ellas."
        )
    elif not st.session_state.database_loaded:
        st.info("💡 Primero actualiza la base de datos para ver las palabras clave disponibles")
    
    # Botón de búsqueda
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        search_clicked = st.button("🔍 Buscar Proyectos", disabled=not selected_keywords)
    
    # Sección de resultados
    st.markdown("""
    <div class="projects-section">
        <div class="projects-title">Proyectos encontrados:</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Mostrar resultados
    if search_clicked and selected_keywords:
        with st.spinner("Buscando proyectos..."):
            filtered_projects = search_projects(selected_keywords)
            
            if not filtered_projects.empty:
                st.success(f"Se encontraron {len(filtered_projects)} proyectos")
                
                # Mostrar cada proyecto como una tarjeta
                for _, project in filtered_projects.iterrows():
                    display_project_card(project)
            else:
                st.warning("No se encontraron proyectos con las palabras clave seleccionadas")
    
    elif st.session_state.get('updating_database', False):
        # Mostrar mensaje de carga en la sección principal también
        st.markdown("""
        <div style='text-align: center; padding: 3rem; background: rgba(255, 193, 7, 0.05); border-radius: 15px; margin: 2rem 0;'>
            <div style='font-size: 4rem; margin-bottom: 1rem; animation: pulse 2s infinite;'>⏳</div>
            <h2 style='color: #ff9800; margin: 0;'>Actualizando Base de Datos</h2>
            <p style='color: #666; margin: 1rem 0; font-size: 1.1rem;'>Este proceso puede tomar varios minutos...</p>
            <p style='color: #999; font-size: 0.9rem;'>Por favor, mantén esta ventana abierta hasta que termine</p>
        </div>
        
        <style>
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.1); }
            100% { transform: scale(1); }
        }
        </style>
        """, unsafe_allow_html=True)
    
    elif not st.session_state.database_loaded:
        st.info("🚀 ¡Bienvenido al Repositorio de Proyectos!\n\nPara comenzar, actualiza la base de datos desde el panel lateral.")
    
    elif not selected_keywords and st.session_state.database_loaded:
        st.info("📝 Selecciona palabras clave y haz clic en 'Buscar Proyectos' para ver los resultados.")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: rgba(255,255,255,0.8); padding: 20px;'>
        <p>🌱 Repositorio de Proyectos - Ser Maestro | Comunidad de Aprendizaje</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
