# 📚 Procesador de Documentos - Diplomados

Una aplicación web desarrollada con Streamlit para procesar automáticamente documentos de sistematización de diplomados y extraer palabras clave relevantes.

## 🌟 Características

- **Procesamiento Automático**: Busca y procesa todos los documentos de sistematización automáticamente
- **Extracción de Keywords**: Utiliza técnicas avanzadas de NLP para extraer las palabras clave más relevantes
- **Interfaz Web Intuitiva**: Interfaz limpia y fácil de usar desarrollada con Streamlit
- **Integración con Google Drive**: Acceso directo a los documentos almacenados en Google Drive
- **Exportación a Excel**: Genera reportes completos en formato Excel

## 🛠️ Tecnologías Utilizadas

- **Streamlit**: Framework para la interfaz web
- **NLTK**: Procesamiento de lenguaje natural
- **KeyBERT**: Extracción de palabras clave
- **Google Drive API**: Integración con Google Drive
- **Pandas**: Manipulación de datos
- **Python-docx**: Procesamiento de documentos Word

## 🚀 Instalación y Configuración

### Requisitos Previos

1. Python 3.8 o superior
2. Credenciales de Google Drive API (`credentials.json`)
3. Acceso a la carpeta de Google Drive con los documentos

### Instalación Local

1. Clona el repositorio:
```bash
git clone <tu-repo-url>
cd procesador-documentos-diplomados
```

2. Instala las dependencias:
```bash
pip install -r requirements.txt
```

3. Configura las credenciales de Google Drive:
   - Descarga el archivo `credentials.json` desde Google Cloud Console
   - Colócalo en el directorio raíz del proyecto

4. Ejecuta la aplicación:
```bash
streamlit run streamlit_app.py
```

### Despliegue en Streamlit Cloud

1. Haz fork de este repositorio
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu repositorio de GitHub
4. Configura las siguientes variables secretas en Streamlit Cloud:
   - Contenido del archivo `credentials.json` como secreto

## 📁 Estructura del Proyecto

```
├── streamlit_app.py          # Aplicación principal de Streamlit
├── main.py                   # Lógica de procesamiento
├── requirements.txt          # Dependencias
├── README.md                 # Documentación
├── credentials.json          # Credenciales de Google Drive (no incluir en repo)
└── token.json               # Token de autenticación (generado automáticamente)
```

## 🔧 Configuración de Google Drive API

1. Ve a [Google Cloud Console](https://console.cloud.google.com/)
2. Crea un nuevo proyecto o selecciona uno existente
3. Habilita la Google Drive API
4. Crea credenciales (OAuth 2.0 Client IDs)
5. Descarga el archivo `credentials.json`
6. Configura los scopes necesarios:
   - `https://www.googleapis.com/auth/drive`

## 💡 Uso de la Aplicación

1. **Configuración**: Ingresa el ID de la carpeta padre en Google Drive
2. **Procesamiento**: Haz clic en "PROCESAR DIPLOMADOS"
3. **Resultados**: Visualiza los resultados y estadísticas
4. **Descarga**: Descarga el reporte en formato Excel

## 📊 Funcionalidades

### Procesamiento Automático
- Busca todas las carpetas de diplomados
- Navega automáticamente a la estructura: `DIPLOMADO > EVIDENCIA DE TRABAJOS > MÓDULO IV`
- Encuentra archivos de sistematización en cada grupo

### Extracción de Datos
- Extrae el título del proyecto
- Procesa el resumen ejecutivo
- Genera 5 palabras clave por documento

### Visualización
- Métricas en tiempo real
- Gráficos de distribución
- Vista previa de datos
- Estadísticas detalladas

## 🔒 Seguridad

- Las credenciales se manejan de forma segura
- No se almacenan datos sensibles en el repositorio
- Autenticación OAuth 2.0 con Google

## 🤝 Contribución

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Añade nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📝 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## ⚠️ Notas Importantes

- Asegúrate de tener permisos de lectura en las carpetas de Google Drive
- El primer uso requiere autenticación manual con Google
- Los tokens se generan automáticamente después de la primera autenticación

## 📞 Soporte

Para reportar bugs o solicitar nuevas funcionalidades, por favor abre un issue en GitHub.

---

Desarrollado con ❤️ usando Streamlit
