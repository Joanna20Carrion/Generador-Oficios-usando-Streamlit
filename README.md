# 📨 Generador de Oficios Personalizados

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat&logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-Web_App-ff4b4b?style=flat&logo=streamlit)
![Pandas](https://img.shields.io/badge/Pandas-Data--Processing-purple?style=flat&logo=pandas)
![python-docx](https://img.shields.io/badge/python--docx-Word_Generator-blueviolet?style=flat)

---

## 📝 Descripción

Aplicación web desarrollada en **Streamlit** que permite generar **oficios personalizados en Word** a partir de una **plantilla `.docx`** y datos contenidos en un **archivo Excel**.  
Los documentos se procesan dinámicamente en memoria y se empaquetan automáticamente en un **archivo ZIP** para descarga inmediata, sin almacenar archivos en el servidor.

---

## 🎯 Funcionalidades

- Subida de archivos `.xlsx` y `.docx` desde la interfaz web  
- Selección múltiple de **empresas** extraídas del Excel  
- Reemplazo automático de campos en la plantilla:
  - `[Nombre del Destinatario]`
  - `[Cargo]`
  - `[Entidad]`
  - `[Dirección]`
  - `[Distrito]`
- Adjuntar PDFs específicos según la **actividad**:
  - Transmisión  
  - Generación  
  - Distribución  
  - Cliente Libre  
- Generación de **ZIP estructurado por carpetas** (`Actividad/Código/Oficio.docx`)
- Diseño limpio y moderno con colores suaves, estilo **Bootstrap-like**

---

## 💻 Tecnologías utilizadas

- ![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat&logo=python)  
- ![Streamlit](https://img.shields.io/badge/Streamlit-UI_Framework-ff4b4b?style=flat&logo=streamlit)  
- ![Pandas](https://img.shields.io/badge/Pandas-Excel_Data-purple?style=flat&logo=pandas)  
- ![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel_Reader-yellowgreen?style=flat)  
- ![python-docx](https://img.shields.io/badge/python--docx-Word_Automation-blueviolet?style=flat)  

---

## ⚙️ Requisitos

Asegúrate de tener las dependencias necesarias instaladas:

```bash
pip install -r requirements.txt
```

**requirements.txt**
```
streamlit
pandas
openpyxl
python-docx
```

---

## 🚀 Ejecución local

1. Clona este repositorio:

```bash
git clone https://github.com/tu_usuario/generador-oficios.git
cd generador-oficios
```

2. Ejecuta la aplicación:
```bash
streamlit run app.py
```

3. Abre el enlace que aparece en la terminal (por defecto: http://localhost:8501)

---

## ☁️ Despliegue en Streamlit Cloud

1. Sube este proyecto a tu repositorio de **GitHub**  
2. Ingresa a [https://share.streamlit.io](https://share.streamlit.io)  
3. Conecta tu cuenta de GitHub y selecciona este repo  
4. Define `app.py` como archivo principal  
5. ¡Listo! 🚀 Tu aplicación estará en línea en pocos segundos

---

## 👤 Autora

**Joanna Alexandra Carrión Pérez**  
🎓 Bachiller en Ingeniería Electrónica  
🚀 Apasionada por la ciencia de datos y sistemas inteligentes  
📧 joannacarrion14@gmail.com  
🔗 [![LinkedIn](https://img.shields.io/badge/LinkedIn-Joanna%20Carrión%20Pérez-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/joanna-carrion-perez/)

---

## 🧩 Contribuciones

¿Tienes ideas o mejoras?  
Haz un **fork**, realiza tus cambios y envía un **pull request**.  
Las contribuciones son siempre bienvenidas 💡✨
