# ⏱️ Control de Horas (por día, semana, mes)  

Aplicación en **Streamlit** para cargar un Excel de fichajes, calcular horas trabajadas por día, agregar subtotales semanales y mensuales, mostrar tablas interactivas y gráficos, y exportar informes en Excel (incluyendo ZIP con un archivo por trabajador).

---

## 🚀 Funcionalidades

- **Carga de Excel**: Se usa la fila 7 como encabezado (índice 6 en pandas).
- **Procesado automático**:
  - Marca como *Sin registro* cuando no hay fichaje de salida.
  - Calcula horas trabajadas por día, semana y mes.
  - Subtotales semanales por usuario.
- **Filtros multiselección** por usuario y semana.
- **Tablas interactivas** con **AgGrid** (filtros, ordenación, redimensionado de columnas).
- **Gráficos** con Plotly:
  - Horas por día.
  - Horas por semana.
  - Horas por mes.
  - Top personas con más horas.
- **Descargas**:
  - Excel global filtrado (Resumen + Totales semana + Totales mes).
  - ZIP con un Excel por trabajador.

---

## 📦 Estructura del proyecto

.
├── app.py # Código principal de Streamlit
├── requirements.txt # Dependencias del proyecto
└── README.md # Este archivo

yaml
Copiar
Editar

---

## 🛠️ Requisitos

Archivo `requirements.txt`:

streamlit==1.37.1
pandas==2.2.2
numpy==1.26.4
openpyxl==3.1.5
xlsxwriter==3.2.0
plotly==5.22.0
streamlit-aggrid==0.3.5
streamlit-extras==0.4.3

markdown
Copiar
Editar

---

## 📤 Despliegue en Streamlit Cloud

1. **Subir a GitHub**  
   Sube tu repositorio con:
   - `app.py`
   - `requirements.txt`
   - `README.md`

2. **Conectar Streamlit Cloud**  
   - Ve a [streamlit.io/cloud](https://streamlit.io/cloud).
   - Conecta tu cuenta con GitHub.
   - Selecciona tu repositorio y la rama (`main` o `master`).

3. **Configurar despliegue**  
   - Archivo principal: `app.py`.
   - Python version: 3.9+ (Streamlit Cloud la elige automáticamente).

4. **Deploy**  
   Pulsa **Deploy** y espera a que se instalen dependencias.

---

## 📂 Uso de la aplicación

1. **Subir archivo Excel** con:
   - Fila 7 como encabezado.
   - Columnas requeridas: `Usuario`, `Nombre`, `Apellidos`, `Inicio`, `Fin`.

2. **Filtrar** por usuario/semana desde la UI.

3. **Analizar** en las tablas interactivas y gráficos.

4. **Descargar**:
   - Excel global filtrado.
   - ZIP con un Excel por trabajador.

---

## 🎨 Estilo

- Colores neutros (blanco, negro, beige) para la interfaz.
- Toques de azul vivo para elementos destacados.
- Tarjetas con borde redondeado y sombra suave.

---

## 📸 Capturas de pantalla (opcional)

*(Añadir capturas una vez desplegada la app para documentar la UI)*

---

## 📝 Notas

- Si el campo **Fin** tiene fecha `01/01/0001 00:00:00` o está vacío, se considera *Sin registro*.
- Los subtotales y totales semanales/mensuales se calculan solo con fichajes válidos.