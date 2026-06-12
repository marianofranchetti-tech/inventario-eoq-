# Analizador de Inventarios — EOQ Probabilístico

App web para análisis de inventarios con modelo EOQ probabilístico, clasificación ABC diferenciada e importación de Excel.

## Stack
- **Backend**: Python + Flask
- **Frontend**: HTML/JS/CSS (servido por Flask)
- **Deploy**: Render.com

---

## Deploy en Render — paso a paso

### 1. Subir el código a GitHub

1. Entrá a [github.com](https://github.com) e iniciá sesión
2. Hacé click en **"New repository"** (botón verde arriba a la derecha)
3. Nombre: `inventario-eoq` (o el que quieras)
4. Dejalo en **Public**
5. Click en **"Create repository"**

Ahora subí los archivos. La forma más fácil sin usar la terminal:

6. En la página del repo vacío, hacé click en **"uploading an existing file"**
7. Arrastrá todos los archivos de esta carpeta:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - La carpeta `templates/` con `index.html` adentro
8. Click en **"Commit changes"**

### 2. Conectar con Render

1. Entrá a [render.com](https://render.com) e iniciá sesión
2. Click en **"New +"** → **"Web Service"**
3. Click en **"Connect account"** para conectar tu GitHub (si no lo hiciste antes)
4. Buscá y seleccioná el repo `inventario-eoq`
5. Completá el formulario:
   - **Name**: `inventario-eoq`
   - **Region**: Oregon (o la más cercana)
   - **Branch**: `main`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
6. **Instance Type**: Free
7. Click en **"Create Web Service"**

Render va a buildear y deployar automáticamente. En 2-3 minutos tenés la URL.

### 3. URL final

Una vez deployado, Render te da una URL del tipo:
```
https://inventario-eoq.onrender.com
```

Esa URL es pública y podés compartirla con cualquier cliente.

---

## Estructura del proyecto

```
inventario-eoq/
├── app.py              # Backend Flask — parser Excel + cálculos EOQ
├── requirements.txt    # Dependencias Python
├── Procfile            # Comando de inicio para Render
└── templates/
    └── index.html      # Frontend completo
```

## Endpoints de la API

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/` | Sirve el frontend |
| POST | `/api/parse-excel` | Recibe .xlsx, devuelve productos + cálculos |
| POST | `/api/recalc` | Recalcula con parámetros modificados |

## Formato del Excel esperado

**Hoja única** con esta estructura:

| Fila | Col A | Col B |
|------|-------|-------|
| 5 | Nivel servicio A | 0.99 |
| 6 | Nivel servicio B | 0.97 |
| 7 | Nivel servicio C | 0.95 |
| 8 | Costo pedido % | 0.04 |
| 9 | Costo m²/mes | 9500 |

**Productos desde fila 14:**
- Col A: Código
- Col B: Nombre
- Col C: Volumen m³
- Col E: Lead time (días)
- Col F–Q: Ventas enero–diciembre
- Col R: Cantidad total anual
- Col S: Precio de venta unitario
- Col T: Costo unitario

## Próximos pasos (Fase 2)

- Integración con Odoo via XML-RPC (reemplaza la carga de Excel)
- Autenticación por usuario
- Guardado de análisis históricos
- Exportación de resultados a PDF/Excel
