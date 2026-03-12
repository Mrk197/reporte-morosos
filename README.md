# 📊 Dashboard MikroWisp - Reporte de Clientes

Dashboard web interactivo para visualizar reportes de clientes sin primer pago de MikroWisp.

## 🚀 Instalación Rápida

### 1. Instalar dependencias
```bash
npm install
```

### 2. Configurar variables de entorno
Crea un archivo `.env` en la raíz del proyecto:
```env
MIKROWISP_API_URL=https://tu-servidor.com
MIKROWISP_TOKEN=tu-token-secreto
```

### 3. Ejecutar el servidor
```bash
npm run server
```

La aplicación se abrirá en `http://localhost:3000`

---

## 📋 Características

✅ **Ejecución de Scripts**
- Ejecutar el reporte de clientes directamente desde el navegador
- Parámetros configurables (páginas, perfil, categoría)

✅ **Visualización de Resultados**
- Tablas interactivas con datos de clientes
- Clasificación por categoría (Para Retirar, Suspendidos, Sin Factura)
- Indicadores visuales (badges de estado, colores)

✅ **Descarga de Archivos**
- Descargar resultados en Excel, CSV o JSON
- Nombres automáticos con fecha del día

✅ **Histórico de Ejecuciones**
- Guardar automáticamente un historial de ejecuciones
- Ver estadísticas de cada ejecución
- Datos guardados en `results/history.json`

✅ **Indicador de Progreso**
- Barra de progreso visual durante la ejecución
- Log en tiempo real del proceso
- Mensajes descriptivos

---

## 📊 Parámetros de Ejecución

| Parámetro | Descripción | Defecto |
|-----------|-------------|---------|
| **Páginas a revisar** | Cantidad de últimas páginas a consultar | 20 |
| **Perfil de plan** | Filtro de plan (ej: "2026") | 2026 |
| **Categoría** | `retirar` / `sinFactura` / `todos` | retirar |
| **Formato** | Formato de descarga: `json` / `excel` / `csv` | json |

---

## 📁 Estructura de Archivos

```
package/
├── server.ts              # Servidor Express con APIs
├── index.html             # Interfaz web
├── new-clients-unpaid.ts  # Script principal
├── mikrowisp.ts           # Cliente API MikroWisp
├── package.json           # Dependencias
├── tsconfig.json          # Configuración TypeScript
├── .env                   # Variables de entorno
└── results/               # Carpeta de resultados
    ├── history.json       # Histórico de ejecuciones
    └── latest.json        # Último resultado completo
```

---

## 🔌 API Endpoints

### GET `/`
Sirve la interfaz HTML principal

### POST `/api/run-script`
Ejecuta el script y retorna eventos en tiempo real (Server-Sent Events)

**Body:**
```json
{
  "format": "json",
  "pages": "20",
  "perfil": "2026",
  "categoria": "retirar"
}
```

**Response:** Stream de eventos
```
data: {"status": "iniciando", "mensaje": "..."}
data: {"status": "procesando", "linea": "..."}
data: {"status": "completo", "resultado": {...}}
```

### GET `/api/results`
Obtiene el último resultado completo

**Response:**
```json
{
  "resumen": {...},
  "retirarModem": [...],
  "suspendidosSinPago": [...],
  "sinFacturaAun": [...]
}
```

### GET `/api/history`
Obtiene el histórico de ejecuciones

**Response:**
```json
[
  {
    "timestamp": "2026-03-11T15:30:00.000Z",
    "formato": "json",
    "resumen": {...},
    "totales": {
      "retirarModem": 45,
      "suspendidosSinPago": 12,
      "sinFacturaAun": 8
    }
  }
]
```

### GET `/api/download/:format`
Descarga el archivo del resultado actual

**Parámetros:**
- `format`: `excel` o `csv`
- `pages`: cantidad de páginas (query param)
- `perfil`: perfil de plan (query param)
- `categoria`: categoría (query param)

---

## 🎨 Interfaz

### Secciones Principales

**1. Parámetros de Ejecución**
- Formulario con parámetros configurables
- Botón "Ejecutar Script"
- Botón "Descargar Archivo"

**2. Indicador de Progreso**
- Barra animada durante ejecución
- Log de consola en tiempo real
- Mensajes de estado

**3. Resultados**
- Tarjetas de resumen (Por Retirar, Suspendidos, Sin Factura)
- Tres pestañas con tablas de datos
- Información de cliente completa

**4. Histórico**
- Tabla con historial de ejecuciones
- Timestamps y estadísticas
- Actualizado automáticamente

---

## 🔧 Desarrollo

### Scripts disponibles
```bash
npm run server              # Iniciar servidor (watch mode con tsx)
npm run clientes-sin-pago   # Ejecutar script directamente
npm run clientes-sin-pago:excel  # Generar Excel directamente
```

### Tecnologías usadas
- **Backend:** Node.js + Express + TypeScript
- **Frontend:** HTML5 + CSS3 + Vanilla JavaScript
- **APIs:** Server-Sent Events (SSE) para tiempo real
- **Excel:** ExcelJS para generación de archivos

---

## 📝 Ejemplo de Uso

1. **Abrir en navegador:** `http://localhost:3000`

2. **Configurar parámetros:**
   - Páginas: 30
   - Perfil: 2026
   - Categoría: Retirar Modem
   - Formato: Excel

3. **Ejecutar:** Click en "Ejecutar Script"

4. **Esperar progreso:** Ver log en tiempo real

5. **Ver resultados:** Tabla con clientes

6. **Descargar:** Click en "Descargar Archivo"

---

## ⚠️ Notas Importantes

- El servidor debe estar corriendo para usar la interfaz
- El archivo `.env` debe estar correctamente configurado
- Los archivos descargados se guardaran en la carpeta `package/`
- El histórico se guarda automáticamente en `results/history.json`
- El último resultado completo se guarda en `results/latest.json`

---

## 🐛 Troubleshooting

**Error: "MikroWisp API no configurada"**
- Verifica que existe `.env` con `MIKROWISP_API_URL` y `MIKROWISP_TOKEN`

**Error: "Puerto 3000 en uso"**
- Cambia el puerto en `server.ts` (variable `PORT`)

**Script se cuelga**
- Verifica conexión a internet
- Revisa el timeout (30 segundos en `mikrowisp.ts`)

**Archivos no se descargan**
- Asegurate que `npm run server` está ejecutándose
- Revisa que tienes permisos de escritura en la carpeta

---

## 📞 Contacto

Para soporte o contribuciones, contacta al equipo de DIGY.

**Versión:** 1.0.0  
**Última actualización:** 2026-03-11
