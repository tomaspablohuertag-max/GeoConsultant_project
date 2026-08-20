# GeoConsult Pro - Gestor y Calculador Inteligente de Kilometraje

## 🎯 Objetivo y Problema a Resolver
En empresas de consultoría y servicios técnicos en terreno, los consultores se desplazan frecuentemente desde sus domicilios particulares hacia diferentes plantas, sedes o clientes durante distintos días del mes. 

Calcular manualmente estas distancias, recordar si son trayectos de ida y vuelta, multiplicarlas por los días trabajados y consolidar todo en reportes mensuales suele generar errores de cálculo, pérdida de tiempo y discrepancias administrativas.

**GeoConsult Pro** resuelve este problema ofreciendo:
1. **Cálculo automatizado y exacto de distancias** conectándose en tiempo real con Google Maps Distance Matrix API y autocompletado de direcciones/empresas con Google Places API.
2. **Cálculo inteligente de trayectos**: Aplica automáticamente la fórmula de ida y vuelta ($\text{distancia} \times 2 \times \text{días}$).
3. **Consolidación automática mensual en Excel**: Genera y mantiene actualizado un archivo Excel estructurado con dos hojas:
   - **Resumen por Consultor:** Total de viajes, días y kilómetros acumulados por cada persona.
   - **Detalle de Trayectos:** Bitácora cronológica completa con origen, destino y fecha.
4. **Sincronización directa con OneDrive**: Mediante OAuth 2.0 y Microsoft Graph API, guarda o actualiza automáticamente el archivo mensual en la nube de Microsoft sin intervención manual.
5. **Exportación local en 1 clic**: Posibilidad de descargar el archivo Excel inmediatamente desde el navegador.

---

## 🛠️ Tecnologías Utilizadas
- **Frontend**: React 19, TypeScript, Tailwind CSS v4, Lucide Icons, Motion (Framer Motion).
- **Mapas & Geocodificación**: `@vis.gl/react-google-maps` (Google Maps Platform).
- **Backend & Servidor**: Node.js, Express, TSX, Axios.
- **Hojas de Cálculo**: SheetJS (`xlsx`).
- **Integración Cloud**: Microsoft Graph API (OneDrive OAuth 2.0).

---

## 🚀 Guía de Instalación y Puesta en Marcha

### 1. Requisitos Previos
- Node.js (versión 18 o superior).
- NPM o Yarn.

### 2. Instalación de Dependencias
Ejecuta en la terminal dentro de la carpeta del proyecto:
```bash
npm install
```

### 3. Configuración de Variables de Entorno
Crea un archivo `.env` en la raíz del proyecto tomando como base `.env.example`:

```env
# Clave de Google Maps (Distance Matrix y Places)
GOOGLE_MAPS_PLATFORM_KEY="AIzaSy..."

# URL donde corre el servidor
APP_URL="http://localhost:3000"

# Credenciales de Microsoft Azure para sincronización con OneDrive (Opcional si usas solo exportación local)
MICROSOFT_CLIENT_ID="tu-client-id-de-azure"
MICROSOFT_CLIENT_SECRET="tu-client-secret-de-azure"
```

### 4. Ejecución en Modo Desarrollo
```bash
npm run dev
```
Abre tu navegador en: [http://localhost:3000](http://localhost:3000)

### 5. Compilación para Producción
```bash
npm run build
npm start
```
