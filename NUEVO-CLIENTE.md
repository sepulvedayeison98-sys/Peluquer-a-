# 🚀 Crear nuevo cliente — Guía rápida (15 minutos)

## Paso 1 — Duplicar el repositorio (2 min)
1. En GitHub: **Fork** o **Use this template** del repositorio
2. Nómbralo: `nombre-del-negocio`
3. Activa GitHub Pages: Settings → Pages → Branch: main

## Paso 2 — Editar config.js (5 min)
Abre `config.js` y cambia **solo estos campos**:

```js
const BUSINESS_CONFIG = {
  name:        'Nombre del Negocio',    // ← CAMBIAR
  tagline:     'Slogan del negocio',    // ← CAMBIAR
  slogan:      'Since 2024',            // ← CAMBIAR
  type:        'Peluquería',            // ← CAMBIAR (Barbería | Spa | etc.)
  monogram:    'NN',                    // ← iniciales del negocio
  established: '2024',                  // ← año de fundación

  phone:     '+57 300 000 0000',        // ← CAMBIAR
  whatsapp:  '+57 300 000 0000',        // ← CAMBIAR
  email:     'negocio@ejemplo.com',     // ← CAMBIAR
  address:   'Calle 1 # 2-3',          // ← CAMBIAR
  city:      'Ciudad',                  // ← CAMBIAR

  schedule: {
    label:    '08:00 AM — 06:00 PM',    // ← CAMBIAR
    openHour:  8,                       // ← hora apertura
    closeHour: 18,                      // ← hora cierre
  },

  // Servicios del negocio
  services: [
    { id: 'servicio-1', name: 'Nombre servicio', price: 30000, duration: 30, active: true },
    // agregar más servicios...
  ],

  // URLs técnicas (ver Paso 3 y 4)
  backendUrl: '',    // ← URL de Apps Script (Paso 3)
  n8nWebhook: '',    // ← URL de N8N (Paso 4)
  barberoEmail: '',  // ← correo que recibe notificaciones

  panel: {
    user: 'admin',          // ← usuario del panel
    pass: 'clave2024',      // ← contraseña del panel
    sessionKey: 'neg_v1'    // ← clave única (cambiar por cliente)
  }
};
```

## Paso 3 — Configurar Google Apps Script (5 min)
1. Ve a [script.google.com](https://script.google.com) → **Nuevo proyecto**
2. Copia el contenido de `backend-apps-script.js`
3. Cambia en el script:
   - `BUSINESS_NAME` → nombre del negocio
   - `BARBERO_EMAIL` → correo del cliente
   - `N8N_WEBHOOK` → URL del webhook N8N (Paso 4)
4. **Desplegar → Nueva implementación → Aplicación web**
   - Ejecutar como: **Yo**
   - Acceso: **Cualquier persona**
5. Copia la URL generada y pégala en `config.js` → `backendUrl`

## Paso 4 — Configurar N8N (3 min)
1. En N8N: importa el archivo `barber-style-workflow.json`
2. Conecta las credenciales de Gmail en cada nodo
3. Cambia el correo del barbero (`sepulvedayeison98@gmail.com`) por el del cliente
4. Activa el workflow
5. Copia la URL del webhook y pégala en `config.js` → `n8nWebhook`

## Paso 5 — Desplegar (1 min)
```bash
git add .
git commit -m "config: nuevo cliente - [nombre negocio]"
git push origin main
```

GitHub Pages desplegará automáticamente en ~2 minutos.

---

## Checklist final
- [ ] `config.js` con datos del negocio
- [ ] Servicios configurados en `config.js`
- [ ] Apps Script desplegado y URL en `config.js`
- [ ] N8N workflow activo y URL en `config.js`
- [ ] Panel accesible con las credenciales configuradas
- [ ] Prueba de agendamiento desde el sitio
- [ ] Prueba de correo de confirmación

## Estructura de archivos
```
/
├── config.js          ← ✏️  EDITAR POR CLIENTE
├── reservas.html      ← página de reservas (no tocar)
├── panel.html         ← panel del barbero (no tocar)
├── backend-apps-script.js  ← copiar a Apps Script (cambiar 3 constantes)
└── NUEVO-CLIENTE.md   ← esta guía
```
