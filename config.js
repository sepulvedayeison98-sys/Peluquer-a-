/**
 * ============================================================
 *  WHITE LABEL CONFIG — Edita solo este archivo por cliente
 * ============================================================
 *  Para crear un nuevo cliente:
 *  1. Duplica el repositorio
 *  2. Edita este archivo con los datos del nuevo negocio
 *  3. Actualiza BACKEND_URL y N8N_WEBHOOK con las nuevas URLs
 *  4. Despliega en GitHub Pages
 * ============================================================
 */

const BUSINESS_CONFIG = {

  /* ── IDENTIDAD DEL NEGOCIO ─────────────────────────────── */
  name:        'Barber Style',
  tagline:     'Barbería de élite',
  slogan:      'Since 2023',
  description: 'Tu barbería de confianza en Medellín',
  type:        'Barbería',           // Barbería | Peluquería | Spa | Consultorio | etc.
  monogram:    'BS',
  established: '2023',

  /* ── CONTACTO ──────────────────────────────────────────── */
  phone:     '+573001234567',
  whatsapp:  '+573001234567',
  email:     'barbero@barberstyle.co',
  address:   'Carrera 55 24A-06',
  city:      'Medellín, Bello',
  country:   'Colombia',

  /* ── HORARIO ───────────────────────────────────────────── */
  schedule: {
    label:    '07:00 AM — 07:00 PM',
    openHour:  7,   // hora de apertura (24h)
    closeHour: 19,  // hora de cierre (24h)
    slotMinutes: 30 // duración mínima de cada slot
  },

  /* ── REDES SOCIALES ────────────────────────────────────── */
  social: {
    instagram: 'https://instagram.com/barberstyle',
    facebook:  '',
    tiktok:    '',
    youtube:   ''
  },

  /* ── RESEÑAS ───────────────────────────────────────────── */
  reviews: {
    score: '5.0',
    count: '218 reseñas'
  },

  /* ── COLORES (CSS variables) ───────────────────────────── */
  colors: {
    primary:   '#b8965a',   // dorado
    secondary: '#1a1410',   // oscuro
    accent:    '#c9a96e',   // dorado claro
    bg:        '#f4f0eb'    // fondo
  },

  /* ── TÉCNICO ───────────────────────────────────────────── */
  backendUrl: 'https://script.google.com/macros/s/AKfycbyo12J2gDyy9-6m75IOTpnYsvdACKHw63WrlEmfkxOpwOgRRCseIZWr1KlwTb-SQhRyMg/exec',
  n8nWebhook: 'https://kodrefe.app.n8n.cloud/webhook/barber-style',
  barberoEmail: 'sepulvedayeison98@gmail.com',

  /* ── PREFIJOS (únicos por cliente, evitan colisiones en localStorage) ── */
  prefix:     'bs',          // prefijo localStorage y códigos de cita
  codePrefix: 'BS-',         // prefijo visible en códigos de reserva (ej: BS-AB1C2D)

  /* ── PANEL ADMIN ───────────────────────────────────────── */
  panel: {
    user:       'barberstyle',
    pass:       'bs2025*',
    sessionKey: 'bs_admin_v2'
  },

  /* ── SERVICIOS ─────────────────────────────────────────── */
  services: [
    { id: 'barba-afeitada',    name: 'Barba o Afeitada',        price: 18000, duration: 30, chip: 'Popular', active: true },
    { id: 'corte-solo',        name: 'Corte solo',              price: 25000, duration: 30, chip: '',        active: true },
    { id: 'corte-afeitada',    name: 'Corte + Afeitada',        price: 38000, duration: 50, chip: '',        active: true },
    { id: 'corte-barba',       name: 'Corte + Barba',           price: 42000, duration: 60, chip: '',        active: true },
    { id: 'corte-barba-cejas', name: 'Corte + Barba + Cejas',   price: 44000, duration: 60, chip: '',        active: true },
    { id: 'corte-cejas',       name: 'Corte + Cejas',           price: 30000, duration: 40, chip: '',        active: true },
    { id: 'cejas',             name: 'Marcación de cejas',      price:  8000, duration: 10, chip: '',        active: true },
    { id: 'contornos',         name: 'Marcación de contornos',  price: 10000, duration: 10, chip: '',        active: true }
  ],

  /* ── LO QUE OFRECE EL NEGOCIO (sección hero/about) ────── */
  offers: [
    { title: 'Cortes Modernos',    description: 'Tendencias actuales y asesoría personalizada' },
    { title: 'Diseño de Barba',    description: 'Perfilado profesional de alta precisión' },
    { title: 'Cejas Perfectas',    description: 'Definición y marcación de contornos' },
    { title: 'Ambiente Premium',   description: 'Experiencia exclusiva en cada visita' }
  ],

  /* ── EMPLEADOS ─────────────────────────────────────────── */
  employees: [
    {
      id:       'barbero-1',
      name:     'Barbero',
      role:     'Status Studio',
      photo:    '',
      services: [] // vacío = todos los servicios
    }
  ]
};
