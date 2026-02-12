# Arquitectura del Portal Unificado

## Diagrama General

┌─────────────────────────────────────────┐
│           PORTAL CENTRAL                │
│  (Login único, roles, menú personalizado)│
├─────────────────────────────────────────┤
│  Módulo Operaciones    │  Módulo RRHH   │
│  - Dashboard KPIs      │  - Fichas      │
│  - Solicitudes turnos  │  - Aprobaciones│
│  - Mi rendimiento      │  - Reportes    │
├─────────────────────────────────────────┤
│      API UNIFICADA (GAS o Firebase)     │
│  - Auth centralizada                    │
│  - Logs de quién hizo qué               │
│  - Notificaciones entre módulos         │
└─────────────────────────────────────────┘

## Niveles de Seguridad Propuestos

### Nivel 1: "Oficina" (MVP)
- Login con usuario/pass en Sheet "Auth"
- Links de GAS expuestos pero ocultos detrás de portal
- Riesgo: Links pueden compartirse

### Nivel 2: "Empresa" (Recomendado)
- Portal + Proxy serverless (Cloudflare Workers)
- URLs de GAS firmadas temporalmente (expiran en 5 min)
- Links no reutilizables

### Nivel 3: "Enterprise" (Futuro)
- Migración completa a Firebase/Supabase
- Auth profesional, BD real, sin Google Sheets

## Roles de Usuario

| Rol | Paneles | Permisos |
|-----|---------|----------|
| Admin | Todos + config | CRUD completo |
| Gerente | Dashboards consolidados | Ver todos los proyectos |
| Jefe de Proyecto | Solo su proyecto | Ver/editar asignados |
| Operador | 1-2 pantallas de trabajo | Solo sus datos |

## Stack Tecnológico Propuesto

**Frontend:** HTML/CSS/JS vanilla (como los actuales)  
**Hosting:** GitHub Pages / Netlify / Vercel (gratis)  
**Backend temporal:** Google Apps Script (lo que ya tienes)  
**Auth:** Sheet propia (Nivel 1) → Cloudflare Workers (Nivel 2)  
**BD futura:** Firebase o Supabase (cuando aprueben presupuesto)
