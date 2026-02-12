# Suites - Portal Unificado

## Estructura del Proyecto

suites/
├── index.html      ← Portal principal (login + navegación)
├── Code.gs         ← Backend Google Apps Script
└── README.md       ← Este archivo


## Configuración Inicial

### 1. Crear Sheet "Auth"
Columnas necesarias:
- `usuario` (texto)
- `password` (MD5 hash)
- `rol` (admin, gerente, operador)

Para generar hash MD5 de una contraseña temporal:
```javascript
=LOWER(ARRAYFORMULA(JOIN("", DEC2HEX(CODE(MID(MD5("tupass"), ROW(INDIRECT("1:"&LEN(MD5("tupass")))), 1)), 2))))
