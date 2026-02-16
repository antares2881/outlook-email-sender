# 🚀 Guía Rápida de Inicio

## Setup Rápido (5 minutos)

### 1. Instalar Dependencias
```bash
pip install -r requirements.txt
```

### 2. Configurar Credenciales
```bash
# Copiar plantilla
cp .env.example .env

# Editar .env y agregar tus credenciales
nano .env  # o usa tu editor favorito
```

Contenido del `.env`:
```
OUTLOOK_EMAIL=tu_email@empresa.com
OUTLOOK_PASSWORD=tu_contraseña_de_aplicacion
```

### 3. Personalizar Configuración (Opcional)
Editar `config.json` para cambiar:
- Nombre del remitente
- Asunto del email
- Delay entre envíos
- Rutas de archivos

### 4. Preparar tus Datos
Editar `data/destinatarios_ejemplo.xlsx` con tus destinatarios reales.

**Columnas requeridas:**
- `email` (requerido)
- `nombre` (requerido)
- `empresa`, `ciudad`, `mensaje_personalizado`, `nombre_pdf` (opcionales)

### 5. Probar el Sistema
```bash
# Enviar email de prueba a ti mismo
python email_sender.py --preview tu_email@ejemplo.com
```

### 6. Envío Real
```bash
# Modo interactivo (recomendado)
python email_sender.py

# O envío directo
python email_sender.py --send
```

## 📋 Comandos Principales

```bash
# Ver ayuda
python email_sender.py --help

# Modo interactivo (menú)
python email_sender.py

# Envío directo (pide confirmación)
python email_sender.py --send

# Email de prueba
python email_sender.py --preview email@ejemplo.com

# Usar configuración personalizada
python email_sender.py --config mi_config.json
```

## ⚠️ Checklist Antes del Primer Envío

- [ ] Credenciales configuradas en `.env`
- [ ] Contraseña de aplicación (no contraseña normal)
- [ ] Excel actualizado con datos reales
- [ ] Emails validados (formato correcto)
- [ ] Plantilla HTML personalizada (opcional)
- [ ] Email de prueba enviado y recibido
- [ ] Verificar que PDF adjunto es correcto

## 🔧 Problemas Comunes

**Error de autenticación:**
→ Verifica que usas contraseña de aplicación, no contraseña normal

**Email no llega:**
→ Revisa carpeta de spam
→ Verifica configuración SMTP en `config.json`

**Columnas faltantes:**
→ Asegúrate que Excel tiene columnas `email` y `nombre`

## 📚 Más Información

Ver `README.md` para documentación completa.
