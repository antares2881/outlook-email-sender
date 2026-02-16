# 📧 Sistema de Envío Masivo de Emails con Outlook

Sistema completo en Python para envío masivo de correos electrónicos personalizados desde una cuenta corporativa de Outlook/Office365, con generación dinámica de archivos PDF adjuntos para cada destinatario.

## 📋 Descripción

Esta aplicación permite:
- ✉️ Envío masivo de emails personalizados desde Outlook corporativo
- 📄 Generación automática de PDFs únicos para cada destinatario
- 📊 Lectura de datos desde archivos Excel
- 🎨 Plantillas HTML personalizables para emails
- 📝 Sistema completo de logging y reportes
- 🔒 Manejo seguro de credenciales
- ⚡ Gestión de errores robusta con reintentos
- 🎯 Modo preview para pruebas

## ⚙️ Requisitos Previos

- **Python 3.8 o superior**
- **Cuenta corporativa de Outlook/Microsoft 365**
- **Contraseña de aplicación configurada** (ver sección de configuración)

## 🚀 Instalación Paso a Paso

### 1. Clonar el Repositorio

```bash
git clone https://github.com/antares2881/outlook-email-sender.git
cd outlook-email-sender
```

### 2. Crear Entorno Virtual (Recomendado)

```bash
# En Windows
python -m venv venv
venv\Scripts\activate

# En Linux/Mac
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar Dependencias

```bash
pip install -r requirements.txt
```

### 4. Configurar Variables de Entorno

Crea un archivo `.env` en la raíz del proyecto (puedes copiar desde `.env.example`):

```bash
cp .env.example .env
```

Edita el archivo `.env` con tus credenciales:

```env
OUTLOOK_EMAIL=tu_email@empresa.com
OUTLOOK_PASSWORD=tu_contraseña_de_aplicacion
```

⚠️ **IMPORTANTE**: Nunca compartas tu archivo `.env` ni lo subas a repositorios públicos.

### 5. Personalizar Configuración

Edita el archivo `config.json` según tus necesidades:

```json
{
  "smtp": {
    "server": "smtp-mail.outlook.com",
    "port": 587,
    "use_tls": true
  },
  "email": {
    "from_name": "Tu Nombre o Empresa",
    "subject": "Asunto del Email"
  },
  "settings": {
    "delay_between_emails": 2,
    "max_retries": 2,
    "preview_mode": false
  },
  "files": {
    "excel_path": "data/destinatarios_ejemplo.xlsx",
    "email_template": "templates/email_template.html",
    "logo_path": "data/logo.png"
  }
}
```

## 📧 Configuración de Outlook

### Obtener Contraseña de Aplicación

Microsoft requiere contraseñas de aplicación para acceder a Outlook vía SMTP:

1. **Accede a tu cuenta Microsoft**: https://account.microsoft.com/security
2. **Habilita la verificación en dos pasos** si no la tienes activada
3. **Ve a "Contraseñas de aplicación"**
4. **Crea una nueva contraseña** con nombre descriptivo (ej: "EmailSender")
5. **Copia la contraseña generada** y úsala en tu archivo `.env`

### Configuración para Office 365 Corporativo

Si tu empresa usa Office 365, puede que necesites:

- **Autenticación moderna habilitada** en tu organización
- **Permisos SMTP** activados por el administrador
- **MFA configurado** para contraseñas de aplicación

💡 **Consejo**: Si tienes problemas de autenticación, contacta con tu departamento de IT.

## 📊 Preparar tu Archivo Excel

### Estructura Requerida

Tu archivo Excel debe contener las siguientes columnas:

| Columna | Descripción | Requerido |
|---------|-------------|-----------|
| `email` | Email del destinatario | ✅ Sí |
| `nombre` | Nombre del destinatario | ✅ Sí |
| `empresa` | Nombre de la empresa | ⭕ Opcional |
| `ciudad` | Ciudad | ⭕ Opcional |
| `mensaje_personalizado` | Mensaje único para cada destinatario | ⭕ Opcional |
| `nombre_pdf` | Nombre del documento PDF | ⭕ Opcional |

### Ejemplo de Datos

```
email                         | nombre           | empresa                    | ciudad    | mensaje_personalizado
------------------------------|------------------|----------------------------|-----------|---------------------
juan.perez@ejemplo.com       | Juan Pérez       | TechCorp Solutions         | Madrid    | Mensaje para Juan...
maria.gonzalez@ejemplo.com   | María González   | Innovación Digital         | Barcelona | Mensaje para María...
```

📁 Puedes usar el archivo de ejemplo incluido: `data/destinatarios_ejemplo.xlsx`

### Validaciones Automáticas

El sistema valida:
- ✅ Formato correcto de emails
- ✅ Columnas requeridas presentes
- ✅ Datos no vacíos en campos obligatorios

## 🎨 Personalizar Plantillas

### Plantilla de Email

Edita `templates/email_template.html` para personalizar el diseño del email.

**Variables disponibles:**

- `{{nombre}}` - Nombre del destinatario
- `{{empresa}}` - Empresa del destinatario
- `{{ciudad}}` - Ciudad del destinatario
- `{{mensaje_personalizado}}` - Mensaje personalizado
- `{{from_name}}` - Nombre del remitente (desde config.json)

**Ejemplo de uso en HTML:**

```html
<h2>Hola {{nombre}},</h2>
<p>{{mensaje_personalizado}}</p>
<p>Empresa: {{empresa}}</p>
```

### Plantilla de PDF

El PDF se genera automáticamente con:
- 📋 Datos personalizados en tabla
- 🖼️ Logo opcional (si existe `data/logo.png`)
- 📅 Fecha de generación
- ✍️ Mensaje personalizado
- 🎨 Diseño profesional predefinido

Para personalizar el PDF, edita `pdf_generator.py`.

## ▶️ Uso de la Aplicación

### Modo Interactivo (Recomendado)

```bash
python email_sender.py
```

Se mostrará un menú con opciones:
1. Enviar emails a todos los destinatarios
2. Modo preview (solo primer destinatario)
3. Ver estadísticas de destinatarios
4. Recargar archivo Excel
5. Salir

### Envío Directo

```bash
python email_sender.py --send
```

⚠️ Se pedirá confirmación antes de enviar.

### Modo Preview/Prueba

Enviar email de prueba a una dirección específica:

```bash
python email_sender.py --preview tu_email@ejemplo.com
```

### Usar Archivo de Configuración Personalizado

```bash
python email_sender.py --config mi_config.json
```

## 📝 Ejemplos de Uso

### Caso 1: Primera Prueba

```bash
# 1. Verifica tu configuración
cat .env

# 2. Envía un email de prueba a ti mismo
python email_sender.py --preview tu_email@empresa.com

# 3. Verifica que el email y PDF se recibieron correctamente
```

### Caso 2: Envío a Pequeño Grupo

```bash
# 1. Inicia modo interactivo
python email_sender.py

# 2. Selecciona opción 2 (Modo preview) para enviar solo al primero
# 3. Verifica el resultado
# 4. Si todo está bien, selecciona opción 1 para envío completo
```

### Caso 3: Envío Masivo Programado

```bash
# Crear script de envío
python email_sender.py --send
```

## 🔍 Logs y Reportes

### Archivos de Log

Los logs se guardan en `logs/` con formato:
```
logs/email_sender_YYYYMMDD_HHMMSS.log
```

Contienen información detallada de:
- ✅ Emails enviados exitosamente
- ❌ Errores con descripción detallada
- ⚙️ Operaciones del sistema

### Reportes CSV

Después de cada envío se genera un reporte:
```
logs/reporte_envios_YYYYMMDD_HHMMSS.csv
```

Con las siguientes columnas:
- `email` - Email del destinatario
- `nombre` - Nombre del destinatario
- `status` - Éxito o Error
- `timestamp` - Fecha y hora del envío
- `error` - Descripción del error (si aplica)

## ⚠️ Solución de Problemas

### Error: "Autenticación fallida"

**Causa**: Credenciales incorrectas o contraseña de aplicación no configurada.

**Solución**:
1. Verifica que estés usando una contraseña de aplicación, no tu contraseña normal
2. Regenera la contraseña de aplicación en Microsoft
3. Verifica que no haya espacios en el archivo `.env`

### Error: "SMTP timeout" o "Connection refused"

**Causa**: Configuración SMTP incorrecta o firewall bloqueando.

**Solución**:
1. Verifica servidor: `smtp-mail.outlook.com` puerto `587`
2. Comprueba tu conexión a internet
3. Verifica que tu firewall permite conexiones SMTP

### Error: "Columnas faltantes en Excel"

**Causa**: El archivo Excel no tiene las columnas requeridas.

**Solución**:
1. Asegúrate de que existan columnas `email` y `nombre`
2. Verifica que los nombres estén escritos exactamente igual
3. Usa el archivo de ejemplo como referencia

### Error: "Email inválido"

**Causa**: Formato de email incorrecto en el Excel.

**Solución**:
1. Revisa que todos los emails tengan formato `usuario@dominio.com`
2. El sistema filtrará automáticamente emails inválidos
3. Revisa los logs para ver qué emails fueron filtrados

### Emails marcados como Spam

**Causa**: Límites de Outlook o contenido sospechoso.

**Solución**:
1. Aumenta el delay entre envíos en `config.json`
2. No envíes más de 1000 emails/hora
3. Evita palabras "spam" en el asunto
4. Pide a destinatarios que agreguen tu email a contactos

## 🔐 Mejores Prácticas de Seguridad

### ✅ Hacer

- ✅ Usar variables de entorno para credenciales
- ✅ Mantener `.env` en `.gitignore`
- ✅ Usar contraseñas de aplicación, no contraseñas principales
- ✅ Rotar contraseñas regularmente
- ✅ Limitar acceso al archivo `.env`
- ✅ Hacer copias de seguridad de logs importantes

### ❌ No Hacer

- ❌ Hardcodear credenciales en el código
- ❌ Compartir archivos `.env`
- ❌ Subir credenciales a Git
- ❌ Usar la misma contraseña para múltiples servicios
- ❌ Compartir contraseñas de aplicación

## 📁 Estructura del Proyecto

```
outlook-email-sender/
├── README.md                          # Este archivo
├── requirements.txt                   # Dependencias Python
├── .gitignore                        # Archivos ignorados por Git
├── .env.example                      # Plantilla de variables de entorno
├── .env                              # Variables de entorno (no incluir en Git)
├── config.json                       # Configuración del sistema
├── email_sender.py                   # Script principal
├── pdf_generator.py                  # Generador de PDFs
├── templates/
│   └── email_template.html          # Plantilla HTML de emails
├── data/
│   ├── destinatarios_ejemplo.xlsx   # Ejemplo de archivo Excel
│   └── logo.png                     # Logo opcional para PDFs
├── logs/
│   ├── .gitkeep
│   ├── email_sender_*.log           # Logs del sistema
│   └── reporte_envios_*.csv         # Reportes de envíos
└── outputs/
    └── .gitkeep                      # Carpeta para archivos temporales
```

## 🚦 Límites de Outlook

Ten en cuenta los límites de Microsoft/Outlook:

- **Cuentas corporativas**: ~10,000 emails/día
- **Cuentas personales**: ~300 emails/día
- **Tasa recomendada**: 1-2 emails/segundo
- **Tamaño de adjuntos**: Máximo 25 MB

💡 El sistema incluye delays automáticos configurables para respetar estos límites.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver archivo `LICENSE` para más detalles.

## 📞 Soporte

Si encuentras algún problema o tienes sugerencias:

1. Revisa la sección "Solución de Problemas"
2. Consulta los logs en la carpeta `logs/`
3. Abre un issue en GitHub con detalles del problema

## 🎯 Roadmap

Características planeadas para futuras versiones:

- [ ] Soporte para imágenes embebidas en emails
- [ ] Plantillas de PDF múltiples
- [ ] Interfaz web
- [ ] Programación de envíos
- [ ] Soporte para otros proveedores SMTP
- [ ] Dashboard de estadísticas
- [ ] Integración con APIs de marketing

## 📚 Recursos Adicionales

- [Documentación oficial de smtplib](https://docs.python.org/3/library/smtplib.html)
- [Guía de ReportLab para PDFs](https://www.reportlab.com/docs/reportlab-userguide.pdf)
- [Configuración SMTP de Outlook](https://support.microsoft.com/en-us/office/pop-imap-and-smtp-settings-8361e398-8af4-4e97-b147-6c6c4ac95353)

---

⭐ Si este proyecto te ha sido útil, considera darle una estrella en GitHub.

Desarrollado con ❤️ para facilitar las comunicaciones empresariales.
