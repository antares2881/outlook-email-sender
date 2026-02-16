"""
Script principal para envío masivo de emails personalizados con Outlook.
"""

import smtplib
import json
import os
import sys
import logging
import time
import argparse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from tqdm import tqdm

from pdf_generator import PDFGenerator


class EmailSender:
    """Clase principal para manejo de envío de emails masivos."""
    
    def __init__(self, config_path: str = "config.json"):
        """
        Inicializa el EmailSender con configuración.
        
        Args:
            config_path: Ruta al archivo de configuración JSON
        """
        # Cargar variables de entorno
        load_dotenv()
        
        # Configurar logging
        self._setup_logging()
        
        # Cargar configuración
        self.config = self._load_config(config_path)
        
        # Validar credenciales
        self.email = os.getenv('OUTLOOK_EMAIL')
        self.password = os.getenv('OUTLOOK_PASSWORD')
        
        if not self.email or not self.password:
            raise ValueError(
                "❌ Credenciales no configuradas. "
                "Asegúrate de crear un archivo .env con OUTLOOK_EMAIL y OUTLOOK_PASSWORD"
            )
        
        # Inicializar generador de PDFs
        logo_path = self.config['files'].get('logo_path')
        if logo_path and os.path.exists(logo_path):
            self.pdf_generator = PDFGenerator(logo_path)
        else:
            self.pdf_generator = PDFGenerator()
        
        # Cargar plantilla de email
        self.email_template = self._load_email_template()
        
        # Lista para tracking de envíos
        self.results = []
        
        self.logger.info("EmailSender inicializado correctamente")
    
    def _setup_logging(self):
        """Configura el sistema de logging."""
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = log_dir / f'email_sender_{timestamp}.log'
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Logging inicializado: {log_file}")
    
    def _load_config(self, config_path: str) -> dict:
        """Carga el archivo de configuración JSON."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.logger.info(f"Configuración cargada desde {config_path}")
            return config
        except FileNotFoundError:
            self.logger.error(f"Archivo de configuración no encontrado: {config_path}")
            raise
        except json.JSONDecodeError as e:
            self.logger.error(f"Error al parsear JSON: {e}")
            raise
    
    def _load_email_template(self) -> str:
        """Carga la plantilla HTML del email."""
        template_path = self.config['files']['email_template']
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template = f.read()
            self.logger.info(f"Plantilla de email cargada desde {template_path}")
            return template
        except FileNotFoundError:
            self.logger.warning(f"Plantilla no encontrada: {template_path}. Usando plantilla básica.")
            return """
            <html>
                <body>
                    <h2>Hola {{nombre}},</h2>
                    <p>{{mensaje_personalizado}}</p>
                    <p>Saludos cordiales,<br>{{from_name}}</p>
                </body>
            </html>
            """
    
    def load_excel(self, excel_path: Optional[str] = None) -> pd.DataFrame:
        """
        Carga y valida el archivo Excel con destinatarios.
        
        Args:
            excel_path: Ruta al archivo Excel (usa config si no se especifica)
        
        Returns:
            DataFrame con los datos validados
        """
        if excel_path is None:
            excel_path = self.config['files']['excel_path']
        
        self.logger.info(f"Cargando Excel: {excel_path}")
        
        try:
            df = pd.read_excel(excel_path)
        except FileNotFoundError:
            self.logger.error(f"Archivo Excel no encontrado: {excel_path}")
            raise
        except Exception as e:
            self.logger.error(f"Error al leer Excel: {e}")
            raise
        
        # Validar columnas requeridas
        required_columns = ['email', 'nombre']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"❌ Columnas faltantes en Excel: {missing_columns}")
        
        # Validar formato de emails con regex más completo
        # Acepta emails con TLDs largos y caracteres internacionales básicos
        email_pattern = r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
        invalid_emails = df[~df['email'].str.match(email_pattern, na=False)]
        
        if not invalid_emails.empty:
            self.logger.warning(f"Emails inválidos encontrados: {invalid_emails['email'].tolist()}")
            df = df[df['email'].str.match(email_pattern, na=False)]
        
        self.logger.info(f"✅ Excel cargado: {len(df)} destinatarios válidos")
        return df
    
    def _render_template(self, template: str, data: Dict[str, str]) -> str:
        """
        Reemplaza variables en la plantilla con datos personalizados.
        
        Args:
            template: Plantilla con variables {{variable}}
            data: Diccionario con valores
        
        Returns:
            Plantilla renderizada
        """
        rendered = template
        
        # Agregar from_name a los datos
        data['from_name'] = self.config['email']['from_name']
        
        # Reemplazar variables
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            rendered = rendered.replace(placeholder, str(value))
        
        return rendered
    
    def send_email(
        self,
        to_email: str,
        subject: str,
        html_body: str,
        pdf_data: Optional[bytes] = None,
        pdf_filename: str = "documento.pdf"
    ) -> Tuple[bool, Optional[str]]:
        """
        Envía un email individual.
        
        Args:
            to_email: Email del destinatario
            subject: Asunto del email
            html_body: Contenido HTML del email
            pdf_data: Datos del PDF adjunto (opcional)
            pdf_filename: Nombre del archivo PDF
        
        Returns:
            Tupla (éxito: bool, error: Optional[str])
        """
        try:
            # Crear mensaje
            msg = MIMEMultipart('alternative')
            msg['From'] = f"{self.config['email']['from_name']} <{self.email}>"
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # Adjuntar HTML
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg.attach(html_part)
            
            # Adjuntar PDF si existe
            if pdf_data:
                pdf_part = MIMEApplication(pdf_data, _subtype='pdf')
                pdf_part.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
                msg.attach(pdf_part)
            
            # Conectar a servidor SMTP
            smtp_config = self.config['smtp']
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
            
            if smtp_config.get('use_tls', True):
                server.starttls()
            
            # Autenticación
            server.login(self.email, self.password)
            
            # Enviar email
            server.send_message(msg)
            server.quit()
            
            self.logger.info(f"✅ Email enviado a: {to_email}")
            return True, None
            
        except smtplib.SMTPAuthenticationError as e:
            error = f"Error de autenticación: {e}"
            self.logger.error(f"❌ {error} - {to_email}")
            return False, error
        except smtplib.SMTPException as e:
            error = f"Error SMTP: {e}"
            self.logger.error(f"❌ {error} - {to_email}")
            return False, error
        except Exception as e:
            error = f"Error inesperado: {e}"
            self.logger.error(f"❌ {error} - {to_email}")
            return False, error
    
    def send_bulk_emails(self, df: pd.DataFrame, preview_mode: bool = False) -> Dict[str, int]:
        """
        Envía emails masivos a todos los destinatarios.
        
        Args:
            df: DataFrame con destinatarios
            preview_mode: Si True, solo envía a los primeros registros
        
        Returns:
            Diccionario con estadísticas de envío
        """
        if preview_mode:
            df = df.head(1)
            print("\n⚠️  MODO PREVIEW ACTIVADO: Solo se enviará 1 email\n")
        
        total = len(df)
        success_count = 0
        failed_count = 0
        
        print(f"\n📧 Iniciando envío de {total} emails...\n")
        
        # Barra de progreso
        for idx, row in tqdm(df.iterrows(), total=total, desc="Enviando emails"):
            # Convertir row a dict y manejar valores NaN
            data = row.to_dict()
            data = {k: (v if pd.notna(v) else '') for k, v in data.items()}
            
            # Generar PDF personalizado
            try:
                pdf_data = self.pdf_generator.generate_personalized_pdf(data)
                # Sanitizar nombre de archivo removiendo caracteres problemáticos
                safe_pdf_name = data.get('nombre_pdf', 'documento').replace('/', '_').replace('\\', '_').replace(':', '_')
                safe_nombre = data['nombre'].replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '_')
                pdf_filename = f"{safe_pdf_name}_{safe_nombre}.pdf"
            except Exception as e:
                self.logger.error(f"Error generando PDF para {data['email']}: {e}")
                pdf_data = None
                pdf_filename = "documento.pdf"
            
            # Renderizar plantilla de email
            html_body = self._render_template(self.email_template, data)
            subject = self._render_template(self.config['email']['subject'], data)
            
            # Intentar enviar email
            max_retries = self.config['settings']['max_retries']
            success = False
            error_msg = None
            
            for attempt in range(max_retries):
                success, error_msg = self.send_email(
                    to_email=data['email'],
                    subject=subject,
                    html_body=html_body,
                    pdf_data=pdf_data,
                    pdf_filename=pdf_filename
                )
                
                if success:
                    break
                
                if attempt < max_retries - 1:
                    time.sleep(2)
            
            # Registrar resultado
            result = {
                'email': data['email'],
                'nombre': data['nombre'],
                'status': 'Éxito' if success else 'Error',
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'error': error_msg if not success else ''
            }
            self.results.append(result)
            
            if success:
                success_count += 1
            else:
                failed_count += 1
            
            # Delay entre envíos
            if idx < total - 1:  # No delay después del último
                delay = self.config['settings']['delay_between_emails']
                time.sleep(delay)
        
        stats = {
            'total': total,
            'success': success_count,
            'failed': failed_count
        }
        
        print(f"\n✅ Envío completado!")
        print(f"   Total: {stats['total']}")
        print(f"   Éxitos: {stats['success']}")
        print(f"   Errores: {stats['failed']}\n")
        
        return stats
    
    def generate_report(self) -> str:
        """
        Genera reporte CSV con resultados de envío.
        
        Returns:
            Ruta al archivo de reporte
        """
        if not self.results:
            self.logger.warning("No hay resultados para generar reporte")
            return ""
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = Path('logs') / f'reporte_envios_{timestamp}.csv'
        
        df_results = pd.DataFrame(self.results)
        df_results.to_csv(report_path, index=False, encoding='utf-8')
        
        self.logger.info(f"📊 Reporte generado: {report_path}")
        print(f"📊 Reporte guardado en: {report_path}")
        
        return report_path


def main():
    """Función principal con menú interactivo."""
    parser = argparse.ArgumentParser(
        description='Sistema de envío masivo de emails con Outlook'
    )
    parser.add_argument(
        '--send',
        action='store_true',
        help='Enviar emails directamente sin menú interactivo'
    )
    parser.add_argument(
        '--preview',
        type=str,
        metavar='EMAIL',
        help='Enviar email de prueba a la dirección especificada'
    )
    parser.add_argument(
        '--config',
        type=str,
        default='config.json',
        help='Ruta al archivo de configuración (default: config.json)'
    )
    
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("📧  SISTEMA DE ENVÍO MASIVO DE EMAILS CON OUTLOOK")
    print("="*60 + "\n")
    
    try:
        # Inicializar EmailSender
        sender = EmailSender(args.config)
        
        # Modo preview
        if args.preview:
            print(f"🔍 Modo Preview: Enviando email de prueba a {args.preview}\n")
            
            # Crear datos de prueba
            test_data = {
                'email': args.preview,
                'nombre': 'Usuario de Prueba',
                'empresa': 'Empresa de Prueba',
                'ciudad': 'Ciudad de Prueba',
                'mensaje_personalizado': 'Este es un email de prueba del sistema.',
                'nombre_pdf': 'Documento de Prueba'
            }
            
            # Generar PDF
            pdf_data = sender.pdf_generator.generate_personalized_pdf(test_data)
            
            # Renderizar email
            html_body = sender._render_template(sender.email_template, test_data)
            subject = sender._render_template(sender.config['email']['subject'], test_data)
            
            # Enviar
            success, error = sender.send_email(
                to_email=args.preview,
                subject=subject,
                html_body=html_body,
                pdf_data=pdf_data,
                pdf_filename="prueba.pdf"
            )
            
            if success:
                print("✅ Email de prueba enviado exitosamente")
            else:
                print(f"❌ Error al enviar email de prueba: {error}")
            
            return
        
        # Cargar Excel
        df = sender.load_excel()
        
        # Modo envío directo
        if args.send:
            print("\n⚠️  ¿Estás seguro de enviar emails a todos los destinatarios?")
            confirm = input(f"Se enviarán {len(df)} emails. Escribe 'SI' para confirmar: ")
            
            if confirm.upper() == 'SI':
                stats = sender.send_bulk_emails(df)
                sender.generate_report()
            else:
                print("❌ Envío cancelado")
            
            return
        
        # Menú interactivo
        while True:
            print("\n" + "-"*60)
            print("MENÚ PRINCIPAL")
            print("-"*60)
            print(f"📊 Destinatarios cargados: {len(df)}")
            print("\nOpciones:")
            print("  1. Enviar emails a todos los destinatarios")
            print("  2. Modo preview (enviar solo al primer destinatario)")
            print("  3. Ver estadísticas de destinatarios")
            print("  4. Recargar archivo Excel")
            print("  5. Salir")
            print("-"*60)
            
            opcion = input("\nSelecciona una opción (1-5): ").strip()
            
            if opcion == '1':
                print("\n⚠️  ¿Estás seguro de enviar emails a TODOS los destinatarios?")
                confirm = input(f"Se enviarán {len(df)} emails. Escribe 'SI' para confirmar: ")
                
                if confirm.upper() == 'SI':
                    stats = sender.send_bulk_emails(df)
                    sender.generate_report()
                else:
                    print("❌ Envío cancelado")
            
            elif opcion == '2':
                stats = sender.send_bulk_emails(df, preview_mode=True)
                sender.generate_report()
            
            elif opcion == '3':
                print("\n📊 ESTADÍSTICAS DE DESTINATARIOS")
                print(f"   Total: {len(df)}")
                print(f"   Columnas: {', '.join(df.columns.tolist())}")
                print(f"\n   Primeros 5 registros:")
                print(df.head().to_string(index=False))
            
            elif opcion == '4':
                df = sender.load_excel()
                print(f"✅ Excel recargado: {len(df)} destinatarios")
            
            elif opcion == '5':
                print("\n👋 ¡Hasta luego!")
                break
            
            else:
                print("❌ Opción inválida. Por favor selecciona 1-5.")
    
    except KeyboardInterrupt:
        print("\n\n⚠️  Proceso interrumpido por el usuario")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Error fatal: {e}")
        logging.exception("Error fatal en main()")
        sys.exit(1)


if __name__ == "__main__":
    main()
