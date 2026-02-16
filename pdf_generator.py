"""
Módulo para generación de PDFs personalizados dinámicamente.
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO
from datetime import datetime
import os
from typing import Dict, Optional


class PDFGenerator:
    """Clase para generar PDFs personalizados."""
    
    def __init__(self, logo_path: Optional[str] = None):
        """
        Inicializa el generador de PDFs.
        
        Args:
            logo_path: Ruta opcional al logo a incluir en el PDF
        """
        self.logo_path = logo_path
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
    
    def _setup_custom_styles(self):
        """Configura estilos personalizados para el PDF."""
        # Estilo para título
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Estilo para subtítulos
        self.styles.add(ParagraphStyle(
            name='CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#34495e'),
            spaceAfter=12,
            fontName='Helvetica-Bold'
        ))
        
        # Estilo para contenido
        self.styles.add(ParagraphStyle(
            name='CustomBody',
            parent=self.styles['BodyText'],
            fontSize=11,
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=12,
            alignment=TA_LEFT
        ))
    
    def generate_personalized_pdf(self, data: Dict[str, str]) -> bytes:
        """
        Genera un PDF personalizado con los datos proporcionados.
        
        Args:
            data: Diccionario con los datos personalizados (nombre, empresa, etc.)
        
        Returns:
            bytes: Contenido del PDF en bytes
        """
        buffer = BytesIO()
        
        # Crear documento PDF
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # Contenedor para los elementos del PDF
        elements = []
        
        # Agregar logo si existe
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                logo = Image(self.logo_path, width=2*inch, height=1*inch)
                logo.hAlign = 'CENTER'
                elements.append(logo)
                elements.append(Spacer(1, 0.3*inch))
            except Exception as e:
                print(f"Advertencia: No se pudo cargar el logo: {e}")
        
        # Título principal
        title = data.get('nombre_pdf', 'Documento Personalizado')
        title_para = Paragraph(title, self.styles['CustomTitle'])
        elements.append(title_para)
        elements.append(Spacer(1, 0.3*inch))
        
        # Información personalizada
        info_data = [
            ['Nombre:', data.get('nombre', 'N/A')],
            ['Empresa:', data.get('empresa', 'N/A')],
            ['Ciudad:', data.get('ciudad', 'N/A')],
            ['Fecha:', datetime.now().strftime('%d/%m/%Y')],
        ]
        
        # Crear tabla con información
        info_table = Table(info_data, colWidths=[2*inch, 4*inch])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#ecf0f1')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#2c3e50')),
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#bdc3c7'))
        ]))
        elements.append(info_table)
        elements.append(Spacer(1, 0.5*inch))
        
        # Mensaje personalizado si existe
        mensaje = data.get('mensaje_personalizado', '')
        if mensaje:
            heading = Paragraph('Mensaje Personalizado', self.styles['CustomHeading'])
            elements.append(heading)
            
            body = Paragraph(mensaje, self.styles['CustomBody'])
            elements.append(body)
            elements.append(Spacer(1, 0.3*inch))
        
        # Pie de página con información adicional
        footer_text = f"""
        <para align=center>
        <font size=9 color="#7f8c8d">
        Este documento ha sido generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}<br/>
        Documento confidencial - Para uso exclusivo del destinatario
        </font>
        </para>
        """
        footer = Paragraph(footer_text, self.styles['Normal'])
        elements.append(Spacer(1, 0.5*inch))
        elements.append(footer)
        
        # Construir PDF
        doc.build(elements)
        
        # Obtener bytes del PDF
        pdf_bytes = buffer.getvalue()
        buffer.close()
        
        return pdf_bytes


if __name__ == "__main__":
    # Ejemplo de uso
    generator = PDFGenerator()
    
    sample_data = {
        'nombre': 'Juan Pérez',
        'empresa': 'Empresa Ejemplo S.A.',
        'ciudad': 'Madrid',
        'mensaje_personalizado': 'Este es un mensaje de prueba para validar la generación de PDFs.',
        'nombre_pdf': 'Documento de Prueba'
    }
    
    pdf_content = generator.generate_personalized_pdf(sample_data)
    
    # Guardar PDF de prueba
    with open('/tmp/test_pdf.pdf', 'wb') as f:
        f.write(pdf_content)
    
    print(f"PDF de prueba generado: /tmp/test_pdf.pdf ({len(pdf_content)} bytes)")
