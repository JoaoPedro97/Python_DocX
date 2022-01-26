from docx import Document

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches, Cm


class App_Central:
    document = Document()
    def Preset_Margens():
        sections = document.sections
        for section in sections:
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(1.6)
            section.right_margin = Cm(1.6)

    def Estilos():

        styles = document.styles

        #Fonte Padrão
        Font_Geral = styles.add_style("PadraoFont", WD_STYLE_TYPE.PARAGRAPH)
        Font_Geral.font.nome = "Arial"
        Font_Geral.font.size = Pt(10)
        Font_Geral.font.bold = False

        #para regioções sem texto
        Estilo_Null = styles.add_style("NULL", WD_STYLE_TYPE.PARAGRAPH)
        Estilo_Null.font.nome = "Arial"
        Estilo_Null.font.size = Pt(1)
        Estilo_Null.font.bold = False

       