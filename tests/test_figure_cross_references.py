from docx import Document
from docx.oxml.ns import qn

from ezsnake.word_template_writer.api import reemplazar_referencias_cruzadas_de_figuras


def test_figure_reference_at_paragraph_end_is_inserted():
    document = Document()
    paragraph = document.add_paragraph("Antes de la <<reffigura_demo>>")

    reemplazar_referencias_cruzadas_de_figuras(
        document,
        {"<<reffigura_demo>>": ["RefFigura_demo"]},
    )

    field_instructions = [
        element.text
        for element in paragraph._p.iter(qn("w:instrText"))
    ]

    assert "<<reffigura_demo>>" not in paragraph.text
    assert " REF RefFigura_demo \\h " in field_instructions