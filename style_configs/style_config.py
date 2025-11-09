from docx.shared import Pt, Cm
from docx.document import Document
from docx import Document as Doc
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from dataclasses import dataclass
from typing import Any

class StyleNames:
    h1 = "ГОСТ_заголовок1"
    h2 = "ГОСТ_заголовок2"
    h3 = "ГОСТ_заголовок3"
    list = "ГОСТ_список"
    normal = "ГОСТ_обычный"
    tb_item = "ГОСТ_текст_в_таблице"
    caption = "ГОСТ_подпись"
    code = "ГОСТ_код"
    formula = "ГОСТ_формула"

    def conf_style_names(self):
        pass

@dataclass
class StyleConf:
    font_name: str = "Times New Roman"
    font_size: Any = Pt(14)
    font_bold: bool = False
    font_italic: bool = False
    font_all_caps: bool = False
    alignment: Any = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    line_spacing: float = 1.25
    space_before: Any = Pt(0)
    space_after: Any = Pt(0)
    page_break_before: bool = False
    first_line_indent: Any = Cm(0)
    left_indent: Any = Cm(0)
    hanging_indent: Any = Cm(0)

class StyleManager:
    def __init__(self):
        self.styles = self.conf_styles()
        self.tb_conf = self.conf_tables()
        self.page_conf = self.conf_page()

    @staticmethod
    def conf_styles():
        return {
            StyleNames.h1: StyleConf(
                font_size=Pt(16),
                font_bold=True,
                font_all_caps=True,
                alignment=WD_PARAGRAPH_ALIGNMENT.CENTER,
                space_before=Pt(12),
                space_after=Pt(6),
                page_break_before=True,
            ),
            StyleNames.h2: StyleConf(
                font_size=Pt(14),
                font_bold=True,
                alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
                space_before=Pt(6),
                space_after=Pt(6),
            ),
            StyleNames.h3: StyleConf(
                font_size=Pt(14),
                alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
                space_before=Pt(6),
                space_after=Pt(3),
            ),
            StyleNames.list: StyleConf(
                line_spacing=1.5,
                hanging_indent=Cm(0.5),
            ),
            StyleNames.normal: StyleConf(
                font_size=Pt(14),
                font_bold=False,
                font_italic=False,
                line_spacing=1.5,
                first_line_indent=Cm(1.25),
            ),
            StyleNames.tb_item: StyleConf(
                line_spacing=1.5,
            ),
            StyleNames.caption: StyleConf(
                line_spacing=1.5,
            ),
            StyleNames.code: StyleConf(
                font_name="Courier New",
                font_size=Pt(12),
                line_spacing=1.0,
            ),
            StyleNames.formula: StyleConf(
                font_italic=True,
                alignment=WD_PARAGRAPH_ALIGNMENT.CENTER,
                space_before=Pt(12),
                space_after=Pt(6),
                line_spacing=1.0,
            ),
        }

    @staticmethod
    def conf_tables():
        return {
            'border_size': 1,
            'border_style': 'single',
            'row_height': Cm(0.8),
            'cell_margins': {
                'left': Cm(0.2), 'right': Cm(0.2),
                'top': Cm(0.2), 'bottom': Cm(0.2)
            }
        }

    @staticmethod
    def conf_page():
        return {
            'page_number_font_size': Pt(12),
            'page_number_font_name': 'Times New Roman'
        }

    def setup_styles(self, doc: Document):
        for style_name, config in self.styles.items():
            self.apply_style(style_name, config, doc)

    def apply_style(self, name: str, conf: StyleConf, doc: Document):
        try:
            style = doc.styles.add_style(name, 1)
        except ValueError:
            style = doc.styles["name"]

        self.conf_style(style, conf)

    @staticmethod
    def conf_style(style, config: StyleConf) -> None:
        font = style.font
        font.name = config.font_name
        font.size = config.font_size
        font.bold = config.font_bold
        font.italic = config.font_italic
        font.all_caps = config.font_all_caps

        pf = style.paragraph_format
        if config.alignment:
            pf.alignment = config.alignment
        if config.line_spacing:
            pf.line_spacing = config.line_spacing
        if config.space_after:
            pf.space_after = config.space_after
        if config.space_before:
            pf.space_before = config.space_before
        if config.page_break_before:
            pf.page_break_before = config.page_break_before
        if config.first_line_indent:
            pf.first_line_indent = config.first_line_indent
        if config.left_indent:
            pf.left_indent = config.left_indent
        if config.hanging_indent:
            pf.hanging_indent = config.hanging_indent
