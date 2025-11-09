from dataclasses import dataclass
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import re
from typing import TypedDict, List

@dataclass
class ParseResult(TypedDict):
    level: int
    ptype: str
    par: Paragraph | None
    text: str
    index: int

class MdParser:
    def __init__(self, text):
        self.text = text

    @staticmethod
    def parse_paragraph(line) -> tuple[str, int | None, str]:
        if re.match(r'-{3,}', line):
            return 'page_break', 0, ''

        if re.match(r"^(>)+ .*", line):
            return 'normal', 0, re.sub(r"^(> )+", "", line)

        if re.match(r"^```.+", line):
            return 'normal', 0, re.sub(r"^(> )+", "", re.sub(r"^```", "", line))

        for i in range(1, 7):
            pat = f'^{"#" * i} (.+)$'
            if re.fullmatch(pat, line):
                return 'header', i, re.sub(f'{"#" * i} ', '', line.strip())

        if re.match(r'^ *[-*+] (.+)$', line):
            return 'unord_list', 0, re.sub(r'^[-*+] ', '', line.strip())

        if re.match(r'^ *\d+\. (.+)$', line):
            return 'ord_list', 0, line.strip()

        if '|' in line and line.count('|') >= 2:
            if re.match(r"^\|(-+)(?:\|(-+))+\|$", line):
                return 'table_del', None, line
            else:
                return 'table', None, line

        if re.match(r"^ *\S+ *", line):
            return 'normal', 0, line

        return 'empty', 0, line

    def parse_(self):
        parse_data = []
        pars = self.text.split('\n')

        for i, p in enumerate(pars):
            if not re.match("^<.+>", p):
                parse_data += [dict(zip(["type", "level", "text"], [*self.parse_paragraph(p)]))]

        return parse_data

class DocParser:
    def __init__(self, document):
        self.doc = document
        self.pars = document.paragraphs
        self.cur_section = 'content'

    # Определяет тип содержимого
    def determine_type(self, p: Paragraph) -> tuple[str, int]:
        return (self.det_heading(p) or self.det_caption(p)
                or self.det_img_caption(p) or self.det_image(p)
                or self.det_non_text(p) or self.det_special_blocks(p) or self.det_list(p) or ('normal', 1))

    # Определение заголовка
    @staticmethod
    def det_heading(p: Paragraph):
        # основной путь - определение на основе xml-разметки
        xml = p._p

        xml_style, xml_outline = (xml.xpath('.//w:pStyle/@w:val')[0].lower() if xml.xpath('.//w:pStyle/@w:val') else None,
                                  xml.xpath('.//w:outlineLvl/@w:val')[0] if xml.xpath('.//w:outlineLvl/@w:val') else None)
        if xml_style:
            if any(tp in xml_style for tp in ["heading", "заголовок", "title"]):
                if re.search(r'\d$', xml_style):
                    level = int(xml_style[-1])
                elif xml_outline is not None:
                    level = min(3, int(xml_outline) + 1)
                else:
                    if "title" in xml_style or "1" in xml_style:
                        level = 1
                    elif "subtitle" in xml_style or "2" in xml_style:
                        level = 2
                    else:
                        level = 3
                return 'heading', level

        # запасной метод: на основе текстовых паттернов
        text = p.text.strip().upper()

        for pattern in [
            r'^(РЕФЕРАТ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК ЛИТЕРАТУРЫ|ПРИЛОЖЕНИЯ?)$',
            r'^ГЛАВА\s+\d+[.:]?\s+[А-Я][А-Яа-яё\s\d\-]{10,}$',
            r'^РАЗДЕЛ\s+\d+[.:]?\s+[А-Я][А-Яа-яё\s\d\-]{10,}$',
            r'^ЧАСТЬ\s+[IVXLCDM]+[.:]?\s+[А-Я][А-Яа-яё\s\d\-]{10,}$'
        ]:
            if re.fullmatch(pattern, text.upper()):
                return 'heading', 1

        return None

    # Определяет специальные блоки (в пределах ГОСТа нужны только формулы и код)
    @staticmethod
    def det_special_blocks(p: Paragraph):
        text = p.text.strip()

        if not text:
            return None

        formula_signs = [
            r'.*[=≠≈<>≤≥∝→←↑↓↔].*',
            r'.*\b(sin|cos|tan|log|ln|exp|sqrt|sum|prod|int|lim)\b.*',
            r'.*\d+/\d+.*',
            r'.*[a-zA-Z]_{.*}.*',
            r'.*[a-zA-Z]\^\{.*\}.*',
            r'.*\b(alpha|beta|gamma|delta|epsilon|zeta|theta|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|chi|psi|omega)\b.*',
        ]
        res = 0
        for pat in formula_signs:
            if re.search(pat, text, flags=re.IGNORECASE):
                res += 1

        if res > 3:
            return "formula", 1

        code_signs = [
            r'^\s*(if|else|for|while|def|class|function|return|import|from)\b',
            r'.*\{.*\}.*',
            r'.*\(.*\).*;.*',
            r'.*//.*|.*/\*.*\*/.*',
            r'.*->.*|.*=>.*',
        ]
        res = 0
        for pat in code_signs:
            if re.search(pat, text, flags=re.IGNORECASE):
                res += 1

        if res > 3:
            return "code", 1

        return None

    # Комплексное определение списка
    @staticmethod
    def det_list(p: Paragraph):
        text = p.text.strip()
        if re.match(r'^[\-*+•]\s+.+', text):
            return 'list_bullet', 1
        elif re.match(r'^[\dа-яa-z]+[).]\s+.+', text, re.IGNORECASE):
            return 'list_number', 1
        xml = p._p
        num_pr = xml.xpath('.//w:numPr')
        if num_pr:
            level = int(xml.xpath('.//w:ilvl/@w:val')[0]) + 1 if xml.xpath('.//w:ilvl/@w:val') else 1
            num_id = xml.xpath('.//w:numId/@w:val')
            if num_id:
                list_type = 'bullet' if int(num_id[0]) % 2 else 'number'
                return f'list_{list_type}', level
            else:
                return 'list_bullet', level

        return None

    # Определение подписи к таблице
    @staticmethod
    def det_caption(p: Paragraph):
        xml, text = p._p, p.text.strip()

        if (re.match(r'^Таблица\s+\d+[.\-—].+', text)
                or re.match( r'^Таблица\s+[A-ZА-Я]+\s*\.\d+[.\-—].+',  text)):
            return 'tb_caption', 1

        xml_prev = xml.getprevious()
        text_match = re.match( r'^Таблица\s.+',  text)
        if xml_prev is not None and xml_prev.tag.endswith('tbl') and text_match:
            return 'tb_caption', 1

        xml_next = xml.getnext()
        if xml_next is not None and xml_next.tag.endswith('tbl') and text_match:
            return 'tb_caption', 1
        return None

    # Определение картинки
    @staticmethod
    def det_image(p: Paragraph):
        for r in p.runs:
            if r._element.find('.//' + qn('w:drawing')) is not None:
                return 'image', 1
        return None

    # Определение подписи к картинке
    @staticmethod
    def det_img_caption(p: Paragraph):
        xml, text = p._p, p.text.strip()

        if (re.match(r'^Рисунок\s+\d+[.\-].+', text)
                or re.match(r'^Рис.\s+[A-ZА-Я]+\s*\.\d+[.\-].+', text)):
            return 'img_caption', 1

        xml_prev = xml.getprevious()
        text_match = re.match(r'^Рисунок\s.+', text)
        if xml_prev is not None and xml_prev.tag.endswith('drawing') and text_match:
            return 'img_caption', 1

        xml_next = xml.getnext()
        if xml_next is not None and xml_next.tag.endswith('drawing') and text_match:
            return 'img_caption', 1
        return None

    # определение всего остального
    @staticmethod
    def det_non_text(p: Paragraph):
        if not p.text or re.fullmatch(r'^\s*$', p.text):
            return 'empty', 0
        return None

    def det_section(self, p: ParseResult):
        text = p["text"].strip()
        section_keywords = {
            'title_page': [
                'МИНИСТЕРСТВО', 'УНИВЕРСИТЕТ', 'КАФЕДРА', 'КУРСОВАЯ', 'ДИПЛОМ',
                'ВЫПУСКНАЯ', 'ТИТУЛЬНЫЙ'
            ],
            'abstract': ['РЕФЕРАТ', 'АННОТАЦИЯ'],
            'table_of_content': ['СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ'],
            'introduction': ['ВВЕДЕНИЕ'],
            'conclusion': ['ЗАКЛЮЧЕНИЕ', 'ВЫВОД'],
            'bibliography': [
                'СПИСОК ЛИТЕРАТУРЫ', 'БИБЛИОГРАФИЯ', 'ЛИТЕРАТУРА',
            ],
            'appendices': ['ПРИЛОЖЕНИЕ', 'ПРИЛОЖЕНИЯ']
        }
        for section, keywords in section_keywords.items():
            for keyword in keywords:
                if keyword in text.upper() and p["ptype"] == 'heading':
                    return section
        if p["ptype"] == 'heading' and re.match(r'^(ГЛАВА|РАЗДЕЛ)\s+[IVXLCDM\d]', text.upper()):
            return 'content'

        return self.cur_section

    def parse_with_sections(self):
        ctx = self.parse()
        structure = []
        for p in ctx:
            structure += [(p["index"], p["par"], self.det_section(p))]
        return ctx, structure

    def parse(self) -> List[ParseResult]:
        parse_ctx = []
        try:
            for i, par in enumerate(self.pars):
                pt, lvl = self.determine_type(par)
                parse_ctx += [ParseResult(level=lvl, ptype=pt, par=par, text=par.text, index=i)]
        except Exception as e:
            print(e)

        return parse_ctx