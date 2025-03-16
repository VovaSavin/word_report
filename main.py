import os
import random
import platform, shutil
from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.text.paragraph import Paragraph

FONT_FAMILY = [
    ("NinaCTT", 14),
    ("BetinaScriptC", 14),
    ("BirchCTT", 14),
    ("FreestyleScript", 18),
    ("MisterK", 22),
]

MAIN_TXT = """
\tПрошу звільнити мене за власним бажанням з лав національної гвардії України. Зобов’язуюсь відслужити два повних тижні 
у розмірі 14 календарних днів. Буду пити та гуляти ці два тижні.
"""


def replacer(txt: str):
    return txt.replace(
        "і", "i"
    ).replace(
        "І", "I"
    ).replace(
        "ї", "i"
    ).replace(
        "Ї", "i"
    )


def p_indent_l(ind: float, paragraph: Paragraph):
    """
    Відступи ліворуч
    :param ind:
    :param paragraph:
    :return:
    """
    format_par = paragraph.paragraph_format
    format_par.left_indent = Inches(ind)


def p_indent_r(ind: float, paragraph: Paragraph):
    """
    Відступи праворуч
    :param ind: A float multiplier to determine the scaling of the paragraph indent.
    :param paragraph: Paragraph object whose indentation is to be adjusted.
    :return: The updated paragraph indent after applying the scaling multiplier.
    :rtype: float
    """
    format_par = paragraph.paragraph_format
    format_par.right_indent = Inches(ind)


def p_indent_first_line(idn: float, paragraph: Paragraph):
    """
    Відступ першого рядка
    :param idn:
    :param paragraph:
    :return:
    """
    fromat_par = paragraph.paragraph_format
    fromat_par.first_line_indent = Inches(idn)


class CreateDocument:
    def __init__(
            self, first_name: str, last_name: str,
            military_rank: str,
            major_data: dict[str, str],
    ):
        self.first_name = first_name
        self.last_name = last_name
        self.military_rank = military_rank
        self.major_data = major_data
        self.doc = Document()
        self.font_type = random.choice(FONT_FAMILY)

    def set_paragraph_head(self, txt: str, left_indent=True):
        pargrph = self.doc.add_paragraph(txt)
        if left_indent:
            p_indent_l(2.6, pargrph)
        else:
            pargrph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_4 = pargrph.runs[0]
        font_4 = run_4.font
        font_4.size = Pt(self.font_type[1])
        font_4.name = self.font_type[0]

    def set_paragraph_main(self):
        text = self.doc.add_paragraph(replacer(MAIN_TXT))
        run_4 = text.runs[0]
        font_4 = run_4.font
        font_4.size = Pt(self.font_type[1])
        font_4.name = self.font_type[0]

    def set_paragraph_date(self, txt: str):
        pargrph = self.doc.add_paragraph(txt)
        run_4 = pargrph.runs[0]
        font_4 = run_4.font
        font_4.size = Pt(self.font_type[1])
        font_4.name = self.font_type[0]

    def start_write_raports(self):
        self.set_paragraph_head(replacer("Командиру військової частини 3027"))
        self.set_paragraph_head(replacer(self.major_data["rank"] + "у"))
        self.set_paragraph_head(replacer(self.major_data["full_name"]))
        self.set_paragraph_head(replacer(f"{self.first_name} {self.last_name}"))
        self.doc.add_paragraph("\n")
        self.set_paragraph_head("Рапорт", left_indent=False)
        self.set_paragraph_main()
        self.doc.add_paragraph("\n")
        self.set_paragraph_date(self.major_data["date"])
        self.doc.save(f"{self.first_name}_{self.last_name}_рапорт.docx")


class PreviousPrepared:
    """
    Поідготовка даних для зчитування та запису рапортів
    Файл з іменами вс
    Шрифти та куди їх копіювати в залежності від ОС
    """

    def __init__(self):
        self.file_names = r"names.txt"
        self.pathes_systems = {
            "Windows": r"C:/Windows/Fonts",
            "Linux": r"/usr/share/fonts",
            "Darwin": r"/Library/Fonts",
        }

    def copy_fonts(self) -> None:
        """
        Копіює шрифти із директорії fonts в належну директорію ОС
        :return:
        """

        os_name = platform.system()
        files_list = os.listdir(r"fonts")
        for x in range(len(files_list)):
            print(
                os.path.exists(
                    self.pathes_systems[os_name] + files_list[x]
                )
            )
            if not os.path.exists(
                    self.pathes_systems[os_name] + files_list[x]
            ) and not os.path.isfile(
                self.pathes_systems[os_name] + files_list[x]
            ):
                shutil.copy(r"fonts/{}".format(files_list[x]), self.pathes_systems[os_name])
            else:
                continue

    def open_get_names(self) -> list[str]:
        """
        Відкриває текстовий файл з іменами в/с та формує з них список
        для подальшого використання в формуванні *.docx файлів
        :return:
        """
        with open(self.file_names, "r", encoding="utf-8") as fl:
            content = fl.read()
            array_content = content.split("\n")
            return array_content


def run(data: dict, switcher: bool = True):
    """
    Запускає цикл перебору імен
    та запускає метод класу CreateDocument для формування *.docx файлів
    :return:
    """
    helper = PreviousPrepared()
    helper.copy_fonts()
    if switcher:
        array_militaries = helper.open_get_names()
        for x in range(len(array_militaries)):
            first = array_militaries[x].split()[0].strip()
            last = array_militaries[x].split()[1].strip()
            rank = array_militaries[x].split()[2].strip()
            my_doc = CreateDocument(
                first_name=first,
                last_name=last,
                military_rank=rank,
                major_data=data,
            )
            my_doc.start_write_raports()
    else:
        pass


run(
    {"rank": "test", "full_name": "test test", "date": "05.03.2025"},
    switcher=True,
)