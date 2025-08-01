from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.shared import Inches, Pt


def text_underline(new_paragraf: object, number: int, count_kv: int, data: str, flag_size: bool):
    # Добавляем текст в список. Первое булевое значение показывает где нужно подчеркивание
    # Второе булевое значение показывает где нужен другой размер  
    text_parts = [
    (f"№{number}", True, False),
    ("  годно до ", False, False),
    (f"{count_kv} кВ", True, False),
    ("\nДата следующего испытания\n", False, False),
    (f"{data}г.", True, False),
    ("\nЭлектротехническая лаборатория\n",False, False),
    ("АО «Коми коммунальные технологии»", False, True),
    ]

    # Добавляем текст в параграф с подчеркиванием там, где нужно
    # И увеличиваем текст там где нужно и если нужна большая бирка
    if flag_size:
        for text, underline, size_big in text_parts:
            run = new_paragraf.add_run(text)
            run.font.size = Pt(7)
            if underline:
                run.underline = WD_UNDERLINE.SINGLE  # Устанавливаем подчеркивание

            if size_big:
                run.font.size = Pt(8)
    else:
        for text, underline, size_big in text_parts:
            run = new_paragraf.add_run(text)
            run.font.size = Pt(10)
            if underline:
                run.underline = WD_UNDERLINE.SINGLE  # Устанавливаем подчеркивание

            if size_big:
                run.font.size = Pt(11)

            if "Электротехническая лаборатория" in text:
                run.font.size = Pt(9)

    return new_paragraf


def generate_birks(count_birk: int, date: list[list], flag_size) -> object:
    doc = Document()

    # Добавляем стили документу
    style = doc.styles["Normal"]
    style.font.name = 'Times New Roman'

    if flag_size:
        cols = 4
    else:
        cols = 3
    
    rows = count_birk // 4 + 1
    
    # Добавление таблицы
    table = doc.add_table(rows=rows, cols=cols)

    # Убираем авто растягивание таблицы
    table.autofit = False
    table.allow_autofit = False

    index_date = 0
    # Заполнение ячеек таблицы
    for i in range(rows):
        for j in range(cols):
            # Достаем данные из переданного списка
            if index_date < len(date):
                number_prot = date[index_date][0]
                count_kv = date[index_date][1]
                data_prot = date[index_date][2]
            else:
                break
 
            # Создает ячейку таблицы
            cell = table.cell(i, j)

            # Убирает лишний параграф в начале
            cell._element.remove(cell.paragraphs[0]._element)

            # Задает ширину ячейки
            if flag_size:
                cell.width = Inches(1.6)
            else:
                cell.width = Inches(2)
            
            # Добавляет параграф, и форматирует его с помошью отдельной функции
            paragraph = cell.add_paragraph()
            text_underline(new_paragraf=paragraph, number=number_prot, count_kv=count_kv, data=data_prot, flag_size=flag_size)

            # Выбирает параграф и добавляет ему выравнивание по центру
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            index_date += 1

    # Добавляет границы таблице
    table.style = "Table Grid"
    return doc
    
