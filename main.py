from pptx import Presentation
from pptx.util import Pt


def make_pptx():
    prs = Presentation('src/template.pptx')
    parse_pptx(prs)
    prs.save('out/prs.pptx')


def add_text_box(slide, text, left, top, width=None, height=None, font_size=28):
    text_box = slide.shapes.add_textbox(
        Pt(left), Pt(top), Pt(width), Pt(height))
    tf = text_box.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)


def add_table(slide, data: list, left, top, width, height):
    rows = len(data)
    cols = len(data[0])
    table = slide.shapes.add_table(rows, cols, Pt(
        left), Pt(top), Pt(width), Pt(height)).table
    for i in range(rows):
        for j in range(cols):
            cell = table.cell(i, j)
            cell.text = str(data[i][j])


def add_picture(slide, path, left, top, width=None, height=None):
    slide.shapes.add_picture(path, Pt(left), Pt(top), Pt(width), Pt(height))


def parse_pptx(prs):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                x = 0
            elif shape.has_table and shape.name == 'Таблица 2':
                table = shape.table
                cell = table.cell(1, 1)
                cell.text = '100%'

                cell = table.cell(1, 2)
                cell.text = '100%'

                cell = table.cell(1, 3)
                cell.text = '100%'

                cell = table.cell(2, 1)
                cell.text = '100%'

                cell = table.cell(2, 2)
                cell.text = '100%'

                cell = table.cell(2, 3)
                cell.text = '100%'


def make_test_presentation():
    # make presentation file
    prs = Presentation()
    prs.slide_width = Pt(1280)
    prs.slide_height = Pt(720)
    # 6 - blank layout
    # make slide
    slide_1 = prs.slides.add_slide(prs.slide_layouts[6])

    header_text = '''Исполнение КПЭ
    «Повышение исполняемости договоров», от 100 млн руб.*'''
    add_text_box(slide_1, header_text, 10, 10, 900, 20)
    add_picture(slide_1, 'src/logo.png', 1280-250, 10, 200, 75)

    data = [
        [
            'Общее количество КС до конца года*',
            'Общее количество КС на дату',
            'Количество КС исполненных в срок',
            'Количество КС  не исполненных в срок',
            'Фактическое значение доли КС, выполненных в срок, %'],
        [100, 100, 100, 100, 100]
    ]
    add_table(slide_1, data, 50, 100, 800, 200)
    prs.save('out/new.pptx')


if __name__ == '__main__':
    make_test_presentation()
