from docxtpl import DocxTemplate
import pandas as pd


# ------------------------------------
#  ------     НОМЕР ID ---------------
num_id = 5
# ------------------------------------

def docgenerator(num_id):

    # загрузка файлов шаблонов
    doc_dgvr = DocxTemplate('templates_dgvr.docx')
    doc_dgvr_p_1 = DocxTemplate('templates_dgvr_p1.docx')
    doc_dgvr_p_2 = DocxTemplate('templates_dgvr_p2.docx')
    doc_dgvr_ds = DocxTemplate('templates_dgvr_ds.docx')
    doc_dgvr_ds_p1 = DocxTemplate('templates_dgvr_ds_p1.docx')

    # сделаю датафрейм пандас из базы данных по компаниям и договорам
    df = pd.read_excel('BD_act.xlsx', 'tb_comp')

    # из датафрейма выберу нужную строку по id позиции, это будет кастом-датафрейм
    df_custom = df[df['{{ id }}'] == num_id]

    # сделаю список заголовков кастом датафрейма, что бы его привести к нужно виду, для этого запущу цикл
    list_columns = df_custom.columns.tolist()

    # в цикле буду перебирать заголовки кастом датафрейма
    for i in range(0, len(list_columns)):
        # и каждое значение заголовка я буду заменять фактически на него же, только убрав последние и первые три
        # символа (пробел и две фигурные скобки)
        df_custom.columns.values[i] = list_columns[i][3:-3]

    # из кастом датафрейма оставлю только значения
    df_custom_iloc = df_custom.iloc[0]
    # сделаю словарь из значений кастом датафрейма
    df_custom_iloc_dict = df_custom_iloc.to_dict()

    # словарь значений кастом датафрейма помещу в переменную context для последующей подстановки в шаблоны
    context = df_custom_iloc_dict

    # подставлю значения содержащиеся в context в шаблоны
    doc_dgvr.render(context)
    doc_dgvr_p_1.render(context)
    doc_dgvr_p_2.render(context)
    doc_dgvr_ds.render(context)
    doc_dgvr_ds_p1.render(context)

    # сохраню в финальный документ
    doc_dgvr.save("Dgv.docx")
    doc_dgvr_p_1.save("Dgvr_P1.docx")
    doc_dgvr_p_2.save("Dgvr_P2.docx")
    doc_dgvr_ds.save("Dgvr_DS.docx")
    doc_dgvr_ds_p1.save("Dgvr_DS_P1.docx")

docgenerator(num_id)