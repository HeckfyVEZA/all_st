import streamlit as st
st.set_page_config(layout="wide")
st.header("Программа по формированию КП по бланкам канального оборудования")
import docx2txt
from re import findall, IGNORECASE
from itertools import dropwhile, takewhile
from pandas import DataFrame, ExcelWriter, set_option
def to_excel(df, HEADER=False, START=1):
    output = __import__("io").BytesIO()
    writer = ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, header=HEADER,startrow=START, startcol=START, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    return output.getvalue()
set_option('display.max_columns', None)
sortset = lambda x: list(sorted(list(set(x))))
def sashQUA(main_system_name:str, the_mask = r'\d?[A-Za-zА-Яа-яЁё]\D?\D?\D?\D?\S?\S?\S?\S?\d{1,}[А-Яа-яЁё]?'):
    all_system_names = []
    for system_name in main_system_name.replace("-", " - ").split(','):
        system_name = system_name.strip()
        system_name_all = findall(the_mask, system_name, IGNORECASE)
        system_name_all = system_name_all if system_name_all else [main_system_name]
        if len(system_name_all) > 1:
            all_positions_in_system_name = tuple(tuple(filter(None, findall(r'(\D?|\d+)', sys_name, IGNORECASE))) for sys_name in system_name_all)
            system_condition = lambda x: x[0] == x[1]
            before_changing_part = ''.join(el[0] for el in takewhile(system_condition, zip(all_positions_in_system_name[0], all_positions_in_system_name[1])))
            after_changing_part = dropwhile(system_condition, zip(all_positions_in_system_name[0], all_positions_in_system_name[1]))
            changing_part_start, changing_part_thend = next(after_changing_part)
            after_changing_part = ''.join(el[0] for el in after_changing_part)
            all_system_names += [f"{before_changing_part}{'0' * (min(len(changing_part_start), len(changing_part_thend)) - len(str(i)))}{i}{after_changing_part}" for i in range(int(changing_part_start), int(changing_part_thend) + 1)]
        else:
            all_system_names += system_name_all
    return len(all_system_names)
def reg_chast(strochka, vent):
    choice = {220: [[0.22, 1, 'Регулятор скорости СРМ1-230В 1А IP20'],[0.44, 2, 'Регулятор скорости СРМ2-230В 2А IP20'],[0.55, 2.5, 'Регулятор скорости СРМ2,5Щ-230В 2,5А DIN IP20'],[0.66, 3, 'Регулятор скорости СРМ3-230В 3А IP20'],[0.88, 4, 'Регулятор скорости СРМ4-230В 4А IP20'],[0.88, 5, 'Регулятор скорости СРМ5Щ-230В 5А DIN IP20'],[1.1, 5, 'Регулятор скорости СРМ5-230В 5А IP20'],[1.5, 7, 'Регулятор скорости СРМ7-230В 7А IP20']],380: [[0.37, 1.5, 'Преобразователь частоты 0,37 кВт'],[0.75, 2.3, 'Преобразователь частоты 0,75 кВт'],[1.1, 2.3, 'Преобразователь частоты 1,1 кВт'],[1.5, 3.7, 'Преобразователь частоты 1,5 кВт'],[2.2, 5, 'Преобразователь частоты 2,2 кВт'],[4, 8.5, 'Преобразователь частоты 4 кВт'],[5.5, 12, 'Преобразователь частоты 5,5 кВт'],[7.5, 16, 'Преобразователь частоты 7,5 кВт'],[11, 24, 'Преобразователь частоты 11 кВт'],[15, 30, 'Преобразователь частоты 15 кВт'],[18.5, 37, 'Преобразователь частоты 18,5 кВт'],[22, 45, 'Преобразователь частоты 22 кВт'],[30, 60, 'Преобразователь частоты 30 кВт'],[37, 75, 'Преобразователь частоты 37 кВт']]}
    ch = choice[380] if "кварк" in vent.lower() else choice[220] if 'канал-вент' in vent.lower() else choice[int(findall(r'Uпит=~(\d+) В', strochka.replace(",","."))[0])]
    for item in ch:
        if item[0] >= float(findall(r'Ny=(\d+\.?\d*) кВт', strochka.replace(",","."))[0]) and item[1] > float(findall(r'Iпот=(\d+\.?\d*) A', strochka.replace(",","."))[0]):
            itog = item[2]
            break
    else:
        return "Подбор невозможен"
    return itog
def c_str(oborud):
    oborud = oborud.replace("M24-SR", "M24SR").replace("F24-S", "F24S").replace("Гермик", "ГЕРМИК")
    oborud = oborud[:-1] if oborud[-1] == "." else oborud
    mosh_slov = {"9":"12", "17":"18", "23":"24", "27":"30"}
    if "ГЕРМИК" in oborud:
        oborud = "-".join(oborud.split("-")[:-1] + ["Н", oborud.split("-")[-1]]) if len(oborud.split("-")) < 7 else oborud
    if "ЭКВ-К" in oborud and (not "," in oborud):
        oborud+=',0'
    elif "ЭКВ" in oborud:
        oborud = "-".join(oborud.split("-")[:-1] + [mosh_slov[oborud.split("-")[-1]]]) if oborud.split("-")[-1] in list(mosh_slov.keys()) else oborud
    groups = ["Канал-Регуляр", 'Канал-БОБ','Канал-ВЕНТ','Канал-ЕС','Канал-КВАРК','Канал-КВАРК-ФУД','Канал-ПКВ','Канал-КВ-','Канал-ЭКВ','Канал-ВКО','Канал-ФКО','Канал-КП','Канал-КВН','Канал-ФУД-Р-КОЖ','Канал-козырек','Канал-ФУД-козырек','Канал-ВИБР','Канал-ФУД-вибр','Канал-ФУД-Р-МК','Канал-крыша','Канал-КВАРК-ФУД-РКА','Канал-КВАРК-ФУД-РКО','Канал-РВК','Канал-РВС','Канал-РКН','Канал-РПВС','Канал-сетка','Канал-ФУД-сетка','Канал-C-PKT','Канал-ПКТ','Канал-ФКК','Канал-ФКП','Канал-МК','Канал-ГКД','Канал-ГКК','Канал-ГКП','КЛАБ', 'Канал-Гермик', "Канал-ГКВ", "Канал-ФУД-ГКВ", "Канал-ФУД-Регуляр", "Канал-КОЛ", "Канал-ФУД-Тюльпан", "ГЕРМИК", 'Канал-КВАРК-П']
    apendix = ["Клапан ", 'Блок обеззараживания ','Вентилятор ','Вентилятор ','Вентилятор ','Вентилятор ','Вентилятор ','Клапан ','Воздухонагреватель ','Воздухоохладитель ','Воздухоохладитель ','Каплеуловитель ','Воздухонагреватель ','Кожух ','Козырек ','Козырек ','Комплект основы виброизолирующий ','Комплект основы виброизолирующий ','Кронштейн ','Крыша ','Решетка ','Решетка ','Решетка ','Решетка ','Решетка ','Решетка ','Сетка ','Сетка ','Теплоутилизатор ','Теплоутилизатор ','Фильтр ','Фильтр ','Хомут ','Шумоглушитель ','Шумоглушитель ','Шумоглушитель ','Клапан ', 'Клапан ', 'Гибкая вставка ', 'Гибкая вставка ', 'Клапан ', 'Клапан ', 'Клапан ','Клапан ', 'Вентилятор ']
    try:
        indexes = [groups.index(gr) for gr in groups if gr in oborud]
        return apendix[max(indexes)] + oborud
    except:
        return oborud
filess = st.file_uploader("Перетащите сюда бланки (допускается ТОЛЬКО формат .DOCX)",type='docx', accept_multiple_files=True)
all_oborud = []
pbi, mlen = 0, 0
if len(filess):
    progr = st.progress(0)

for f in filess:
    try:
        info = [item for item in docx2txt.process(f).split("\n") if len(item)>0]
        reggii = []
        r = 0
        name = info[info.index("Название:") + 1]
        quan = sashQUA(name)
        oborud = []
        for item in info:
            try:
                if len(findall(r'(\d+\.)', item.split()[0])):
                    oborud.append([c_str((info[info.index(item)+1]+ ";  ").split(";")[0].split()[1].strip()), quan, name])
            except:
                pass
            try:
                if "эл.двиг" in item.lower():
                    vent = oborud[-1][0]
                    reggii.append(reg_chast(item, vent))
            except:
                pass
            if "Дополнительное оборудование:" in item:
                i = info.index(item)+1
                while not ("габаритная схема" in info[i].lower() or "габаритные размеры" in info[i].lower()):
                    try:
                        oborud.append([c_str(info[i].split(": ")[1].split(" - ")[0].strip()), int(findall(r'(\d+) шт.', info[i].split(": ")[1])[0])* quan, name])
                    except:
                        if "регулятор" in info[i].lower() and "да" in info[i].lower():
                            oborud.append([reggii[r], quan, name])
                            r+=1
                    i+=1
        all_oborud += oborud
    except:
        pass
    try:
        mlen = len(oborud) if len(oborud) > mlen else mlen
    except:
        mlen = 0
    pbi+=1
    progr.progress(pbi/len(filess))
final = [[None, n, 0, ''] for n in sorted(sortset([obo[0] for obo in all_oborud]))]
for ob in all_oborud:
    nom_idx = sorted(sortset([obo[0] for obo in all_oborud])).index(ob[0])
    final[nom_idx][2] += ob[1]
    if not (ob[2] in final[nom_idx][3].split(", ")):
        final[nom_idx][3] += ob[2] + ", "
for i in range(len(final)):
    final[i][3] = final[i][3][:-2]
fintable = DataFrame(list(map(lambda x: x[1:],final)))
fintable_withcols = DataFrame(list(map(lambda x: x[1:],final)), columns=["Номеклатура", "Количество", "Номера систем"])
final = [[None, None, None, None]] + final
checklist = {}
if len(final)>1:
    for syss in sortset([obo[2] for obo in all_oborud]):
        checklist[syss] = [f'{obo[0]} | {obo[1]}шт.' for obo in all_oborud if obo[2]==syss]
    list_of_cols = [c for c in checklist.keys()]
    num_of_cols = 5
    appp = 0 if not len(list_of_cols)%num_of_cols else num_of_cols-len(list_of_cols)%num_of_cols
    list_of_cols+= [None]*appp
    col = {}
    st.write("#### ТАБЛИЦА ДЛЯ КП ###")
    st.dataframe(fintable_withcols, width=2000)
    with st.expander("Состав систем"):
        for i in range(0, len(list_of_cols), num_of_cols):
            col[i] = st.columns(num_of_cols)
            for j in range(num_of_cols):
                exec(f"""
with col[{i}][{j}]:\n
\tif not list_of_cols[{i+j}] is None:
\t\tst.write(f"#### Состав системы {list_of_cols[i+j]} ###")
\t\tfor item in checklist[list_of_cols[{i+j}]]:
\t\t\tst.write(item.replace("*", "\*"))""")
    for syss in sortset([obo[2] for obo in all_oborud]):
        while len(checklist[syss]) < mlen+1:
            checklist[syss].append(None)
st.download_button(label='💾 Скачать файл для выгрузки в КП',data=to_excel(fintable) ,file_name= 'для кп.xls')
# st.write([len(checklist[chk]) for chk in checklist.keys()])
try:
    st.download_button(label='💾 Скачать проверочный файл',data=to_excel(DataFrame(checklist), HEADER=True, START=0) ,file_name= 'проверка.xlsx')
except:
    cormlen = max([len(checklist[chk]) for chk in checklist.keys()])
    for chk in checklist.keys():
        checklist[chk] = checklist[chk] + [None]*(cormlen-len(checklist[chk]))
    st.download_button(label='💾 Скачать проверочный файл',data=to_excel(DataFrame(checklist), HEADER=True, START=0) ,file_name= 'проверка.xlsx')