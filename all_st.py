import streamlit as st
st.set_page_config(layout="wide")
tabsis = st.tabs(["КП каналка", "Частотники", "Автоматика"])
with tabsis[0]:
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
        if not HEADER:
            format1 = workbook.add_format({'num_format': '0.00'})
            worksheet.set_column('A:A', None, format1)
        else:
            for idx, col in enumerate(df):
                series = df[col]
                max_len = max((
                    series.astype(str).map(len).max(), len(str(series.name))
                )) + 1
                worksheet.set_column(idx,idx,max_len)
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
        oborud = oborud.replace("M24-SR", "M24SR").replace("F24-S", "F24S").replace("Гермик", "ГЕРМИК").replace('-H','-Н')
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
    def ficula(x):
        hlop = []
        ficula = []
        for i in range(len(x)):
            masshlop = []
            masshlop2 = []
            masshlop = list(map(lambda x: [x[1],x[2],x[0]],x[i]))
            for i in range(len(masshlop)):
                for y in range(len(masshlop)):
                    if hlop == []:
                        hlop.append(masshlop[i][0])
                        hlop.append(0)
                    if hlop[0] == masshlop[y][0] and masshlop[y] != masshlop[-1]:
                        hlop[1] += masshlop[y][1]
                        hlop.extend(masshlop[y][2:])
                    elif hlop[0] == masshlop[y][0] and y == len(masshlop)-1:
                        hlop[1] += masshlop[y][1]
                        hlop.extend(masshlop[y][2:])
                        if hlop not in masshlop2 and hlop[1] != 0:
                            masshlop2.append(hlop)
                        hlop = []
                    elif hlop[0] != masshlop[y] and y == len(masshlop)-1:
                        if hlop not in masshlop2 and hlop[1] != 0:
                            masshlop2.append(hlop)
                        hlop = []
            for i in range(len(masshlop2)):
                masshlop2[i] = masshlop2[i][:2] + [', '.join(masshlop2[i][2:])]
            ficula.append(sorted(sorted(masshlop2,key=lambda x:x[0]), key=lambda x:x[-1]))
        return ficula
    filess = st.file_uploader("Перетащите сюда бланки (допускается ТОЛЬКО формат .DOCX)",type=('docx',), accept_multiple_files=True)
    all_oborud = []
    pbi, mlen = 0, 0
    if len(filess):
        progr = st.progress(0)
    all_inf_stroka, all_inf_contruct, all_inf, error_files = [], [], [[], [], []], []
    for f in filess:
        try:
            info = [item for item in docx2txt.process(f).split("\n") if item]
            reggii = []
            r = 0
            name = info[info.index("Название:") + 1]
            quan = sashQUA(name)
            oborud = []
            for item in info:
                try:
                    if len(findall(r'(\d+\.)', item.split()[0])):
                        oborud.append([c_str((info[info.index(item)+1]+ ";  ").split(";")[0].split('Индекс:')[1].strip()), quan, name])
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
        try:
            info = [item.strip() for item in docx2txt.process(f).split("\n") if len(item.strip())]
            system, info = [item.replace('БЛАНК-ЗАКАЗ ','')[:-13].strip() for item in info if len(findall(r'(от \d\d.\d\d.\d\d\d\d)', item))][0], [list(map(lambda x :x.replace('шт.','')[:-1].strip() if x[-1]=='-' else x.replace('шт.','').strip(),[' '.join(item.replace('–','-').split(' - ')[0].split(' ')[1:]), item.replace('–','-').split(' - ')[1] if len(item.replace('–','-').split(' - ')) == 2 else '1шт.'])) for item in info if ((len(findall(r'(\d+\. )', item)) and (not 'аэро' in item.lower ()) and (not 'габар' in item.lower())))]
            info = [[system]+item for item in info]
            idx = 0 if info[0][1][:3] =='ВРА' or (info[0][1][:3]=='ОСА' and (not info[0][1][4] in 'ЕE') and (not ('ДУ' in info[0][1]))) else 1 if not 'ВКОП' in info[0][1] else 2
            info[0] = [info[0][0], 'Вентилятор '+info[0][1] if idx!=2 else 'Установка приточная '+info[0][1], info[0][2]]
            all_inf[idx]+=[inf[:2]+[int(inf[-1])] for inf in info]
        except:
            error_files.append(f.name)
        pbi+=1
        progr.progress(pbi/len(filess))
    sorted_stroks = ficula(all_inf)
    all_oborud+=sorted_stroks[1]
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
    
    # st.write(all_oborud)
    st.download_button(label='Разбивка для кп по системам',data=to_excel(DataFrame([[None, None, None, None]] + all_oborud), file_name='NNV.xls')
    # st.write([len(checklist[chk]) for chk in checklist.keys()])
    try:
        st.download_button(label='💾 Скачать проверочный файл',data=to_excel(DataFrame(checklist), HEADER=True, START=0) ,file_name= 'проверка.xlsx')
    except:
        cormlen = max([len(checklist[chk]) for chk in checklist.keys()])
        for chk in checklist.keys():
            checklist[chk] = checklist[chk] + [None]*(cormlen-len(checklist[chk]))
        st.download_button(label='💾 Скачать проверочный файл',data=to_excel(DataFrame(checklist), HEADER=True, START=0) ,file_name= 'проверка.xlsx')
with tabsis[1]:
    from docx2txt import process
    from re import findall, IGNORECASE
    from itertools import dropwhile, takewhile
    def sashQUA(main_system_name:str, the_mask = r'\d?[A-Za-zА-Яа-яЁё]\D?\D?\D?\D?\S?\S?\S?\S?\d{1,}[А-Яа-яЁё]?'):
        system_names = main_system_name.replace("-", " - ").split(',')
        all_system_names = []
        for system_name in system_names:
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
    fty = st.radio(
        "#### Какой тип частотников нужно подобрать? ###",
        ("Обезличенный", "ESQ", "INNOVERT"))
    def NyI_rec(fan):
        kvt, pole = list(map(int, fan.split("/")))
        toks = {2: {0.37: 1, 0.55: 1.43, 0.75: 1.92, 1.1:2.74, 1.5:3.46, 2.2:5.21, 3.0:7.03, 4.0:7.9, 5.5:10.7, 7.5:15, 11.0:22, 15.0:30, 18.5:35, 22.0:42, 30.0:56, 37.0:70},4: {0.25: 1.16, 0.37: 1.37, 0.55:1.8, 0.75:2.23, 1.1: 3.03, 1.5:3.78, 2.2:5.78, 3.0:7.17, 4.0:8.5, 5.5:12, 7.5:15.6, 11.0:22, 15.0:29, 18.5:35, 22.0:42, 30.0:56, 37.0:70, 45.0:86, 55.0:104, 75.0:139, 90.0:169, 110.0:195, 132.0:231, 160.0:279},6: {0.18:0.99, 0.25:1.29, 0.37:1.55, 0.55:2, 0.75:2.61, 1.1:3.39, 1.5:4.74, 2.2:6.1, 3.0:7.6, 4.0:9.4, 5.5:12.2, 7.5:17.5, 11.0:23, 15.0:31, 18.5:37, 22.0:46, 30.0:59, 37.0:66, 45.0:81, 55.0:97, 75.0:133},8: {0.25:1.39, 0.37:1.87, 0.55:2.62, 0.75:2.99, 1.1:4.09, 1.5:4.83, 2.2:6.74, 3.0:9.1, 4.0:9.6, 5.5:13, 7.5:18, 11.0:26, 15.0:35, 18.5:40, 22.0:49, 30.0:64, 37.0:73, 45.0:111, 55.0:113, 75.0:153}}
        return (float(kvt)/100)*1.1, (toks[int(pole)][float(kvt)/100])*1.1
    def chast_rec(path, fretype):
        freq = [[0.37, 1.5, 'Преобразователь частоты 0,37 кВт'],[0.75, 2.3, 'Преобразователь частоты 0,75 кВт'],[1.1, 2.3, 'Преобразователь частоты 1,1 кВт'],[1.5, 3.7, 'Преобразователь частоты 1,5 кВт'],[2.2, 5, 'Преобразователь частоты 2,2 кВт'],[4, 8.5, 'Преобразователь частоты 4 кВт'],[5.5, 12, 'Преобразователь частоты 5,5 кВт'],[7.5, 16, 'Преобразователь частоты 7,5 кВт'],[11, 24, 'Преобразователь частоты 11 кВт'],[15, 30, 'Преобразователь частоты 15 кВт'],[18.5, 37, 'Преобразователь частоты 18,5 кВт'],[22, 45, 'Преобразователь частоты 22 кВт'],[30, 60, 'Преобразователь частоты 30 кВт'],[37, 75, 'Преобразователь частоты 37 кВт']]
        esq = [[0.75,2.3,'Преобразователь частоты ESQ-760-4T-0007 0,75/1,5кВт 380В арт. 08.04.000642'],[1.5,3.7,'Преобразователь частоты ESQ-760-4T-0015 1,5/2,2кВт 380В арт. 08.04.000643'],[2.2,5.1,'Преобразователь частоты ESQ-760-4T-0022 2,2/4кВт 380В арт. 08.04.000644'],[4,8.5,'Преобразователь частоты ESQ-760-4T-0040 4/5,5кВт 380В арт. 08.04.000645'],[7.5,17,'Преобразователь частоты ESQ-760-4T0055G/0075P 5,5/7,5кВт 380В арт. 08.04.000477'],[11,25,'Преобразователь частоты ESQ-760-4T0075G/0110P 75/11кВт 380В арт. 08.04.000478'],[15,32,'Преобразователь частоты ESQ-760-4T0110G/0150P 11/15кВт 380В арт. 08.04.000479'],[18.5,37,'Преобразователь частоты ESQ-760-4T0150G/0185P 15/18,5кВт 380В арт. 08.04.000480'],[22,45,'Преобразователь частоты ESQ-760-4T0185G/0220P 18,5/22кВт 380В арт. 08.04.000481'],[30,60,'Преобразователь частоты ESQ-760-4T0220G/0300P 22/30кВт 380В арт. 08.04.000482'],[37,75,'Преобразователь частоты ESQ-760-4T0300G/0370P-BU 30/37кВт 380В арт. 08.04.000728'],[45,91,'Преобразователь частоты ESQ-760-4T0370G/0450P-BU 37/45кВт 380В арт. 08.04.000729'],[55,112,'Преобразователь частоты ESQ-760-4T0450G/0550P-BU 45/55кВт 380В арт. 08.04.000706'],[75,150,'Преобразователь частоты ESQ-760-4T0550G/0750P-BU 55/75кВт 380В арт. 08.04.000707'],[90,176,'Преобразователь частоты ESQ-760-4T0750G/0900P 75/90кВт 380В арт. 08.04.000487'],[110,210,'Преобразователь частоты ESQ-760-4Т0900G/1100P 90/110кВт 380В арт. 08.04.000488']]
        innovert = [[0.4,1.5,'Преобразователь частоты INNOVERT 0,4кВт IP65 арт. IPD401P43B'],[0.75,2.7,'Преобразователь частоты INNOVERT 0,75кВт IP65 арт. IPD751P43B'],[1.1,3,'Преобразователь частоты INNOVERT 1,1кВ IP65 арт. IPD112P43B'],[1.5,4,'Преобразователь частоты INNOVERT 1,5кВт IP65 арт. IPD152P43B'],[2.2,5,'Преобразователь частоты INNOVERT 2,2кВт IP65 арт. IPD222P43B'],[3,6.8,'Преобразователь частоты INNOVERT 3кВт IP65 арт. IPD302P43B'],[4,8.6,'Преобразователь частоты INNOVERT 4кВт IP65 арт. IPD402P43B'],[5.5,12.5,'Преобразователь частоты INNOVERT 5,5кВт IP54 арт. IPD552P43B'],[7.5,17.5,'Преобразователь частоты INNOVERT 7,5кВт IP54 арт. IPD752P43B'],[11,24,'Преобразователь частоты INNOVERT 11кВт IP54 арт. IPD113P43B'],[15,33,'Преобразователь частоты INNOVERT 15кВт IP54 арт. IPD153P43B'],[18.5,40,'Преобразователь частоты INNOVERT 18,5кВт IP54 арт. IPD183P43B'],[22,45,'Преобразователь частоты INNOVERT 22кВт IP54 арт. IPD223P43B'],[30,60,'Преобразователь частоты INNOVERT 30кВт IP54 арт. IPD303P43B'],[37,80,'Преобразователь частоты INNOVERT 37кВт IP54 арт. IPD373P43B'],[45,90,'Преобразователь частоты INNOVERT 45кВт IP54 арт. IPD453P43B']]
        fff = {"Обезличенный":freq, "ESQ":esq, "INNOVERT":innovert}[fretype]
        info = [item for item in process(path).split("\n") if item]
        name = path.name.split("\\")[-1].replace(".docx", "").replace(".doc","")
        try:
            fan = [findall(r"(\d\d\d\d\d\/\d\d?)", N)[0] for N in info if len(findall(r"(\d\d\d\d\d\/\d\d?)", N))][0]
            Ny, I = NyI_rec(fan)
            for fre in fff:
                if Ny<=fre[0] and I<fre[1]:
                    chastotnik = fre[2]
                    break
            else:
                chastotnik = "Подбор невозможен"
            if any(["ЧР: да" in s for s in info]) or any([["ВОСК" in s for s in info]]) or ([["ВР" in s for s in info]]):
                try:
                    st.write(f"Ny = {round((Ny/1.1),1)} кВт; I = {round((I/1.1),1)} A")
                    yield name + " | " + chastotnik + " | "+str(sashQUA(name))+" шт."
                except:
                    st.write(f"Ny = {round((Ny/1.1),1)} кВт; I = {round((I/1.1),1)} A")
                    yield name + " | "+ chastotnik + " | необрабатываемое имя"
            else:
                yield False
        except:
            # try:
                infos = info[:]
                info = [item for item in infos if (len(item)>0 and "ВОСК" in item and "блок" in item) and 'сифон' not in item] + [item for item in infos if (len(item)>0 and "ВР" in item and "блок" in item) and 'сифон' not in item]
                st.write(info)
                rezervs = [item for item in infos if 'резерв' in item.lower()]
                # st.write(info)
                # st.write(rezervs)
                if len(rezervs):
                    if infos.index(rezervs[0])<infos.index(info[0]):
                        rezervs = rezervs[1:]
                al_rez = len(rezervs)       
                for infi in info:
                    Ny = float(findall(r"Ny=(\d+[\.,]?\d?\d?)кВт", infi)[0])
                    obmin = float(findall(r"nдв=(\d+)об/мин", infi)[0])
                    poles = 2 if obmin > 1500 else 4 if obmin > 1000 else 6
                    I = {2: {0.37: 1, 0.55: 1.43, 0.75: 1.92, 1.1:2.74, 1.5:3.46, 2.2:5.21, 3.0:7.03, 4.0:7.9, 5.5:10.7, 7.5:15, 11.0:22, 15.0:30, 18.5:35, 22.0:42, 30.0:56, 37.0:70},4: {0.25: 1.16,      0.37: 1.37, 0.55:1.8, 0.75:2.23, 1.1: 3.03, 1.5:3.78, 2.2:5.78, 3.0:7.17, 4.0:8.5, 5.5:12, 7.5:15.6, 11.0:22, 15.0:29, 18.5:35, 22.0:42, 30.0:56, 37.0:70,  45.0:86, 55.0:104, 75.0:139, 90.0:169, 110.0:195, 132.0:231, 160.0:279},6: {0.18:0.99, 0.25:1.29, 0.37:1.55, 0.55:2, 0.75:2.61, 1.1:3.39, 1.5:4.74, 2.2:6.1, 3.0:7.6, 4.0:9.4, 5.5:12.2, 7.5:17.5, 11.0:23, 15.0:31, 18.5:37, 22.0:46, 30.0:59, 37.0:66, 45.0:81, 55.0:97, 75.0:133},8: {0.25:1.39, 0.37:1.87, 0.55:2.62, 0.75:2.99, 1.1:4.09, 1.5:4.83, 2.2:6.74, 3.0:9.1, 4.0:9.6, 5.5:13, 7.5:18, 11.0:26, 15.0:35, 18.5:40, 22.0:49, 30.0:64, 37.0:73, 45.0:111, 55.0:113, 75.0:153}}[poles][Ny]
                    for fre in fff:
                        if Ny<=fre[0] and I<fre[1]:
                            chastotnik = fre[2]
                            break
                    else:
                        chastotnik = "Подбор невозможен"
                    kolvorez = 2 if al_rez > 0 else 1
                    al_rez-=1
                    yield name + " | " + chastotnik + " | "+str(kolvorez)+" шт."
            # except:
            #     choice = {220: [[0.22, 1, 'Регулятор скорости СРМ1-230В 1А IP20'],[0.44, 2, 'Регулятор скорости СРМ2-230В 2А IP20'],[0.55, 2.5, 'Регулятор скорости СРМ2,5Щ-230В 2,5А DIN IP20'],[0.66, 3, 'Регулятор скорости СРМ3-230В 3А IP20'],[0.88, 4, 'Регулятор скорости СРМ4-230В 4А IP20'],[0.88, 5, 'Регулятор скорости СРМ5Щ-230В 5А DIN IP20'],[1.1, 5, 'Регулятор скорости СРМ5-230В 5А IP20'],[1.5, 7, 'Регулятор скорости СРМ7-230В 7А IP20']],380: [[0.37, 1.5, 'Преобразователь частоты 0,37 кВт'],[0.75, 2.3, 'Преобразователь частоты 0,75 кВт'],[1.1, 2.3, 'Преобразователь частоты 1,1 кВт'],[1.5, 3.7, 'Преобразователь частоты 1,5 кВт'],[2.2, 5, 'Преобразователь частоты 2,2 кВт'],[4, 8.5, 'Преобразователь частоты 4 кВт'],[5.5, 12, 'Преобразователь частоты 5,5 кВт'],[7.5, 16, 'Преобразователь частоты 7,5 кВт'],[11, 24, 'Преобразователь частоты 11 кВт'],[15, 30, 'Преобразователь частоты 15 кВт'],[18.5, 37, 'Преобразователь частоты 18,5 кВт'],[22, 45, 'Преобразователь частоты 22 кВт'],[30, 60, 'Преобразователь частоты 30 кВт'],[37, 75, 'Преобразователь частоты 37 кВт']]}
            #     ch = choice[380] if "кварк" in vent.lower() else choice[220] if 'канал-вент' in vent.lower() else choice[int(findall(r'Uпит=~(\d+) В', strochka.replace(",","."))[0])]
            #     for item in ch:
            #         if item[0] >= float(findall(r'Ny=(\d+\.?\d*) кВт', strochka.replace(",","."))[0]) and item[1] > float(findall(r'Iпот=(\d+\.?\d*) A', strochka.replace(",","."))[0]):
            #             itog = item[2]
            #             break
            #         else:
            #             return "Подбор невозможен"
            #         return itog
    st.write("#### ПОДБОР ЧАСТОТНИКОВ ###")
    filess = st.file_uploader("Перетащите сюда бланки", type=("docx", 'doc'), accept_multiple_files=True)
    def to_excel(df, HEADER=False, START=1):
        output = __import__("io").BytesIO()
        writer = ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, header=HEADER,startrow=START, startcol=START, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # format1 = workbook.add_format({'num_format': '0.00'})
        # worksheet.set_column('A:A', None, format1)
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(), len(str(series.name))
            )) + 1
            worksheet.set_column(idx,idx,max_len)
        writer.close()
        return output.getvalue()
    df_chast = []
    with st.expander("Частотники"):
        for item in filess:
            soo = chast_rec(item,fty)
            if soo:
                for item in list(soo):
                    df_chast.append(item.split(" | "))
                    st.write(item)
                    st.markdown("---")
    st.download_button(label='💾 Скачать файл с частотниками',data=to_excel(DataFrame(df_chast, columns=["Имя файла БЗ", "Маркировка частотника", "Количество"]), HEADER=True, START=0) ,file_name= 'частотники.xlsx')
with tabsis[2]:
    __import__("AUTOMATA_WEB").streamlit_version()
