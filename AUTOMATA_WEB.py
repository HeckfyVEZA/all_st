"""Веб-версия программы подбора автоматики. Делает самое основное, что требуется - генерирует бланки, выплёвывая их в заархивированном виде
"""

import streamlit as st
import AUTOMATA as au
from pandas import DataFrame, concat
from vezamodule import ideal_message, xlsx_file_beautifulication, SUPPORTED_EXCTENTIONS_FOR_BLANK
from io import BytesIO
from docx import Document
from zipfile import ZipFile, BadZipFile
from time import perf_counter
from streamlit import session_state
from datetime import datetime
from collections import Counter
from itertools import takewhile

# streamlit run c:\\users\\ovchinnikov\\documents\\python\\21_automata\\automata_web.py
# streamlit run C:\\Users\\ovchinnikov\\Documents\\Python\\21_AUTOMATA\\AUTOMATA_WEB.py
# streamlit run C:\\Users\\krinitsin.da\\Documents\\AUTOMATA\\AUTOMATA_WEB.py
# Копировать, когда надо запустить

def streamlit_version():
    """Веб-версия программы подбора автоматики. Делает самое основное, что требуется - генерирует бланки, выплёвывая их в заархивированном виде
    """

    def create_tab(tab, dictionaty:dict, key:str, key_for_key=list()):
        """Рекурсивная функция по созданию кнопок выбора

        Args:
            tab (_type_): текущая вкладка, остаётся неизменной на всё рекурсивном пути
            dictionaty (dict): словарь, либо со словарями, либо со значениями
            key (str): ключ
            key_for_key (_type_, optional): По умолчанию пустой список путя к текущему значению. В рекурсивно вызванных функциях будет содержать в себе значения
        """

        def set_parameters(value:str):
            """Выставление некоторых параметров прямо по ходу дела

            Args:
                value (str): текущее заданное значение

            Returns:
                str: value - отредактированное значение
            """
            global value_V, value_w, value_I, value_N, honeycomb_index
            
            if 'W[I]' in value:
                col1, col2 = tab.columns(2)
                if code_key != '1.5':
                    with col1:
                        value_w = col1.text_input('Введите мощность, кВт', 0.31 if 'СУ' in new_new_key else 'W', key=new_new_key + '_W', help='Введите числовое значение данного параметра, чтобы программа сработала корректно')
                    with col2:
                        value_I = col2.text_input('Введите силу тока, А', 1.43 if 'СУ' in new_new_key else 'I', key=new_new_key + '_I', help='Введите числовое значение данного параметра, чтобы программа сработала корректно')
                value = value.replace('W', str(value_w)).replace('I', str(value_I))
            if 'EK' in value:
                value_N = tab.text_input('Введите мощность ступени/ей, кВт', 'N', key=new_new_key + '_N', help='Введите числовое значение данного параметра, чтобы программа сработала корректно')
                value = value.replace('N', value_N)
            if 'W[P]' in value:
                value_P = tab.text_input('Введите мощность обогрева клапана, кВт', 'P', key=new_new_key + '_P', help='Введите числовое значение данного параметра, чтобы программа сработала корректно')
                value = value.replace('P', value_P)
                pass
            if value == 'SU':
                honeycomb_index = tab.checkbox('Предусмотреть подключение в ШСАУ поплавкового датчика уровня ПДУ', False, f"{current_code}_honey", 'При наличии сотового увлажнителя этот параметр будет влиять на один из пунктов доптребований в бланке для заказчика') # and au.CODES_FOR_BLANK_WITH_CODES[current_code]['СУ']['Наличие'] == au.CODES_WITH_CODES[current_code]['СУ']['Наличие']['присутствует']
            return value
            pass

        global value_V, value_w, value_I

        key_for_key.append(key)
        if all(isinstance(value, str) for value in dictionaty[key].values()):
            if len(key_for_key) == 1:
                tab.subheader(key)
                
            new_new_key = '_'.join(key_for_key)
            real_value = tab.radio(key, dictionaty[key].keys(), key=new_new_key, horizontal=True, help='; '.join(f"{key}: {value}" for key, value in dictionaty[key].items()))
            code_key = f"1.{CODES_keys.index(key_for_key[0]) + 1}"
            match len(key_for_key):    
                case 1:
                    main_value = set_parameters(au.CODES[key_for_key[0]][real_value])
                    au.CODES_FOR_BLANK[key_for_key[0]] = main_value
                    au.CODES_FOR_BLANK_WITH_CODES[code_key] = main_value
                case 2:
                    main_value = set_parameters(au.CODES[key_for_key[0]][key_for_key[1]][real_value])
                    au.CODES_FOR_BLANK[key_for_key[0]][key_for_key[1]] = main_value
                    au.CODES_FOR_BLANK_WITH_CODES[code_key][key_for_key[1]] = main_value
                case 3:
                    main_value = set_parameters(au.CODES[key_for_key[0]][key_for_key[1]][key_for_key[2]][real_value])
                    au.CODES_FOR_BLANK[key_for_key[0]][key_for_key[1]][key_for_key[2]] = main_value
                    au.CODES_FOR_BLANK_WITH_CODES[code_key][key_for_key[1]][key_for_key[2]] = main_value
            return key_for_key[:-1]
        else:
            if key == 'Насос' and liquid:
                value_V = ''
                # tab.write(dictionaty[key])
                return key_for_key[:-1]
            tab.subheader(key)
            for new_key in dictionaty[key].keys():
                key_for_key = create_tab(tab, dictionaty[key], new_key, key_for_key if len(key_for_key) > 1 else [key])
            
            return key_for_key[:-1]

    def final_part():
        """Заключительная часть программы, где обрабатываются все созданные бланки, которые пока существуют в виде док-объектов и словарей с запутанными структурами
        """

        with st.form(f'additional_work_{postfix}'):
            result_zip_file = BytesIO()
            
            count, all_files = 0, 0
            for filename in session_state['unusual_files'].keys():
                if session_state['unusual_files'][filename][0] is None:
                    st.write(filename)
                    continue
                all_files += 1

                if session_state['unusual_files'][filename][1]['JTU_blank'][0] or session_state['unusual_files'][filename][1]['glycol_blank'][0]:
                    st.write('Обрабатываемый файл: ' + filename.split('\\')[-1])

                if session_state['unusual_files'][filename][1]['JTU_blank'][0]:
                    session_state['unusual_files'][filename][1]['JTU_blank'][1] = st.text_input('Введите номер бланк-заказа для жидкостного теплоутилизатора ЖТУ', key=f'jtu_{filename}_{postfix}', help='Сюда надо ввести номер бланк-заказа для жидкостного теплоутилизатора ЖТУ')

                if session_state['unusual_files'][filename][1]['glycol_blank'][0]:
                    for i in range(len(session_state['unusual_files'][filename][1]['glycol_blank'][1])):
                        session_state['unusual_files'][filename][1]['glycol_blank'][1][i] = st.text_input(f"Введите номер бланк-заказа для водосмесительного узла ТО{i+1}", key=f"to_{filename}_{i+1}_{postfix}", help='Сюда надо ввести номер бланк-заказа для водосмесительного узла')

            progress_bar_2 = st.progress(count, 'Запустите финальную обработку файлов, чтобы загрузить бланки')

            if st.form_submit_button('Запустить финальную обработку файлов', 'Программа не даст скачать бланки, пока не будет нажата эта кнопка! Заполнять все поля, к слову, необязательно. Но желательно'):
                time_start = perf_counter()
                for filename, value in session_state['unusual_files'].items():
                    if session_state['unusual_files'][filename][0] is None:
                        st.write(filename)
                        continue

                    for paragraph in session_state['unusual_files'][filename][0].paragraphs:
                        if session_state['unusual_files'][filename][1]['JTU_blank'][0]:
                            if 'Предусмотрено управление водосмесительным узлом ЖТУ' in paragraph.text:
                                paragraph.text = paragraph.text.replace('НОМЕР_БЛАНКА', session_state['unusual_files'][filename][1]['JTU_blank'][1])
                        
                        if session_state['unusual_files'][filename][1]['glycol_blank'][0]:
                            for i in range(len(session_state['unusual_files'][filename][1]['glycol_blank'][1])):
                                if f'Предусмотрено управление водосмесительным узлом ТО{i+1}' in paragraph.text:
                                    paragraph.text = paragraph.text.replace(f'НОМЕР_БЛАНКА_{i}', session_state['unusual_files'][filename][1]['glycol_blank'][1][i])

                    bio_docx = BytesIO()
                    session_state['unusual_files'][filename][0].save(bio_docx)
                    with ZipFile(result_zip_file, mode='a') as archive:
                        archive.writestr(filename.split('\\')[-1], bio_docx.getvalue())

                    count += 1
                    progress_bar_2.progress(count / all_files, ideal_message(count, all_files, 'файлов', time_start, True, True))

                with ZipFile(result_zip_file, mode='a') as archive:
                    archive.writestr(au.FILE_PIVOT_TABLE, xlsx_file_beautifulication(BytesIO(), session_state['pivot_table']).getvalue())

                result_zip_name = f"{datetime.today().strftime('%d-%m-%y %H-%M-%S')}.zip"
                st.write('Можно загружать бланки')

        try:
            st.download_button('Загрузите бланки', result_zip_file.getvalue(), result_zip_name, 'application/zip', help='Файлы будут выгружены в zip-архив')
        except NameError:
            st.write('Запустите финальную обработку файлов, чтобы загрузить бланки')
        pass

    #=============================================================================================================================================================================================

    global value_V, value_w, value_I, honeycomb_index

    st.title(f'Подбор автоматики v{au.VERSION}', help='Это программа подбора автоматики. Есть два режима работы - автоматический (загружается бланк, и на его основе программа автоматически рассчитывает бланки КА) и ручной (параметры для КА вводятся в программе - возможно даже создание бланков КА без исходного бланка автоматики, в настоящий режим ручной режим находится в стадии разработки)')
    col1, col2, col3 = st.columns(3)
    with col1:
        username = col1.text_input('Введите имя пользователя для начала работы', key='username', help='Данное поле абсолютно обязательно для ввода - программа не запустится без имени пользователя!')
    with col2:
        start_ka_number = datetime.today().year % 2000 * 10000000
        ka_number = col2.number_input('Введите номер бланк-заказа', start_ka_number, start_ka_number + 9999999, start_ka_number, 1, key='ka_number_number', help='Номер бланк-заказа, вводить, если это не ВЕРОСА. Код должен быть девятизначным, первые две цифры означают год, и за вас уже введены. Необязательное поле!')
    with col3:
        filial = col3.radio('Выберите отдел', ('ОПР', 'СПБ', 'ННВ', 'ДОН'), key='filial', help='Выбор филиала. По умолчанию поставлено то, где развёрнута эта программа', horizontal=True)
    
    CODES_keys = tuple(au.CODES.keys())
    if 'username' not in session_state:
        session_state['username'] = username
    if 'ka_number' not in session_state:
        session_state['ka_number'] = ka_number
    if 'run' not in session_state:
        session_state['run'] = 0
    if 'unusual_files' not in session_state:
        session_state['unusual_files'] = dict()
    if 'uploaded_files' not in session_state:
        session_state['uploaded_files'] = list()
    if 'pivot_table' not in session_state:
        session_state['pivot_table'] = DataFrame(columns=au.PIVOT_TABLE_COLUMNS[1:])

    if 'add_it' not in session_state:
        session_state['add_it'] = dict()
    if 'remove_it' not in session_state:
        session_state['remove_it'] = dict()

    if 'all_results' not in session_state:
        session_state['all_results'] = {f'1.{i}' : [] for i in range(1, 17)}
    if 'all_results_strings' not in session_state:
        session_state['all_results_strings'] = {f'1.{i}' : '' for i in range(1, 17)}
    if 'amount' not in session_state:
        session_state['amount'] = {f'1.{i}' : 0 for i in range(1, 17)}
    if 'automation_devices' not in session_state:
        session_state['automation_devices'] = {f'1.{i}' : [] for i in range(1, 17)}
    if 'circuit_design ' not in session_state:
        session_state['circuit_design'] = list()

    if 'liquid' not in session_state:
        session_state['liquid'] = 0
    if 'vector_name' not in session_state:
        session_state['vector_name'] = {
            'нагревателя' : [],
            'охладителя' : []
        }
    if 'illumination_theory' not in session_state:
        session_state['illumination_theory'] = {  
            "1.4" : [],
            "1.12" : [],
            "1.13" : []
        }
    honeycomb_index = False
    all_valve_drives = ('M24-V', 'M24-S-V', 'M230-V', 'M230-S-V', 'M24-SR-V', 'M24-SR-S2-V', 'M230-SR-V', 'M230-SR-S2-V', 'F24-V', 'F24-S-V', 'F230-V', 'F230-S-V')

    with st.expander("Раскройте, чтобы ввести параметры самостоятельно", expanded=False):
        for k in (0, 1):
            start, end = 8 * k, 8 * (k + 1)
            for tab, key, j in zip(st.tabs([f"1.{i}. {codes_key}" for i, codes_key in enumerate(CODES_keys[start:end], start + 1)]), CODES_keys[start:end], range(start + 1, end + 1)):
                current_code = f"1.{j}"
                if j in (5, 8):
                    if j == 5:
                        liquid = tab.slider('Укажите процентное содержание пропиленгликоля/этилегликоля', 0, 100, 0, 10, key=current_code + '_liquid', help='0 - только вода, 0-40 - формируется нестандартный Вектор, 40-100 - указывать вручную!')
                        if liquid != session_state['liquid']:
                            session_state['liquid'] = liquid
                            au.CODES_FOR_BLANK_WITH_CODES[current_code] = au.linear_dict_change_values(au.CODES_FOR_BLANK_WITH_CODES[current_code], '', True)
                        # tab.write(liquid)
                        if liquid and liquid > 40:
                            st.write('Высокое содержание пропиленгликоля/этиленгликоля. Заполняйте это поле вручную!')
                            continue
                        schema = 2 if liquid == 0 else 3
                    if j == 8:
                        schema = 3

                    col1, col2 = tab.columns(2)
                    with col1:
                        value_Gh = col1.number_input("Введите расход теплоносителя, м³/ч", 0.0, 60.0, 0.0, 0.1, key=current_code + '_Gh')
                    with col2:
                        value_ctrl_dev = 'С' if schema == 3 else col2.radio('Выберите тип клапана (сидельный или шаровой)', ('С', 'Ш'), key=current_code + '_sh', horizontal=True)
                    
                    for data in au.VECTOR_DATA:
                        if eval(data[5].replace('G', str(value_Gh)).replace(',', '.')) and data[2] == value_ctrl_dev and data[0] == schema:
                            if schema == 2:
                                value_V, value_w, value_I = data[9], data[11], data[10]
                            else:
                                value_V, value_w, value_I = 'V', 'W', 'I'
                            vector_name = f"ВЕКТОР-{schema}-{value_ctrl_dev}-{data[3]}-П(Л)-С+"
                            tab.write(vector_name)
                            break
                    else:
                        value_V, value_w, value_I = 'V', 'W', 'I'
                create_tab(tab, au.CODES, key, list())

                real_value = au.the_meaning(current_code, au.CODES_FOR_BLANK_WITH_CODES[current_code])
                if j in (1, 2, 15):
                    match len(session_state['all_results'][current_code]):
                        case 0:
                            session_state['all_results'][current_code].append(real_value)
                        case 1:
                            session_state['all_results'][current_code][0] = real_value
                else:
                    if j in (4, 12, 13):
                        illumination_theory = tab.checkbox('Добавить освещение в этот отсек', False, f"{current_code}_light", 'В клапанах и вентиляторах возможно наличие освещения')
                        pass
                    if j == 3:
                        # tab.write(au.CODES_FOR_BLANK_WITH_CODES[current_code])
                        # tab.write(all_valve_drives)
                        con1 = '230' if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Питание привода'] == au.CODES_WITH_CODES[current_code]['Питание привода']['питание 230В АС'] else '24'
                        con2 = ('-SR-', ) if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Управление приводом'] == au.CODES_WITH_CODES[current_code]['Управление приводом']['аналоговое управление (0... 10В)'] else ('-S-', '-S2-') if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Управление приводом'] == au.CODES_WITH_CODES[current_code]['Управление приводом']['дискретное управление'] else ''
                        con3 = '-S-' if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Дополнительные параметры'] == au.CODES_WITH_CODES[current_code]['Дополнительные параметры']['наличие одного, встроенного в привод выключателя положения с переключающим «сухим» контактом'] else '-S2-' if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Дополнительные параметры'] == au.CODES_WITH_CODES[current_code]['Дополнительные параметры']['наличие двух, встроенных в привод выключателей положения с переключающими «сухими» контактами'] else au.CODES_WITH_CODES[current_code]['Дополнительные параметры']['отсутствует']
                        valve_drive = tab.radio('Выберите клапан из возможных при данной конфигурации', tuple(filter(lambda x:(con1 in x) and (any(s in x for s in con2) if con2 else True) and (con3 in x if con3 != au.CODES_WITH_CODES[current_code]['Дополнительные параметры']['отсутствует'] else True), all_valve_drives)), key=f"{current_code}_valve_drive", help='На основе выбранных параметров предполагается, какой клапан будет использоваться. Если выбор был сделан неправильно, сообщите об этом или впишите сами, нажав на кнопочку чуть выше', horizontal=True) if not tab.checkbox('Вписать обозначение привода вручную', False, f"{current_code}_valve_drive_choice", 'Если хотите вписать маркировку бланка руками, то нажмите сюда') else tab.text_input('Впишите маркировку привода клапана', key=f"{current_code}_valve_drive_name", help='Впишите маркировку привода клапана. Постарайтесь вписать правильно!')
                        # tab.write(valve_drive)

                    if j == 14:
                        all_sensors = au.for_1_14(session_state['all_results_strings'], session_state['amount'], False)
                        all_devices = au.create_automation_devices(current_code, session_state['amount'][current_code], all_sensors)
                        all_sensors = au.INNER_DELIMITER.join(sensor for sensor in all_sensors[:-1] if sensor)
                        if all_sensors:
                            if len(session_state['all_results'][current_code]) == 0:
                                session_state['all_results'][current_code].append(all_sensors)
                        else:
                            session_state['all_results'][current_code] = []
                    if j == 16:
                        if len(session_state['all_results'][current_code]) == 0:
                            session_state['all_results'][current_code].append(real_value)
                    tab.markdown('---')

                    col = tab.columns(8)
                    with col[0]:
                        session_state['add_it'][current_code] = col[0].button('Добавить', f"{current_code}_add")
                    with col[1]:
                        session_state['remove_it'][current_code] = col[1].button('Убрать', f"{current_code}_remove")

                    if session_state['add_it'][current_code]:
                        if real_value:
                            if j in (6, 10, 11):
                                berbers = ('ПО', 'ГН', 'ДН') if j == 6 else ('ТР', 'ТП', 'ЖТУ') if j == 10 else ('ФУ', 'СУ', 'ПУ')
                                aburval = ('', '', '') if j == 6 else ('', '', '') if j == 10 else ('', '', '')
                                for part_value, part_berb_part_abur in zip(real_value.split(au.INNER_DELIMITER), ((au.CODES_FOR_BLANK_WITH_CODES[current_code][bebe]['Наличие'], abur) for bebe, abur in zip(berbers, aburval) if au.CODES_FOR_BLANK_WITH_CODES[current_code][bebe]['Наличие'] != '-')):
                                    session_state['all_results'][current_code].append((part_value, au.create_circuit_design(current_code, len(session_state['all_results'][current_code]), session_state['amount'][current_code] + 1, *part_berb_part_abur)))
                            else:
                                match j:
                                    case 3:
                                        the_args = (
                                            au.CODES_FOR_BLANK_WITH_CODES[current_code]['Типы клапанов'],
                                            au.CODES_FOR_BLANK_WITH_CODES[current_code]['Обогрев клапана'] if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Обогрев клапана'] != '-' else '', valve_drive)
                                    case 4:
                                        session_state['illumination_theory'][current_code].append(illumination_theory)
                                        the_args = tuple()
                                    case 5:
                                        the_args = (value_V[0],)
                                        session_state['vector_name']['нагревателя'].append(vector_name)
                                    case 7:
                                        the_args = (value_N, au.CODES_FOR_BLANK_WITH_CODES[current_code]['Питание'])
                                    case 8:
                                        session_state['vector_name']['охладителя'].append(vector_name)
                                        the_args = tuple()
                                    case 12 | 13:
                                        session_state['illumination_theory'][current_code].append(illumination_theory)
                                        the_args = (
                                            'АВ*п' if au.CODES_FOR_BLANK_WITH_CODES[current_code]['Резервирование'] != '-' else f"В{'п' if j == 12 else 'в'}",
                                            f"(Nу={value_w}кВт; Iпот={value_I}А; ~{au.CODES_FOR_BLANK_WITH_CODES[current_code]['Параметры двигателя']['Питание'][-2]})")
                                    case _:
                                        the_args = tuple()
                                session_state['all_results'][current_code].append((real_value, au.create_circuit_design(current_code, len(session_state['all_results'][current_code]), session_state['amount'][current_code] + 1, *the_args)))
                            # session_state['circuit_design'].append()
                            # tab.write()

                    if session_state['remove_it'][current_code]:
                        shorting_list = lambda li:li[:-1] if len(li) > 0 else li
                        session_state['all_results'][current_code] = shorting_list(session_state['all_results'][current_code])
                        if j in (5, 8):
                            bebe = 'нагревателя' if j == 5 else 'охладителя'
                            session_state['vector_name'][bebe] = shorting_list(session_state['vector_name'][bebe])
                        if j in (4, 12, 13):
                            session_state['illumination_theory'][current_code] = shorting_list(session_state['illumination_theory'][current_code])

                if j == 4:
                    all_filters = Counter(val[0] for val in session_state['all_results'][current_code])
                    session_state['all_results_strings'][current_code] = au.INNER_DELIMITER.join(f"{all_filters[key] if all_filters[key] > 1 else ''}{key}" for key in sorted(all_filters.keys(), reverse=True)) if all_filters else '0'
                    pass
                elif j in (1, 2, 14, 15, 16):
                    session_state['all_results_strings'][current_code] = au.INNER_DELIMITER.join(val for val in session_state['all_results'][current_code]) if session_state['all_results'][current_code] else '0'
                else:
                    session_state['all_results_strings'][current_code] = au.INNER_DELIMITER.join(val[0] for val in session_state['all_results'][current_code]) if session_state['all_results'][current_code] else '0'
                session_state['amount'][current_code] = len(au.INNER_DELIMITER.join(val[0] for val in session_state['all_results'][current_code]).split(au.INNER_DELIMITER)) if session_state['all_results'][current_code] else 0

                if session_state['all_results_strings'][current_code] != '0':
                    match j:
                        case 3:
                            # if and 
                            pass
                        case 4 | 12 | 13:
                            session_state['automation_devices'][current_code] = au.create_automation_devices(current_code, session_state['amount'][current_code], tab.radio('Выбор диапазона работы реле перепада', ['30-300Па', '50-500Па', '100-1500Па'], key=f"{current_code}_perepad", horizontal=True))
                        case 5:
                            session_state['automation_devices'][current_code] = au.create_automation_devices(current_code, session_state['amount'][current_code], liquid, tab.radio('Капилляры', (s[0] for s in Counter(s[2] for s in au.SHLANG_DATA).most_common()), horizontal=True), 2)
                        case 10:
                            session_state['automation_devices'][current_code] = au.create_automation_devices(current_code, session_state['amount'][current_code], '30-300Па', session_state['all_results_strings'][current_code])
                        case 14:
                            session_state['automation_devices'][current_code] = all_devices
                else:
                    session_state['automation_devices'][current_code] = []

                if j == 3:
                    recycling_split = tab.checkbox('Предусмотреть выбор алгоритма работы камеры смешения в меню контроллера: «Температура в камере смешения контролируется автоматически путем регулирования степени открытия рециркуляционного и воздухозаборного клапанов по датчику, установленному в камере смешения»; «Степень открытия воздухозаборного и рециркуляционного клапанов задается вручную для каждого сезонного режима, датчик в камере смешения необходим только для наблюдения за показаниями»', False, f"{current_code}_rec", 'Имеет смысл, если') if session_state['amount'][current_code] >= 2 and au.CODES_WITH_CODES[current_code]['Типы клапанов']['рециркуляционный клапан'] in session_state['all_results_strings'][current_code] else False
                    pass
                if j == 7:
                    air_preparation_unit = tab.checkbox('Предусмотреть возможность выбора работы первого нагревателя: «Включение совместно с открытием клапанов», «Включение совместно с вентилятором» (заводская настройка)', False, f"{current_code}_apu", 'Имеет смысл включать данную функцию, если у вас два и более электрокалорифера, тогда один из них, скорее всего, будет работать в данном режиме') if session_state['amount'][current_code] >= 2 else False

                tab.text(f"Текущее выставленное значение: {real_value}", help='Здесь отображается, что выбрано на основании того, как проставлены галочки в меню выбора')
                tab.text(f"Текущая строка по данному полю: {session_state['all_results_strings'][current_code]}", help='Здесь отображается, что уже внесено в итоговые поля и будет использоваться при формировании бланков')
    st.markdown('---')

    if session_state['ka_number'] != ka_number:
        session_state['run'] = 0
        session_state['pivot_table'] = DataFrame(columns=au.PIVOT_TABLE_COLUMNS[1:])
        session_state['ka_number'] = ka_number
    
    conditions = (
            st.checkbox('Использовать параметры, выбранные вручную', False, 'use', 'Если будет выбрана эта опция, то в бланк-заказ будут внесены те параметры, что выбраны выше'),
            st.checkbox('Добавлять КИП при наличии ШСАУ', True, 'kip_shsau', 'Пока выбран этот параметр, в бланк-заказы будут добавлять все КИПы'),
            st.checkbox('Освещение', True, 'illumi', 'Я же правильно понимаю, что этот параметр означает, что нам надо включить освещение в бланк-заказы?'),
            st.radio('Выберите контроллер', ('ОВЕН ПР200', 'Zentec M245', 'нет'), help='От наличия того или иного типа контроллера пока ничего не зависит, но если его не будет, об этом будет пометка', horizontal=True))

    tab_automated, tab_byhanded = st.tabs(('Создать бланки автоматики на основе введённых бланков', 'Создать бланки на основе выбранных выше параметров БЕЗ бланков'))
    with tab_automated:
        st.subheader('Автоматический подбор комплекта автоматики на основе введённых бланков', help='')
        postfix = 'tab_automated'

        
        # conditions = (False, True, True)
        
        uploaded_files = st.file_uploader("Перетащите сюда бланки для обработки", SUPPORTED_EXCTENTIONS_FOR_BLANK[:-1], True, help='Программа не заработает, пока как минимум не будут перетащены сюда файлы!')
        count, all_files = 0, len(uploaded_files)
        progress_bar = st.progress(count, 'Введите имя и загрузите файл(ы)')

        if username:
            progress_bar.progress(count, 'Загрузите файл(ы)')
            if uploaded_files:
                if session_state['uploaded_files'] != uploaded_files:
                    session_state['uploaded_files'] = uploaded_files
                    session_state['run'] = 0

                if session_state['run'] == 0:
                    session_state['unusual_files'] = dict()
                    time_start = perf_counter()

                    for file in uploaded_files:
                        # st.write(file)
                        try:
                            doc = Document(file)
                        except (BadZipFile, ValueError):
                            st.write(f'Преобразуйте файл {file.name} в docx-формат, после чего загрузите ещё раз!')
                        except Exception as err:
                            st.write(err)
                        else:
                            try:
                                auf, acps = au.main_part(file.name, (doc, username, conditions, ka_number, filial), streamlit_version=True)
                            except Exception as err:
                                st.write(err)
                            else:
                                # st.write(auf, acps)
                                session_state['unusual_files'] |= auf
                                session_state['pivot_table'] = concat([session_state['pivot_table'], acps], ignore_index=True).drop_duplicates().reset_index(drop=True)
                            pass
                        count += 1
                        ka_number += 1
                        # session_state['ka_number'] = ka_number
                        progress_bar.progress(count / all_files, ideal_message(count, all_files, 'файлов', time_start, True, True))
                else:
                    progress_bar.progress(100/100, 'Бланки отработаны')
                session_state['run'] += 1

                final_part()

            else:
                session_state['run'] = 0
                session_state['pivot_table'] = DataFrame(columns=au.PIVOT_TABLE_COLUMNS[1:])
                pass
    
    with tab_byhanded:
        st.subheader('Автоматический подбор комплекта автоматики на основе введённых параметров БЕЗ бланков', help='')
        postfix = 'tab_byhanded'

        session_state['circuit_design'] = [s[1] for sss in tuple(session_state['all_results'].values())[2:-3] for s in sss if sss]
        circuit_design_long = '-'.join(sss[1] for sss in session_state['circuit_design'] if sss)
        circuit_design_shot = '-'.join(''.join(s for s in takewhile(lambda sx: sx != '(', sss[1])) for sss in session_state['circuit_design'] if sss)
        
        automation_devise = []
        for ad in session_state['automation_devices'].values():
            automation_devise += ad
        # automation_devise

        main_info = (
            st.text_input('Введите название объекта', key='object_name', help='Собственно, здесь от вас просят просто ввести название объекта'),
            st.text_input('Введите название системы', key='system_name', help='При записи системы придерживайтесь следующих правил: Если у вас не-вероса и несколько систем, то записывайте их через запятую с пробелом, а если несколько подряд идущих систем, то записывайте через дефис (можно слитно, но если удастся отделить дефис пробелами, вам цены не будет!)'),
            st.text_input('Введите название организации', key='organi_name', help=''),
            st.text_input('Введите фамилию и имя менеджера', key='manage_name', help='')
        )

        # main_info

        with st.expander("Проверить выбранные параметры", expanded=False):
            cols = st.columns(3)
            with cols[0]:
                st.write(session_state['all_results_strings'])
                st.write(session_state['amount'])
            with cols[1]:
                st.write(automation_devise)
            with cols[2]:
                st.text(circuit_design_long)
                st.text(circuit_design_shot)
                st.write(session_state['vector_name'])
                st.write(session_state['illumination_theory'])
                st.write(honeycomb_index)
                st.write(air_preparation_unit)
                st.write(recycling_split)

        if username and all(bool(mi) for mi in main_info):
            auf, acps = au.main_part(
                '',
                (None, username, conditions, ka_number, filial),
                (session_state['all_results_strings'], automation_devise, circuit_design_long, circuit_design_shot, session_state['amount']),
                (session_state['vector_name'], session_state['illumination_theory'], honeycomb_index, air_preparation_unit, recycling_split),
                main_info,
                streamlit_version=True)
            session_state['unusual_files'] |= auf
            session_state['pivot_table'] = concat([session_state['pivot_table'], acps], ignore_index=True).drop_duplicates().reset_index(drop=True)
            
            final_part()


# streamlit run c:\\users\\ovchinnikov\\documents\\python\\21_automata\\automata_web.py
# streamlit run C:\\Users\\ovchinnikov\\Documents\\Python\\21_AUTOMATA\\AUTOMATA_WEB.py
# streamlit run C:\\Users\\krinitsin.da\\Documents\\AUTOMATA\\AUTOMATA_WEB.py
# Копировать, когда надо запустить

if __name__ == '__main__':
    st.set_page_config(layout="wide")
    streamlit_version()
