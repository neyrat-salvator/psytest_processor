from bs4 import BeautifulSoup
from selenium import webdriver
import excel_process


# Метод для некоторых глобальных атрибутов
class attributes:
    
    # Глобальные значения, которыми можно пользоваться
    global url, driver, filename, sheet_name, workbook, sheet_data
    

# Метод для открытия браузера, из которого осуществляется вход на psytests.org
def open_browser():
    
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "\
        "AppleWebKit/537.36 (KHTML, like Gecko) "\
        "Chrome/61.0.3163.100 Safari/537.36")
    options.add_argument("--disable-blink-features")
    options.add_argument('--disable-blink-features=AutomationControlled') 
    attributes.driver = webdriver.Chrome(options=options)
    attributes.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")


# Метод для выгрузки данных из ссылки на анкету
def get_content_from_url(url):
    
    attributes.driver.get(url)
    response = attributes.driver.page_source
    soup = BeautifulSoup(response, 'html.parser')
    content = soup.find_all(['td', 'div'], class_=['nisTitle', 'nisName', 'nisVal'])

    return content


# Получить HTML-контент в виде списка
def convert_html(data):
    
    # Получить и обработать данные
    soup = BeautifulSoup(data, 'html.parser')
    tags_content = soup.find_all(['td', 'div'], class_=['nisTitle', 'nisName', 'nisVal'])
    # content = BeautifulSoup(tags_content, 'html.parser').text.replace('[', '').replace(']', '').split(', ')
    content = []
    
    for values in tags_content:
        content.append(values.string)
    
    # Проверка и удаление лишних значений. Удаляются значения, но сами индексы в списке не удаляются. Требует доработки
    # exception = ['Основные шкалы (станайны)', 'Субшкалы (средний балл)']
    # cycle_number = 0
    # for item in content:
    #     if item in exception:
    #         content.pop(cycle_number)
    #     else:
    #         cycle_number += 1
    #         continue
    
    return content
    
    
# Метод для прохода по всем строкам в столбце для записи значений из psytests.org
def write_user_values(sheet, columns: list):
    
    cycle_column = 0
    
    for column in columns:
    
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True):
            
            if row[0] == '№':
                continue
            
            row_number = row[0]
            resource = row[column-2]
            content = get_content_from_url(url=resource)
            excel_process.set_value_to_cells(sheet=sheet, 
                                            target_column=columns[column-2], 
                                            row_index=row_number, 
                                            value=content)
            
        cycle_column += 1
    
    
# Метод для заполнения всех таблиц (анкет) из готовых данных лсита Main
def set_tables_form(source_excel, forms: list, data_columns: list):
    
    source_excel_sheet = source_excel.get('sheet')
    # Эта переменная хранит в себе индекс обрабатываемов анкеты
    cycle_form_number = 0
    
    # В этом цикле обрабатываются анкеты вместе с листами под них
    for form in forms:
        
        form_sheet = form.get('sheet')
        form_columns = form.get('columns')
        # В этой переменной отсчитывается максимальное число строк, которые будут обрабатываться
        list_of_rows = source_excel_sheet.iter_rows(min_row=1, 
                                                    max_row=source_excel_sheet.max_row, 
                                                    values_only=True)
    
        # В этом цикле обрабатываются все строки главной таблицы с исходными данными
        for row in list_of_rows:
            
            # Эта переменная ссылается на значения из первого столбца, отображающие номер анкетируемого
            row_number = row[0]
            
            # Здесь я пытался отследить, почему последняя строка не обрабатывалась. Пока не получилось
            # if row_number in (1, 48, 49):
            #     True
            
            # Этой проверкой пропускается первая строка на листе, где указаны наименования с толбцов
            if row[0] == '№':
                continue
            
            # В этой переменной хранится список значений, которые будут вписаны в таблицы на листах под анкеты
            data_forms = convert_data_forms(sheet=source_excel_sheet, 
                                            target_column=data_columns[cycle_form_number], 
                                            row_number=row_number)
            # Эта переменная отсчитывает индекс каждого обрабатываемого значения
            cycle_value_number = 0
            
            # В этом цикле обрабатываются все значения из анкеты
            for value in data_forms:
                
                # Эта переменная высчитывает каждое нечетное значение для записи их в талицы под анкеты
                column_number = form_columns[int((cycle_value_number-1)/2)]
                
                # Первое значение пропускается для недопущения отрицательного числа, оно все равно не обрабатывается
                if cycle_value_number == 0:
                    cycle_value_number += 1
                    continue
                
                # Эта проверка запускает запись значения в таблицу под анкету, когда занчение в цикле корректное
                if (cycle_value_number % 2) != 0:
                    excel_process.set_value_to_cells(sheet=form_sheet, 
                                                     target_column=column_number, 
                                                     row_index=row_number, 
                                                     value=value)
                    
                cycle_value_number += 1
            
        cycle_form_number += 1
    
# Метод для формирования списка из форм по каждому анкетируемому
def convert_data_forms(sheet, target_column: int, row_number: int):
    
    # if row_number == 1:
    #     return
    
    content = excel_process.get_value_from_cell(sheet=sheet, target_column=target_column, row_index=row_number)
    parse_content = convert_html(data=content)
    
    return parse_content    


# Основной исполняемый метод
def main():
    
    # Имя файла, который требуется обработать
    attributes.filename = 'user_form.xlsx'
    # Имена листов с исходными данными. Под каждый лист/анкету нужна своя запись в этом списке
    attributes.sheet_name = ['Main', 'Form_1', 'Form_2', 'Form_3']
    # Объявление переменных Excel. Под каждый лист/анкету нужная своя переменная Form_[номер листа]
    attributes.workbook = excel_process.get_workbook(filename=attributes.filename)
    Main = excel_process.get_sheet_data(workbook=attributes.workbook, 
                                         sheet_name=attributes.sheet_name[0])
    Form_1 = excel_process.get_sheet_data(workbook=attributes.workbook, 
                                         sheet_name=attributes.sheet_name[1])
    Form_2 = excel_process.get_sheet_data(workbook=attributes.workbook, 
                                         sheet_name=attributes.sheet_name[2])
    Form_3 = excel_process.get_sheet_data(workbook=attributes.workbook, 
                                         sheet_name=attributes.sheet_name[3])
    # Под каждый лист/анкету нужна своя запись в этом списке
    all_excel = [Form_1, Form_2, Form_3]
    # Под каждый столбец с сырыми исходными данными с ресурса psytests.org нужа последовательная своя запись в этом списке
    data_columns = [7, 9, 11]
    
    # Блок для первичной выгрузки данных с сайта
    # open_browser()
    write_user_values(sheet=Main.get('sheet'), columns=data_columns)
    
    # Блок для формирования таблицы
    set_tables_form(source_excel=Main, 
                    forms=all_excel, 
                    data_columns=data_columns)
    
    # Этот исполняемый метод использоваля для точечной выгрузки данных последней анкеты, так как почему-то последняя строка игнорируется
    # convert_data_forms(sheet=Main.get('sheet'), target_column=data_columns[2], row_number=48)
    
    # Блок для одиночной выгрузки данных из ресурсы
    # content = get_content_from_url('https://psytests.org/result?v=emqB7q6-C&b=51_B9x6DwWNjUJ')
    
    # Сохранение и закрытие прогресса
    excel_process.save_and_close(filename=attributes.filename, workbook=attributes.workbook)  


main()