import collections
import openpyxl
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#Consts. Months -> Numbers
months_words = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
months_numbers = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

##Consts. The list of pharmacies' names
apteki_list = ['Izumrud', 'Ostojenka', 'Michur', 'Redmayak_4k1', 'Samark', 'Poletaeva',
               'Prechistenka', 'MalayaDmitrovka', 'Sherbkina', 'Kommunarka',
               'Chelyabinskaya', 'Tverskaya', 'Sh_Entuziastov', 'Lusinovskaya',
               'Blvr_Yana_Rajninsa', 'Barklaya', 'Warshavskoe_shosse', 'Kornejchuka',
               'Rudnevka', 'Julebinskiy_blvr', 'Nijnaya_Krasnoselksaya', 'Kosmonavtov',
               'Sholakskogo', 'Kashirskoe_shosse', 'Redmayak_9', 'Kastanaevskaya', 'Birulovskaya',
               'Kotelniki_Kuzminksaya', 'Ramenskoe_Vokzalnaya_sqr', 'Lubertsi_3-e_Pochtovoe', 'Dolgoprudnij_Moskovskaya']


apteki_rus_address = ['UM001, ул. Изумрудная, д. 18', 'UM002, ул. Остоженка, д.40/1', 'UM004, Мичуринский пр-т, д. 27',
                      'ул. Красного Маяка, д. 4к1', 'UM007, Самарканский б-р, д. 18/26', 'UM008, ул. Ф. Полетаева, д. 40',
                      'UM009, ул. Пречистенка, д. 25', 'UM010,  ул. Малая Дмитровка, д. 15',
                      'UM011, г. Щербинка, ул. Барышевская роща, д.10', 'UM012, пос. Коммунарка, ул. Липовый парк, д. 4к1',
                      'UM013,ул. Челябинская, д. 21', 'UM015, ул. Тверская, д.19/31',
                      'UM016, ш. Энтузиастов, д. 22/18', 'UM017, ул. Люсиновская, д. 36/50',
                      'UM018, б-р Яна Райниса, д. 8', 'UM019, ул.Барклая, д.12',
                      'UM021, Варшавское ш., д.34', 'UM022, ул. Корнейчука, д. 36',
                      'UM023, ул. Рудневка, д. 4', 'UM024, Жулебинский б-р, д. 9',
                      'UM025, ул. Нижняя Красносельская, д. 45/17', 'UM026, ул. Космонавтов, д. 15',
                      'UM027, ул. Шокальского, д. 25а', 'UM028, Каширское ш., д. 53, корп. 4',
                      'UM029, ул. Красного Маяка, д. 9', 'UM030, ул. Кастанаевская, д. 42, корп. 2',
                      'UM033, ул. Бирюлевская, д. 13, корп. 4',
                      'UO003, МО, г. Котельники, ул. Кузьминская, д. 17', 'UO014, МО, г. Раменское, Вокзальная пл., д. 4б',
                      'UO031, МО, г. Люберцы, ул. 3-е Почтовое отделение д. 49, корп. 2',
                      'UO032, МО, г. Долгопрудный, ул.Московская, д. 56, корп. 3']

#Apteki dict
apteki = collections.OrderedDict()

for apteka in apteki_list:
    apteki[apteka] = {}

# Creating list for storing all the Pharmacy coordinates
for k in apteki.keys():
    apteki[k]['name_occurrence_coor'] = []

# Adding rus_addresses
for counter, apteka in enumerate(apteki.keys()):
    apteki[apteka]['rus_address'] = apteki_rus_address[counter]

def open_xlsx(xlsx):
    # Open book
    book = openpyxl.load_workbook(xlsx)
    worksheet = book.active
    return worksheet

def identify_xlsx_active_range(worksheet):
    # Creating the Excel Columns-Rows range OrderDict
    excel_range = collections.OrderedDict()
    for row in worksheet.iter_rows():
        for cell in row:
            # Saving the non-full Columns letters and Rows range in the excel_range dict
            if cell.coordinate[0] not in excel_range:
                # [0] - a letter(Column), i.e. A [1:] - the int
                excel_range[cell.coordinate[0]] = [int(cell.coordinate[1:])]
            else:
                excel_range[cell.coordinate[0]].append(int(cell.coordinate[1:]))
    columns_list = list(excel_range.keys())
    min_value_of_range = min(excel_range[columns_list[0]])
    max_value_of_range = max(excel_range[columns_list[-1]])
    active_cells = worksheet[min_value_of_range:max_value_of_range]
    return active_cells

def extract_xlsx_pharmacies_coordinates(active_cells):
    for row in active_cells:
        for cell in row:
            # Searching for pharmacy names and their coordinates
            if isinstance(cell.value, str):
                for key, value in apteki.items():
                    if key == 'Redmayak_4k1' or key == 'Redmayak_9':
                        if fuzz.token_sort_ratio('Красного Маяка,', cell.value) > 50:
                            templist = []
                            for character in list(cell.value):
                                try:
                                    templist.append(int(character))
                                except:
                                    pass
                            if templist[-1] == 9:
                                apteki['Redmayak_9']['name_occurrence_coor'].append(cell.coordinate)
                            else:
                                apteki['Redmayak_4k1']['name_occurrence_coor'].append(cell.coordinate)
                    else:
                        if fuzz.token_sort_ratio(value['rus_address'], cell.value) > 80:
                            # Appending the coordinates
                            apteki[key]['name_occurrence_coor'].append(cell.coordinate)

def extract_xlsx_search_range():
    ### Creating a tuple of search ranges for Revenue and Checks number
    for apteka in apteki_list:
        # The index in the apteki list
        the_index = apteki_list.index(apteka)
        # First range coordinate
        try:
            first_name_cooridnate = int(apteki[apteka]['name_occurrence_coor'][0][1:])  # Coorinate
        except IndexError:
            # The crutch for the last pharmacy
            the_index = the_index - 1
            first_name_cooridnate = int(apteki[apteki_list[the_index]]['name_occurrence_coor'][-1][1:]) + 1
        # Last range coordinate
        try:
            # The coordinate of the next pharmacy index minus one
            the_index = the_index + 1
            last_name_cooridnate = int(apteki[apteki_list[the_index]]['name_occurrence_coor'][0][1:]) - 1
        except IndexError:
            # The crutch for the last pharmacy. The discrepancy between the pharmacies coordinates minus one
            the_index = the_index - 1
            position_calc = int(apteki[apteki_list[the_index]]['name_occurrence_coor'][-1][1:]) - \
                            int(apteki[apteki_list[the_index - 1]]['name_occurrence_coor'][-1][1:]) - 1
            last_name_cooridnate = first_name_cooridnate + position_calc
        # The final search Range.
        # If revenue and cheques are found in those ranges, it would be assigned to a specific pharmacy
        apteki[apteka]['search_range'] = (first_name_cooridnate, last_name_cooridnate)

def extract_xlsx_coordinates_revenue_checks(active_cells):
    for row in active_cells:
        for cell in row:
            ##### Searching for Revenue str
            if isinstance(cell.value, str):
                if fuzz.token_sort_ratio('выручка', cell.value) > 50:
                    for apteka in apteki.keys():
                        if int(cell.coordinate[1:]) in range(apteki[apteka]['search_range'][0],
                                                             apteki[apteka]['search_range'][1]):
                            apteki[apteka]['revenue_coordinate'] = cell.coordinate
                ##### Searching for Check number str
                if fuzz.token_sort_ratio('Кол во чеков', cell.value) > 50:
                    for apteka in apteki.keys():
                        if int(cell.coordinate[1:]) in range(apteki[apteka]['search_range'][0],
                                                             apteki[apteka]['search_range'][1]):
                            apteki[apteka]['checks_coordinate'] = cell.coordinate

# The date container
the_date =[]
def extract_xslx_data(active_cells):
    for row in active_cells:
        for cell in row:
            #### Searching for Revenue(int or float) and Checks number(int or float)
            if isinstance(cell.value, int) or isinstance(cell.value, float):
                for apteka in apteki.keys():
                    if 'revenue_coordinate' in apteki[apteka]:
                        if cell.coordinate[1:] == apteki[apteka]['revenue_coordinate'][1:]:
                            apteki[apteka]['revenue'] = cell.value
                    if 'checks_coordinate' in apteki[apteka]:
                        if cell.coordinate[1:] == apteki[apteka]['checks_coordinate'][1:]:
                            apteki[apteka]['number_of_checks'] = cell.value

            ##### Searching for date str
            if isinstance(cell.value, str):
                if fuzz.token_sort_ratio('00.00.2000', cell.value) >= 45:
                    try:
                        if cell.value[2] and cell.value[5] == '.':
                            the_date.append(cell.value)
                    except IndexError:
                        pass

def open_google_sheet(scope='https://spreadsheets.google.com/feeds',
                      credentials='client_secret.json',
                      book_name='Farma2'):
    scope = [scope]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
    client = gspread.authorize(credentials)
    datesdict = dict(zip(months_numbers, months_words))
    name_of_the_sheet = datesdict[the_date[0][3:5]]
    book = client.open(book_name)
    sheet = book.worksheet(name_of_the_sheet)
    parsed_sheet = sheet.get_all_values()
    return sheet, parsed_sheet

def extract_googlesheet_coordinates(parsed_sheet):
    # Диапазон координат аптек
    for row_index, row in enumerate(parsed_sheet):
        for column_index, cell in enumerate(row):
            if isinstance(cell, str):
                for couner, apteka in enumerate(apteki.keys()):
                    # Костыль для маяка
                    if fuzz.token_sort_ratio('Красного Маяка,', cell) > 60:
                        if cell[-1] == '9':
                            apteki['Redmayak_9']['google_name_occurence_excel_format'] = (
                                column_index + 1, row_index + 1)
                            apteki['Redmayak_9']['google_name_occurence_list_format'] = (column_index, row_index)
                        if cell[-1] == '1':
                            apteki['Redmayak_4k1']['google_name_occurence_excel_format'] = (
                                column_index + 1, row_index + 1)
                            apteki['Redmayak_4k1']['google_name_occurence_list_format'] = (column_index, row_index)
                    # Для всех остальных
                    elif fuzz.token_sort_ratio(apteki[apteka]['rus_address'], cell) > 80:
                        apteki[apteka]['google_name_occurence_excel_format'] = (column_index + 1, row_index + 1)
                        apteki[apteka]['google_name_occurence_list_format'] = (column_index, row_index)
                    # Выручка
                    if fuzz.token_sort_ratio('Сумма выручки', cell) > 70:
                        try:
                            if column_index in range(apteki[apteka]['google_name_occurence_list_format'][0],
                                                     apteki[apteka]['google_name_occurence_list_format'][0] + 2):
                                apteki[apteka]['google_viruchka_column_index_excel_format'] = column_index + 1
                                apteki[apteka]['google_viruchka_column_index_list_format'] = column_index
                        except KeyError:
                            pass
                    # Чеки
                    if fuzz.token_sort_ratio('Кол-во чеков', cell) > 70:
                        try:
                            if column_index in range(apteki[apteka]['google_name_occurence_list_format'][0],
                                                     apteki[apteka]['google_name_occurence_list_format'][0] + 2):
                                apteki[apteka]['google_checks_column_index_excel_format'] = column_index + 1
                                apteki[apteka]['google_checks_column_index_list_format'] = column_index
                        except KeyError:
                            pass
                    # Номер строки по дате
                    if fuzz.token_sort_ratio(the_date[0], cell) > 90:
                        apteki[apteka]['google_date_row_index_excel_format'] = row_index + 1
                        apteki[apteka]['google_date_row_index_list_format'] = row_index

def update_google_sheet(googlesheet):
    for apteka in apteki.keys():
        try:
            # Выручка
            googlesheet.update_cell(apteki[apteka]['google_date_row_index_excel_format'],
                              apteki[apteka]['google_viruchka_column_index_excel_format'],
                              apteki[apteka]['revenue'])
        except KeyError:
            pass

            # Чеки
        try:
            googlesheet.update_cell(apteki[apteka]['google_date_row_index_excel_format'],
                              apteki[apteka]['google_checks_column_index_excel_format'],
                              apteki[apteka]['number_of_checks'])
        except KeyError:
            pass




ws = open_xlsx('2.xlsx')
active_cells = identify_xlsx_active_range(ws)
extract_xlsx_pharmacies_coordinates(active_cells)
extract_xlsx_search_range()
extract_xlsx_coordinates_revenue_checks(active_cells)
extract_xslx_data(active_cells)
googlesheet, parsedsheet = open_google_sheet()
extract_googlesheet_coordinates(parsedsheet)
update_google_sheet(googlesheet)
print('Updated')
print(apteki)

