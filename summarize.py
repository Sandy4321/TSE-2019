import os
import csv
import xlrd

barriers = ['Finding a task to start with', # NO
            'Finding a mentor', # NO
            'Poor "How to contribute"', #NO
            'Newcomers don’t know what is the contribution flow', #NO
            'Lack of Patience', # NC
            'Shyness', # NC
            'Lack of domain expertise', # NC
            'Lack of knowledge in project processes and practice', # NC
            'Knowledge on technologies and tools used', # NC
            'Knowledge of versioning control systems', # NC
            'Receiving answers with too advanced/complex contents', # RI
            'Impolite answers', # RI
            'Not receiving an answer', # RI
            'Delayed Answer', # RI
            'Some newcomers need to contact a real person', # DC
            'Documentation Outdated', # DP
            'Documentation Overload', # DP
            'Documentation Unclear', # DP
            'Documentation Spread', # DP
            'Documentation General', # DP
            'Building workspace locally', # TH
            'Lack of information on how to send a contribution', # TH
            'Getting contribution accepted', # TH
            'Issue to create a patch'] # TH

categories = {
                'NO': ['Finding a task to start with', # NO
                    'Finding a mentor', # NO
                    'Poor "How to contribute"', #NO
                    'Newcomers don’t know what is the contribution flow'], #NO
                'NC' : ['Lack of Patience', # NC
                    'Shyness', # NC
                    'Lack of domain expertise', # NC
                    'Lack of knowledge in project processes and practice', # NC
                    'Knowledge on technologies and tools used', # NC
                    'Knowledge of versioning control systems'], # NC
                'RI' : ['Receiving answers with too advanced/complex contents', # RI
                    'Impolite answers', # RI
                    'Not receiving an answer', # RI
                    'Delayed Answer'], # RI
                'DC' : ['Some newcomers need to contact a real person'], # DC
                'DP' : ['Documentation Outdated', # DP
                    'Documentation Overload', # DP
                    'Documentation Unclear', # DP
                    'Documentation Spread', # DP
                    'Documentation General'], # DP
                'TH' : ['Building workspace locally', # TH
                    'Lack of information on how to send a contribution', # TH
                    'Getting contribution accepted', # TH
                    'Issue to create a patch'] # TH
            }

facets = ['Females motivation',
          'Females lower computer self-efficacy',
          'Females risk-averse than males',
          'Females process information comprehensively',
          'Females learn in process-oriented learning styles']

def get_diary_number(filename):
    if str(filename).endswith('(Igor).xlsx'):
        filename = filename.replace('(Igor).xlsx','')
    else:
        filename = filename.replace('.xlsx', '')
    filename = filename.replace('Diary ', '')
    return int(filename)

def parse_diaries(folders, labels):
    diaries = {}

    for genre in folders.keys():
        diaries[genre] = {}
        diaries[genre]['barriers'] = {}
        diaries[genre]['facets'] = {}

        if not os.path.isdir(folders[genre]):
            raise Exception(folders[genre] + ' does not exist.')

        for filename in os.listdir(folders[genre]):
            workbook = xlrd.open_workbook(folders[genre] + filename)
            spreadsheet = workbook.sheet_by_name('Sheet')
            diary_number = get_diary_number(filename)

            if str(filename).endswith('(Igor).xlsx'):
                diaries[genre]['barriers'][diary_number] = {}

                for row in range(1, spreadsheet.nrows):
                    row_value = '(Barriers ' + str(diary_number) + ', row ' + str(row) + ') ' + spreadsheet.cell_value(row, 0)

                    diaries[genre]['barriers'][diary_number][row] = {'diary': diary_number,
                                                                     'row_value': row_value,
                                                                     'occurred': []}

                    for column, cell in enumerate(spreadsheet.row_slice(row, 1, 25)):
                        if cell.ctype != 0: # If cell is not empty
                            if cell.value in labels:
                                diaries[genre]['barriers'][diary_number][row]['occurred'].append(barriers[column])
            else:
                diaries[genre]['facets'][diary_number] = {}

                for row in range(1, spreadsheet.nrows):
                    row_value = '(Facets ' + str(diary_number) + ', row ' + str(row) + ') ' + spreadsheet.cell_value(row, 0)

                    diaries[genre]['facets'][diary_number][row] = {'diary': diary_number,
                                                                   'row_value': row_value,
                                                                   'occurred': []}

                    for column, cell in enumerate(spreadsheet.row_slice(row, 1, 6)):
                        if cell.ctype != 0: # If cell is not empty
                            if cell.value in labels:
                                diaries[genre]['facets'][diary_number][row]['occurred'].append(facets[column])
    return diaries

def adjacency_matrix(diaries, filename, data_type):
    matrix = {}
    vertices = barriers + facets + ['NO FACET', 'NO BARRIER']

    for adjacent_i in vertices:
        matrix[adjacent_i] = {}
        for adjacent_j in vertices:
            matrix[adjacent_i][adjacent_j] = []

    for diary in diaries['barriers']:
        for row in diaries['barriers'][diary]:
            occurrences = diaries['barriers'][diary][row]['occurred'] + diaries['facets'][diary][row]['occurred']

            for adjacent_i in occurrences:
                if adjacent_i in barriers and not any(facet in occurrences for facet in facets):
                    if data_type == 'frequency':
                        matrix[adjacent_i]['NO FACET'].append(1)
                    elif data_type == 'diaries_frequency':
                        matrix[adjacent_i]['NO FACET'].append(diary)
                    elif data_type == 'rows':
                        matrix[adjacent_i]['NO FACET'].append(diaries['barriers'][diary][row]['row_value'])
                    else:
                        raise Exception('You are using an invalid data type!')

                if adjacent_i in facets and not any(barrier in occurrences for barrier in barriers):
                    if data_type == 'frequency':
                        matrix[adjacent_i]['NO BARRIER'].append(1)
                    elif data_type == 'diaries_frequency':
                        matrix[adjacent_i]['NO BARRIER'].append(diary)
                    elif data_type == 'rows':
                        matrix[adjacent_i]['NO BARRIER'].append(diaries['barriers'][diary][row]['row_value'])
                    else:
                        raise Exception('You are using an invalid data type!')

                for adjacent_j in occurrences:
                        if adjacent_i != adjacent_j:
                            if data_type == 'frequency':
                                matrix[adjacent_i][adjacent_j].append(1)
                            elif data_type == 'diaries_frequency':
                                matrix[adjacent_i][adjacent_j].append(diary)
                            elif data_type == 'rows':
                                matrix[adjacent_i][adjacent_j].append(diaries['barriers'][diary][row]['row_value'])
                            else:
                                raise Exception('You are using an invalid data type!')

        with open(filename, 'w', encoding='utf-8') as export_file:
            writer = csv.DictWriter(export_file, fieldnames= ['#'] +  vertices)
            writer.writeheader()

            for adjacent_i in vertices:
                row = {'#': adjacent_i}

                for adjacent_j in vertices:
                    if data_type == 'frequency':
                        if matrix[adjacent_i][adjacent_j]:
                            row.update({adjacent_j: sum(matrix[adjacent_i][adjacent_j])})
                        else:
                            row.update({adjacent_j: 0})
                    elif data_type == 'diaries_frequency':
                        if matrix[adjacent_i][adjacent_j]:
                            row.update({adjacent_j: ', '.join(str(diary) for diary in matrix[adjacent_i][adjacent_j])})
                        else:
                            row.update({adjacent_j: ''})
                    elif data_type == 'rows':
                        if matrix[adjacent_i][adjacent_j]:
                            row.update({adjacent_j: ', '.join(matrix[adjacent_i][adjacent_j])})
                        else:
                            row.update({adjacent_j: ''})
                    else:
                        raise Exception('You are using an invalid data type! ' + data_type + ' is not accepted.')

                writer.writerow(row)


def diaries_per_category(diaries, filename):
    matrix = {}
    categories_list = ['NO', 'NC', 'RI', 'DC', 'DP', 'TH']
    vertices = categories_list + facets + ['NO FACET', 'NO BARRIER']

    for adjacent_i in vertices:
        matrix[adjacent_i] = {}
        for adjacent_j in vertices:
            matrix[adjacent_i][adjacent_j] = []

    for diary in diaries['barriers']:
        for row in diaries['barriers'][diary]:
            occurrences = diaries['barriers'][diary][row]['occurred'] + diaries['facets'][diary][row]['occurred']

            for adjacent_i in occurrences:
                adjacent_i_category = None
                adjacent_j_category = None

                for category in categories:
                    if adjacent_i in categories[category]:
                        adjacent_i_category = category

                if adjacent_i_category != None and not any(facet in occurrences for facet in facets):
                    if not diary in matrix[adjacent_i_category]['NO FACET']:
                        matrix[adjacent_i_category]['NO FACET'].append(diary)

                if adjacent_i in facets and not any(barrier in occurrences for barrier in barriers):
                    if not diary in matrix[adjacent_i]['NO BARRIER']:
                        matrix[adjacent_i]['NO BARRIER'].append(diary)

                for adjacent_j in occurrences:
                    if adjacent_i != adjacent_j:
                        for category in categories:
                            if adjacent_j in categories[category]:
                                adjacent_j_category = category
                        
                        if adjacent_i_category != None:
                            if adjacent_j_category != None:
                                if not diary in matrix[adjacent_i_category][adjacent_j_category]:
                                    matrix[adjacent_i_category][adjacent_j_category].append(diary)
                            else:
                                if not diary in matrix[adjacent_i_category][adjacent_j]:
                                    matrix[adjacent_i_category][adjacent_j].append(diary)
                        else:
                            if adjacent_j_category != None:
                                if not diary in matrix[adjacent_i][adjacent_j_category]:
                                    matrix[adjacent_i][adjacent_j_category].append(diary)
                            else:
                                if not diary in matrix[adjacent_i][adjacent_j]:
                                    matrix[adjacent_i][adjacent_j].append(diary)

                    adjacent_j_category = None


    with open(filename, 'w', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames= ['#'] +  vertices)
        writer.writeheader()

        for adjacent_i in vertices:
            row = {'#': adjacent_i}

            for adjacent_j in vertices:
                if matrix[adjacent_i][adjacent_j]:
                    row.update({adjacent_j: len(matrix[adjacent_i][adjacent_j])})
                else:
                    row.update({adjacent_j: 0})

            writer.writerow(row)


if __name__ == "__main__":
    men_diaries_folder = "men_diaries/"
    women_diaries_folder = "women_diaries/"

    output_folder = 'results/'    
    frequency_folder = output_folder + "frequency/"
    rows_folder = output_folder +  'rows/'
    diaries_frequency_folder = output_folder + 'diaries_frequencies/'
    categories_frequency_folder = output_folder + 'categories_frequency/'

    if not os.path.isdir(output_folder):
        os.mkdir(output_folder)
    if not os.path.isdir(frequency_folder):
        os.mkdir(frequency_folder)
    if not os.path.isdir(rows_folder):
        os.mkdir(rows_folder)
    if not os.path.isdir(diaries_frequency_folder):
        os.mkdir(diaries_frequency_folder)
    if not os.path.isdir(categories_frequency_folder):
        os.mkdir(categories_frequency_folder)

    labels = {'a': ['A','X','x'],
              'a_minus': ['A-','X','x'],
              'a_plus': ['A+','X','x'],
              'all': ['A+','A-','A','X','x']}

    for key in labels.keys():
        diaries = parse_diaries(folders={'men': men_diaries_folder, 'women': women_diaries_folder}, labels=labels[key])
        # Occurrences
        # adjacency_matrix(diaries=diaries['men'], filename=frequency_folder + 'men_' + key + '_frequency.csv', data_type='frequency')
        # adjacency_matrix(diaries=diaries['women'], filename=frequency_folder + 'women_' + key + '_frequency.csv', data_type='frequency')
        # Quotes
        # adjacency_matrix(diaries=diaries['men'], filename=rows_folder + 'men_' + key + '_rows.csv', data_type='rows')
        # adjacency_matrix(diaries=diaries['women'], filename=rows_folder + 'women_' + key + '_rows.csv', data_type='rows')
        # Diary Number
        # adjacency_matrix(diaries=diaries['men'], filename=diaries_frequency_folder + 'men_' + key + '_diaries_frequency.csv', data_type='diaries_frequency')
        # adjacency_matrix(diaries=diaries['women'], filename=diaries_frequency_folder + 'women_' + key + '_diaries_frequency.csv', data_type='diaries_frequency')
        # Diaries per category
        diaries_per_category(diaries['men'], filename=categories_frequency_folder + 'men_' + key + '_categories_frequency.csv')
        diaries_per_category(diaries['women'], filename=categories_frequency_folder + 'women_' + key + '_categories_frequency.csv')
