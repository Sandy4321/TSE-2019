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

def read_diaries(folders, labels):
    diaries = {}

    for genre in folders.keys():
        diaries[genre] = {}
        diaries[genre]['barriers'] = {}
        diaries[genre]['facets'] = {}

        if not os.path.isdir(folders[genre]):
            raise Exception(folders[genre] + ' does not exist.')

        for filename in os.listdir(folders[genre]):
            if str(filename).endswith('(Igor).xlsx'):
                workbook = xlrd.open_workbook(folders[genre] + filename)
                spreadsheet = workbook.sheet_by_index(0)
                diary_number = get_diary_number(filename)
                diaries[genre]['barriers'][diary_number] = {}

                for row in range(1, spreadsheet.nrows):
                    diaries[genre]['barriers'][diary_number][row] = {'diary': diary_number, 
                                                                     'row_value': '(Barriers ' + str(diary_number) + ', row ' + str(row) + ') ' + spreadsheet.cell_value(row, 0), 
                                                                     'occurred': []}

                    for column, cell in enumerate(spreadsheet.row_slice(row, 1, 25)):
                        if cell.ctype != 0:
                            if cell.value in labels:
                                diaries[genre]['barriers'][diary_number][row]['occurred'].append(barriers[column])
            else:
                workbook = xlrd.open_workbook(folders[genre] + filename)
                spreadsheet = workbook.sheet_by_index(0)
                diary_number = get_diary_number(filename)
                diaries[genre]['facets'][diary_number] = {}

                for row in range(1, spreadsheet.nrows):
                    diaries[genre]['facets'][diary_number][row] = {'diary': diary_number,
                                                                   'row_value': '(Facets ' + str(diary_number) + ', row ' + str(row) + ') ' + spreadsheet.cell_value(row, 0), 
                                                                   'occurred': []}

                    for column, cell in enumerate(spreadsheet.row_slice(row, 1, 5)):
                        if cell.ctype != 0:
                            if cell.value in labels:
                                diaries[genre]['facets'][diary_number][row]['occurred'].append(facets[column])
    return diaries

def adjacency_matrix(diaries, filename, output):
    matrix = {}
    labels = barriers + facets + ['NO FACET', 'NO BARRIER']

    for adjacent_i in labels:
        matrix[adjacent_i] = {}

        for adjacent_j in labels:
            matrix[adjacent_i][adjacent_j] = []

    for diary in diaries['barriers']:
        for row in diaries['barriers'][diary]:
            dependencies = diaries['barriers'][diary][row]['occurred'] + diaries['facets'][diary][row]['occurred']

            for adjacent_i in dependencies:
                if adjacent_i in barriers and not any(facet in dependencies for facet in facets):
                    if output == 'occurrences':
                        matrix[adjacent_i]['NO FACET'].append(1)
                    elif output == 'diary' and not diary in matrix[adjacent_i]['NO FACET']:
                        matrix[adjacent_i]['NO FACET'].append(diary)
                    else:
                        matrix[adjacent_i]['NO FACET'].append(diaries['barriers'][diary][row][output])

                if adjacent_i in facets and not any(barrier in dependencies for barrier in barriers):
                    if output == 'occurrences':
                        matrix[adjacent_i]['NO BARRIER'].append(1)
                    elif output == 'diary' and not diary in matrix[adjacent_i]['NO BARRIER']:
                        matrix[adjacent_i]['NO BARRIER'].append(diary)
                    else:
                        matrix[adjacent_i]['NO BARRIER'].append(diaries['barriers'][diary][row][output])

                for adjacent_j in dependencies:
                        if adjacent_i != adjacent_j:
                            if output == 'occurrences':
                                matrix[adjacent_i][adjacent_j].append(1)
                            else:
                                matrix[adjacent_i][adjacent_j].append(diaries['barriers'][diary][row][output])

    with open(filename, 'w', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames= ['adjacents'] +  labels)
        writer.writeheader()

        for adjacent_i in labels:
            row = {'adjacents': adjacent_i}

            for adjacent_j in labels:
                if matrix[adjacent_i][adjacent_j]:
                    if output == 'occurrences':
                       row.update({adjacent_j: sum(matrix[adjacent_i][adjacent_j])})
                    else:
                       row.update({adjacent_j: matrix[adjacent_i][adjacent_j]})
                else:
                   row.update({adjacent_j: ''})

            writer.writerow(row)

if __name__ == "__main__":
    male_folder = "./MaleDiaries/"
    female_folder = "./FemaleDiaries/"
    output_folder = "./Matrix/"

    if not os.path.isdir(output_folder):
        os.mkdir(output_folder)

    labels = {
        'a': ['A','X','x'],
        'a_minus': ['A-','X','x'],
        'a_plus': ['A+','X','x'],
        'all': ['A+','A-','A','X','x']
    }

    for key in labels.keys():
        diaries = read_diaries(folders={'male': male_folder, 'female': female_folder}, labels=labels[key])
        # Occurrences
        adjacency_matrix(diaries=diaries['male'], filename=output_folder + 'male_' + key + '.csv', output='occurrences')
        adjacency_matrix(diaries=diaries['female'], filename=output_folder + 'female_' + key + '.csv', output='occurrences')
        # Quotes
        adjacency_matrix(diaries=diaries['male'], filename=output_folder + 'male_' + key + '_quotes.csv', output='row_value')
        adjacency_matrix(diaries=diaries['female'], filename=output_folder + 'female_' + key + '_quotes.csv', output='row_value')
        # Diary Number
        adjacency_matrix(diaries=diaries['male'], filename=output_folder + 'male_' + key + '_diaries.csv', output='diary')
        adjacency_matrix(diaries=diaries['female'], filename=output_folder + 'female_' + key + '_diaries.csv', output='diary')

'''
Insert the item description into the hyperlink range.
Set the cursor to the mouse pointer.

When clicked, the description of an item was not redirecting
 the user to the related page. The description was moved to the
hyperlink range, and now all the item elements are clickable.
To assure that the item will look like a clickable element,
the cursor was set to the mouse pointer, so when the user point
over the item, it looks like a clickable element.

Add the cursor:pointer rule to the gallery div of item_fig_desc.html
Signed-off-by: Felipe Fronchetti's avatarFelipe Fronchetti <fronchetti@usp.br>
'''