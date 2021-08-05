import os
import numpy as np
import pandas as pd


##### CONSTANTS #####
directory = 'Scoresheets'

starting_round = 1
ending_round = 5
#####################


num_rounds = ending_round - starting_round + 1
sheets_to_use = [i + starting_round + 1 for i in range(num_rounds)]

tossup_stats = [[0 for j in range(9)] for i in range(20*num_rounds + 1)]
bonus_stats  = [[0 for j in range(9)] for i in range(20*num_rounds + 1)]

for i in range(20*num_rounds + 1):
    tossup_stats[i][3] = (i - 1)%20 + 1
    bonus_stats[i][3]  = (i - 1)%20 + 1

# Header format for the tossup_stats and bonus_stats arrays (which are also the final spreadsheets)
tossup_stats[0] = ['Category', 'Subcategory', 'Packet', '#', 15, 10, -5, 'DT', 'Average']
bonus_stats[0]  = ['Category', 'Subcategory', 'Packet', '#', 30, 20, 10,    0, 'Average']


# Read all of the spreadsheets in `directory`
for filename in os.listdir(directory):
    filepath = directory + '/' + filename
    print('starting', filepath)

    sheet_counter = -1
    for game in pd.read_excel(filepath, sheet_name=sheets_to_use).values():
        sheet_counter += 1

        # converts the pandas dataframe to a np.array
        game = np.append(np.array([game.columns]), game.to_numpy(), axis=0)

        for i in range(20):  # compiling tossup stats
            row = i + 20*sheet_counter + 1
            tossup_stats[row][2] = sheet_counter + starting_round
            went_dead = True
            for j in [2, 3, 4, 5, 6, 7, 12, 13, 14, 15, 16, 17]:
                cell = str(game[i+3][j]).strip()
                if cell == '15':
                    tossup_stats[row][4] += 1
                    went_dead = False
                if cell == '10':
                    tossup_stats[row][5] += 1
                    went_dead = False
                if cell == '-5':
                    tossup_stats[row][6] += 1

            if went_dead:
                tossup_stats[row][7] += 1

        for i in range(20):  # compiling bonus stats
            row = i + 20*sheet_counter + 1
            bonus_stats[row][2] = sheet_counter + starting_round
            for j in [8, 18]:
                cell = str(game[i+3][j]).strip()
                if cell == '30': bonus_stats[row][4] += 1
                if cell == '20': bonus_stats[row][5] += 1
                if cell == '10': bonus_stats[row][6] += 1
                if cell ==  '0': bonus_stats[row][7] += 1


for array in tossup_stats:
    if array[0] == 'Category': continue
    array[8] = (15*array[4] + 10*array[5] - 5*array[6]) / sum(array[4:8])
    array[8] = round(array[8], 2)

for array in bonus_stats:
    if array[0] == 'Category': continue
    array[8] = (30*array[4] + 20*array[5] + 10*array[6]) / sum(array[4:8])
    array[8] = round(array[8], 2)


# Writes the data to excel spreadsheets
with pd.ExcelWriter(directory + '_stats.xlsx') as writer:
    stat_sheet = pd.DataFrame(np.array(tossup_stats))
    stat_sheet.to_excel(writer, sheet_name='Tossups', header=None, index=False)

    stat_sheet = pd.DataFrame(np.array(bonus_stats))
    stat_sheet.to_excel(writer, sheet_name='Bonuses', header=None, index=False)
