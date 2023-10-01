import os
import numpy as np
import pandas as pd


##### CONSTANTS #####
directory = 'Scoresheets'

starting_round = 1
ending_round = 10
#####################


# Header format for the tossup_stats and bonus_stats arrays (which are also the final spreadsheets)
tossup_header = ['Packet', '#', 15, 10, -5, 'DT', 'Average']
bonus_header  = ['Packet', '#', 30, 20, 10,    0, 'Average']

num_rounds = ending_round - starting_round + 1
sheets_to_use = [i + starting_round + 1 for i in range(num_rounds)]

tossup_stats = [[0 for j in range(len(tossup_header))] for i in range(20*num_rounds + 1)]
bonus_stats  = [[0 for j in range(len(bonus_header))] for i in range(20*num_rounds + 1)]

for i in range(20*num_rounds + 1):
    tossup_stats[i][1] = (i - 1)%20 + 1
    bonus_stats[i][1]  = (i - 1)%20 + 1


# Read all of the spreadsheets in `directory`
for filename in os.listdir(directory):
    if filename == '.DS_Store': continue

    filepath = directory + '/' + filename
    print('starting', filepath)

    sheet_counter = -1
    for game in pd.read_excel(filepath, sheet_name=sheets_to_use).values():
        sheet_counter += 1

        # converts the pandas dataframe to a np.array
        game = np.append(np.array([game.columns]), game.to_numpy(), axis=0)

        for i in range(20):  # compiling tossup stats
            row = i + 20*sheet_counter + 1
            tossup_stats[row][0] = sheet_counter + starting_round
            went_dead = True
            for j in [2, 3, 4, 5, 6, 7, 8, 12, 13, 14, 15, 16, 17, 18]:
                cell = str(game[i+3][j]).strip()
                if cell == '15' or cell == '15.0':
                    tossup_stats[row][2] += 1
                    went_dead = False
                if cell == '10' or cell == '10.0':
                    tossup_stats[row][3] += 1
                    went_dead = False
                if cell == '-5' or cell == '-5.0':
                    tossup_stats[row][4] += 1
                if cell == 'DT':
                    actually_went_dead = True

            if went_dead:
                tossup_stats[row][5] += 1

            if went_dead and not actually_went_dead:
                print(f'warning - tossup {i+1} is dead but not marked with DT')

        for i in range(20):  # compiling bonus stats
            row = i + 20*sheet_counter + 1
            bonus_stats[row][0] = sheet_counter + starting_round
            for j in [8, 18]:
                cell = str(game[i+3][j]).strip()
                if cell == '30' or cell == '30.0': bonus_stats[row][2] += 1
                if cell == '20' or cell == '20.0': bonus_stats[row][3] += 1
                if cell == '10' or cell == '10.0': bonus_stats[row][4] += 1
                if cell ==  '0' or cell == '0.0' : bonus_stats[row][5] += 1


for array in tossup_stats:
    if array[0] == 'Packet': continue
    if sum(array[2:6]) == 0:
        array[6] = 0
    else:
        array[6] = (15*array[2] + 10*array[3] - 5*array[4]) / sum(array[2:6])
        array[6] = round(array[6], 2)

for array in bonus_stats:
    if array[0] == 'Category': continue
    if sum(array[2:6]) == 0:
        array[6] = 0
    else:
        array[6] = (30*array[2] + 20*array[3] + 10*array[4]) / sum(array[2:6])
        array[6] = round(array[6], 2)

tossup_stats[0] = tossup_header
bonus_stats[0] = bonus_header

# Writes the data to excel spreadsheets
with pd.ExcelWriter(directory + '_stats.xlsx') as writer:
    stat_sheet = pd.DataFrame(np.array(tossup_stats))
    stat_sheet.to_excel(writer, sheet_name='Tossups', header=None, index=False)

    stat_sheet = pd.DataFrame(np.array(bonus_stats))
    stat_sheet.to_excel(writer, sheet_name='Bonuses', header=None, index=False)
