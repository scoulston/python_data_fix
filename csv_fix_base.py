#!/usr/bin/env python
# -*- coding: utf-8 -*-


import pandas as pd
import csv
import random
import re
from time import process_time

def order(frame,var):
    if type(var) is str:
        var = [var] #let the command take a string or list
    varlist =[w for w in frame.columns if w not in var]
    frame = frame[var+varlist]
    return frame

default_base_file_path = 'C:/Users/user/fixed_data' #Just the base file path to send the fixed file to.
default_junk_file_path = 'C:/Users/user/junk' #Just the base junk file path to send the junk/intermediate files to.

def twent_cha_fix(str1): #takes the two or three letters with the dash off at the beginning and only saves the first 20 characters after that
    str2 = ""
    if re.search("^DB-", str1):
        str2 = str1[3:]
    elif re.search("^DMS-", str1):
        str2 = str1[4:]
    if len(str2) > 20:
        str2 = str2[:21]
    return str2

#All "start" and "end" variables are to calculate the time it takes to do certain sections of code
#Takes a full path to a csv_file to fix, a base file path where you want your fixed csv to go, and a junk_file_path where you want the junk files to go.
def fix_csv(csv_file_path, fixed_file_name = 'fixed_data.csv', base_file_path=default_base_file_path, junk_file_path= default_junk_file_path):
    csv.register_dialect('star', delimiter='*', quoting=csv.QUOTE_NONE)

    start = process_time()
    data = pd.read_csv(csv_file_path)

    monster_descriptions = data['Item Description']

    monster_descriptions.to_csv(junk_file_path + 'monster_descriptions.csv', index = False)

    end = process_time()
    csv_open_export_descrip_time = "Opening csv and exporting description: " + str(end - start)

    start = process_time()
    with open(junk_file_path + 'monster_descriptions.csv', newline='') as monster_descriptions_read_csv, \
            open(junk_file_path + 'monster_descriptions_fixed.csv', 'w', newline='') as monster_descriptions_write_csv:

        missing = [6, 7]  # 1-indexed positions of missing values
        missing.sort()  # enforce the increasing order
        reader = csv.reader(monster_descriptions_read_csv, delimiter='*', skipinitialspace=True)
        writer = csv.writer(monster_descriptions_write_csv, delimiter = ',' )
        next(reader)
        header = 'metal_type','idk1','idk2','OD_or_H','od_name','ID_or_W','id_name','Length','length_name','cut','company','idk3','idk4', 'idk5' # get first row (header)
        writer.writerow(header)  # write it back
        for row in reader:
            if len(row) < 12:
                 #row shorter than header -> insert empty strings
                #inserting changes indices so `missing` must be sorted
                for idx in missing:
                    row.insert(idx - 1, '')
            writer.writerow(row)
    end = process_time()
    fix_monster_descrip_time = "Time to fix monster description: " + str(end - start)

    start = process_time()
    data=data.rename(columns = {'Item Description':'metal_type*idk1*idk2*OD_or_H*od_name*ID_or_W*id_name*Length*length_name*cut*company*idk3*idk4*idk5', 'Customer': 'Company'})
    #converts Day of Date column to a datetime column (to easily extract month, day and year from the column)
    data['Day of Date'] = pd.to_datetime(data['Day of Date'])
    # Create new columns for month, day and year
    data['Month'] = data['Day of Date'].dt.month
    data['Day'] = data['Day of Date'].dt.day
    data['YR'] = data['Day of Date'].dt.year
    end = process_time()
    fix_date_time = "Time to fix date: " + str(end - start)





    monster_data = pd.read_csv(junk_file_path + 'monster_descriptions_fixed.csv')   
    start = process_time()

    #applies the fix to company where it takes out DMS- or DB- and only keeps 20characters after that, and strips extra space if any
    data['Company'] = data['Company'].apply(twent_cha_fix)
    data['Company'] = data['Company'].str.strip()


    #old, chunky, slow code
    #for index, row in data.iterrows():
        #data.loc[index, 'Company'] = twent_cha_fix(data.loc[index, 'Company'])
        #data.loc[index,'Shape'] = monster_data.loc[index, 'metal_type']
        #data.loc[index, 'Cond'] = monster_data.loc[index, 'idk1']
        #data.loc[index, 'Grade'] = monster_data.loc[index, 'idk2']
        #data.loc[index, 'OD_or_H'] = monster_data.loc[index, 'OD_or_H']
        #data.loc[index, 'ID_or_W'] = monster_data.loc[index, 'ID_or_W']
        #data.loc[index, 'Len'] = monster_data.loc[index, 'Length']
        #data.loc[index, 'cut'] = monster_data.loc[index, 'cut']
        #data.loc[index, 'company'] = monster_data.loc[index, 'company']
        #data.loc[index, 'idk3'] = monster_data.loc[index, 'idk3']

    #merges the description and rest of the data back together
    data = pd.concat([data, monster_data], axis = 1, sort=False)

    #gets rid of unwanted columns
    del data['Day of Date']
    del data['metal_type*idk1*idk2*OD_or_H*od_name*ID_or_W*id_name*Length*length_name*cut*company*idk3*idk4*idk5']
    del data['od_name']
    del data['id_name']
    del data['length_name']
    del data['company']
    del data['idk4']
    del data['idk5']

    #renames columns to better names
    data.rename(columns = {'metal_type' : 'Shape', 'idk1' : 'Cond', 'idk2': 'Grade', 'Length': 'Len'}, inplace=True)
    data.to_csv(junk_file_path + 'intermed_fix.csv', index=False) #This is the point fixing everything else but before fixing the triplicates
    end = process_time()
    add_columns_from_monster_data_time = "Time to add columns from monster data: " + str(end - start)

    start = process_time()
    triple_data = pd.read_csv(junk_file_path + 'intermed_fix.csv')
    measure_dict = {
        'Pounds Shipped': [],
        'Shipment Dollars': [],
        'QTY': []
    }
    #Converts appropriate numbers into a dictionary for the appropriate measure from the Measure Names Column
    for index, row in triple_data.iterrows():
        for key, value in measure_dict.items():
            if triple_data.loc[index, 'Measure Names'] == key:
                value.append(triple_data.loc[index, 'Measure Values'])
    measure_df = pd.DataFrame(measure_dict)

    fixed_triple_df = triple_data[triple_data.index % 3 == 0].reset_index(drop=True)  # Selects every 3rd raw starting from 0
    fixed_triple_df.reset_index(drop=True)
    final_df = pd.concat([fixed_triple_df, measure_df], axis=1)
    final_df.drop(columns=['Measure Names', 'Measure Values'], inplace = True)
    fixed_triple_df.to_csv(junk_file_path + 'test.csv')
    measure_df.to_csv(junk_file_path + 'measuretest.csv')
    #renames the Shipment Dollars to Sales
    final_df=final_df.rename(columns = {'Shipment Dollars': 'Sales'})

    #rounds the sales and pounds 
    final_df['Sales'] = final_df['Sales'].apply(round)
    final_df['Pounds Shipped'] = final_df['Pounds Shipped'].apply(round)

    #orders the columns
    final_df = order(final_df, ['Company', 'Month', 'Day', 'YR', 'Grade', 'QTY', 'OD_or_H', 'ID_or_W', 'Len', 'Cond', 'Shape', 'Pounds Shipped', 'Sales', 'Salesperson'])
    final_df.to_csv(base_file_path + fixed_file_name, index= False)
    end = process_time()
    fix_triplicate_time = "Time to fix triplicate rows: " + str(end - start)

    #printing calc for time to do each section (can be commented out until needed)
    #print(csv_open_export_descrip_time)
    #print(fix_monster_descrip_time)
    #print(fix_date_time)
    #print(add_columns_from_monster_data_time)
    #print(fix_triplicate_time)

def input_csv_fix():
    path = input("Enter the path of the csv file to be fixed: ")
    #path = path[1:-1] #use this if the copied path has quotes around it already (copying a path from a windows 10 computer for instance)
    fixed_name = input("Enter what you want as the name of the fixed file: ")
    fix_csv(r"{path}".format(path = path), '{fixed_name}.csv'.format(fixed_name = fixed_name))
    return fixed_name #returns the fixed name so it can be used for other functions such as the summary data fix