import openpyxl as op
import pandas as pd
import re
from datetime import timedelta, date

def get_status(remarksstr):
    curtail_status = remarksstr.split(' ')[0]
    if curtail_status == 'curtailed' or curtail_status == 'Curtailed':
        return 1
    elif curtail_status == 'released' or curtail_status == 'Released' or curtail_status == 'releaseed':
        return 0
    else:
        return -1
    
def contain_ghani(row):
    sub_station = 'Ghani'
    station = row['Remarks']
    if sub_station.lower() in station.lower():
        return row['Solar cutailment MW']
    else:
        return 0

def contain_jam(row):
    sub_station = 'Jammalamadugu'
    station = row['Remarks']
    if sub_station.lower() in station.lower():
        return row['Solar cutailment MW']
    else:
        return 0

def contain_tala(row):
    sub_station = 'Talarichruvu'
    sub_station1 = 'Talaricheruvu'
    station = row['Remarks']
    if sub_station.lower() in station.lower() or sub_station1.lower() in station.lower():
        return row['Solar cutailment MW']
    else:
        return 0

def contail_myal(row):
    sub_station = 'Mylavaram'
    station = row['Remarks']
    if sub_station.lower() in station.lower():
        return row['Solar cutailment MW']
    else:
        return 0

def contain_all(row):
    sub_station = 'Distributed stations'
    station = row['Remarks']
    if sub_station.lower() in station.lower():
        return row['Solar cutailment MW']
    else:
        return 0

def data_entry_problems(row):
    ret = ["" for _ in row.index]
    if row['Status'] and row['Solar cutailment MW'] > 0:
        return ret
    elif row['Status'] and row['Solar cutailment MW'] < 0:
        ret[row.index.get_loc('Solar cutailment MW')] = "background-color: red"
        return ret
    elif not row['Status'] and row['Solar cutailment MW'] > 0:
        ret[row.index.get_loc('Solar cutailment MW')] = "background-color: red"
        return ret
    elif not row['Status'] and row['Solar cutailment MW'] < 0:
        return ret

def color_negative_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for negative
    strings, black otherwise.
    """
    color = 'red' if val < 0 else 'black'
    return 'color: %s' % color


start_dt = date(2019, 9, 24)
end_dt = date(2019, 12, 31)
date_list = [(start_dt + timedelta(n)).strftime('%d-%m-%Y') for n in range(int ((end_dt - start_dt).days)+1)]
df = pd.read_excel('apsolarcurtailment.xlsx')
df.rename(columns = {'Reasons for Backing down': 'Remarks'},inplace = True)
df['Status']  = df['Remarks'].apply(get_status)
df['Ghani'] = df.apply(contain_ghani,axis=1)
df['Jammalamadugu'] = df.apply(contain_jam,axis=1)
df['Talarichruvu'] = df.apply(contain_tala,axis=1)
df['Mylavaram'] = df.apply(contail_myal,axis=1)
df['All Stations'] = df.apply(contain_all,axis=1)
##df['Date'] =  pd.to_datetime(df['Date'], infer_datetime_format=True)
df.style.apply(data_entry_problems,axis=1)

##Seperate for All Stations
with pd.ExcelWriter('output.xlsx',engine='openpyxl') as writer:
    df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Data')
    sub_df = df[df['Ghani'] != 0]
    sub_df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Ghani')
    sub_df = df[df['Jammalamadugu'] != 0]
    sub_df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Jammalamadugu')
    sub_df = df[df['Talarichruvu'] != 0]
    sub_df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Talarichruvu')
    sub_df = df[df['Mylavaram'] != 0]
    sub_df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Mylavaram')
    sub_df = df[df['All Stations'] != 0]
    sub_df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='All Stations')

final_df = pd.DataFrame(columns=['Date','Block No','Cumulative Curtailment MW','Solar cutailment MW','Remarks'])
with pd.ExcelWriter('BlockWiseData.xlsx',engine='openpyxl') as writer:
    for given_date in date_list:
        date_df = df[df['Date'] == given_date]
        curtailment_list = [0 for x in range(1,97)]
        for index, row in date_df.iterrows():
            curtailment_list[row['Block No'] - 1] = curtailment_list[row['Block No'] - 1] + row['Solar cutailment MW']
        df_curtailment_list = [0 for x in range(1,97)]
        temp_no = 0
        for x in range(len(curtailment_list)):
            df_curtailment_list[x] = curtailment_list[x] + temp_no
            temp_no = df_curtailment_list[x]
        df_block_list = [x for x in range(1,97)]
        df_remark_list = ["" for x in range(1,97)]
        for index, row in date_df.iterrows():
            df_remark_list[row['Block No'] - 1] = df_remark_list[row['Block No'] - 1] + row['Remarks']
        df_date_list = [given_date for x in range(1,97)]
        df_data = {
            'Date' : df_date_list,
            'Block No' : df_block_list,
            'Cumulative Curtailment MW' : df_curtailment_list,
            'Solar cutailment MW': curtailment_list,
            'Remarks': df_remark_list
        }
        print_df = pd.DataFrame(data = df_data)
        final_df = final_df.append(print_df)
    final_df.to_excel(writer,sheet_name='Date_Data',index=False)
