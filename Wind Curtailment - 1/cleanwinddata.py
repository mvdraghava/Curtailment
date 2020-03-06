import openpyxl as op
import pandas as pd
import re
from datetime import timedelta, date

global cumulative_curt
global cumulative_percent

cumulative_curt = 0
cumulative_percent = 0

def get_status(remarksstr):
    curtail_status = remarksstr.split(' ')[0]
    if curtail_status == 'curtailed' or curtail_status == 'Curtailed':
        return 1
    elif curtail_status == 'released' or curtail_status == 'Released' or curtail_status == 'releaseed':
        return 0
    else:
        return -1


def data_entry_problems(row):
    ret = ["" for _ in row.index]
    if row['Status'] and row['Wind Curtailment'] > 0:
        return ret
    elif row['Status'] and row['Wind Curtailment'] < 0:
        ret[row.index.get_loc('Solar cutailment MW')] = "background-color: red"
        return ret
    elif not row['Status'] and row['Wind Curtailment'] > 0:
        ret[row.index.get_loc('Solar cutailment MW')] = "background-color: red"
        return ret
    elif not row['Status'] and row['Wind Curtailment'] < 0:
        return ret


def color_negative_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for negative
    strings, black otherwise.
    """
    color = 'red' if val < 0 else 'black'
    return 'color: %s' % color

def get_curtailment_mw(row):
    global cumulative_curt
    global cumulative_percent
    if row['Wind Curtailment'] > 0 :
        curt_mw =  round((row['Wind Curtailment'] * row['Wind generation MW'])/100 , 2)
    else:
        curt_mw = round(row['Wind Curtailment'] * (cumulative_curt / cumulative_percent), 2)
##    if (cumulative_curt + curt_mw) < 0:
##        import pdb
##        pdb.set_trace()
    cumulative_curt = round(cumulative_curt + curt_mw , 2)
    cumulative_percent = cumulative_percent + row['Wind Curtailment']
    return curt_mw


start_dt = date(2019, 9, 24)
end_dt = date(2019, 12, 31)
date_list = [(start_dt + timedelta(n)).strftime('%d-%m-%Y') for n in range(int ((end_dt - start_dt).days)+1)]


df = pd.read_excel('apwindcurtialment.xlsx')
df.rename(columns = {'Reasons for Backing down': 'Remarks'},inplace = True)
df['Status']  = df['Remarks'].apply(get_status)
df['Wind Curtailment MW'] = df.apply(get_curtailment_mw,axis=1)


df.style.apply(data_entry_problems,axis=1)
##with pd.ExcelWriter('output.xlsx',engine='openpyxl') as writer:
##    df.style.apply(data_entry_problems,axis=1).to_excel(writer, sheet_name='Data')

final_df = pd.DataFrame(columns=['Date',
                                 'Block No',
                                 'Cumulative Curtailment MW',
                                 'Wind Curtailment MW',
                                 'Curtailement Residual MW',
                                 'Cumulative Curtailment Percent',
                                 'Wind Curtailment Percent',
                                 'Wind Curtailment Percent Seperated',
                                 'Wind generation MW',
                                 'Remarks'])
with pd.ExcelWriter('BlockWiseData.xlsx',engine='openpyxl') as writer:
    temp_no = 0
    mw_temp = 0
    one_percent_curt = 0
    for given_date in date_list:
        date_df = df[df['Date'] == given_date]

        df_date_list = [given_date for x in range(1,97)]  #Date Column

        df_block_list = [x for x in range(1,97)]          #Block Column

        curtailment_list = [0 for x in range(1,97)]       #Wind Curtailment Percent Column
        for index, row in date_df.iterrows():
                curtailment_list[row['Block No'] - 1] = curtailment_list[row['Block No'] - 1] + row['Wind Curtailment']

        df_curtailment_list = [0 for x in range(1,97)]    #Cumulative Wind Curtailment Percent Column
        for x in range(len(curtailment_list)):
            df_curtailment_list[x] = curtailment_list[x] + temp_no
            temp_no = df_curtailment_list[x]

        
        mw_curtailment_list = [0 for x in range(1,97)]    #Wind Curtailment MW Column
        for index, row in date_df.iterrows():
                 mw_curtailment_list[row['Block No'] - 1] = mw_curtailment_list[row['Block No'] - 1] + row['Wind Curtailment MW'] 

        residual_curtailement = [0 for x in range(1,97)]  #Residual Column
        
        df_mw_curtailment_list = [0 for x in range(1,97)] #Cumulative Wind Curtailment MW Column with released as zero
        for x in range(len(curtailment_list)):
            df_mw_curtailment_list[x] = round(mw_curtailment_list[x] + mw_temp, 2)
            mw_temp = df_mw_curtailment_list[x]
            if df_curtailment_list[x] == 0:
                mw_temp = 0
                residual_curtailement[x] = df_mw_curtailment_list[x]
                if residual_curtailement[x] != 0:
                    print(residual_curtailement[x])
                df_mw_curtailment_list[x] = 0

        
        df_wind_total_gen = ["" for x in range(1,97)]     #Wind Generation MW Column
        for index, row in date_df.iterrows():
            if not df_wind_total_gen[row['Block No'] - 1]:
                df_wind_total_gen[row['Block No'] - 1] = str(row['Wind generation MW'])
            else:
                df_wind_total_gen[row['Block No'] - 1] = df_wind_total_gen[row['Block No'] - 1] + ' , ' + str(row['Wind generation MW'])

        seperate_wind_curtail_percent = ["" for x in range(1,97)] #Wind Curtailment Seperated Column
        for index, row in date_df.iterrows():
            if not seperate_wind_curtail_percent[row['Block No'] - 1]:
                seperate_wind_curtail_percent[row['Block No'] - 1] = str(row['Wind Curtailment'])
            else:
                seperate_wind_curtail_percent[row['Block No'] - 1] = seperate_wind_curtail_percent[row['Block No'] - 1] + ' , ' + str(row['Wind Curtailment'])

        df_remark_list = ["" for x in range(1,97)]       #Remarks Column
        for index, row in date_df.iterrows():
            df_remark_list[row['Block No'] - 1] = df_remark_list[row['Block No'] - 1] + row['Remarks']
                
        df_data = {
            'Date' : df_date_list,
            'Block No' : df_block_list,
            'Cumulative Curtailment MW' : df_mw_curtailment_list,
            'Wind Curtailment MW': mw_curtailment_list,
            'Curtailement Residual MW':residual_curtailement,
            'Cumulative Curtailment Percent' : df_curtailment_list,
            'Wind Curtailment Percent': curtailment_list,
            'Wind Curtailment Percent Seperated': seperate_wind_curtail_percent,
            'Wind generation MW': df_wind_total_gen,
            'Remarks': df_remark_list
        }
        print_df = pd.DataFrame(data = df_data)
        final_df = final_df.append(print_df)
    final_df.to_excel(writer,sheet_name='Date_Data',index=False) 
