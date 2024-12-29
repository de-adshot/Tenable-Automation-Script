from openpyxl import load_workbook
import shutil
import pandas as pd

primary_file = input("Exter the file name of the initial assessment:\t")
primary_df = pd.read_excel(primary_file)
excel_length=len(primary_df)
#print(excel_length)
secondary_file = input("Exter the file name of the secondary assessment:\t")
secondary_df = pd.read_excel(secondary_file)

result_file = 'revalidation_' + primary_file
shutil.copy(primary_file,result_file)
booki = load_workbook(result_file)

resultant_df = pd.read_excel(result_file)

for p_rw in range(0,excel_length):
    print("row number: ",p_rw)
    primary_ip = primary_df.iloc[p_rw, 2]
    primary_vname = primary_df.iloc[p_rw, 0]
    column_name = 'Plugin Name'
    matching_rows = secondary_df[secondary_df[column_name] == primary_vname]
    #print(matching_rows)
    #print(type(matching_rows))
    secondary_index = matching_rows.index
    secondary_index = secondary_index.tolist()
    #print(secondary_index)
    if len(secondary_index) == 0:
     #   print(type(secondary_index))
        print("Empty secondary index")
        print('Primary IP is',primary_ip)
        print('Vulnerability:',primary_vname)
        print("Status value: ", 'Closed')
        resultant_df.loc[(resultant_df['IP Address'] == primary_ip) & (resultant_df[column_name] == primary_vname), 'Revalidation Status'] = 'Closed'
        resultant_df.to_excel(result_file, index=False)
    else:
        print("Length of secondary_index is:",len(secondary_index))
        temp_list = []
        for rw in secondary_index:
            #print("row number: ",rw)
            secondary_ip = secondary_df.iloc[rw, 4]
            #secondary_vname = secondary_df.iloc[rw, 1]
            #print(secondary_df.iloc[rw])
            #print("Adding secondary_ip to list: ",secondary_ip)
            temp_list.append(secondary_ip)
        print('Length of Temp List:',len(temp_list))
        #print(temp_list)
        print('Primary IP is',primary_ip)
        print('Vulnerability:',primary_vname)
        if primary_ip in temp_list:
            #print('Primary IP is',primary_ip)
            #print('Vulnerability:',primary_vname)
            print("Status value: ", 'Open')
            #col_insert(status_value,p_rw,primary_ip,primary_vname,secondary_ip,secondary_vname)
            resultant_df.loc[(resultant_df['IP Address'] == primary_ip) & (resultant_df['Plugin Name'] == primary_vname), 'Revalidation Status'] = 'Open'
            resultant_df.to_excel(result_file, index=False)
        else:
            print("Status value: ", 'Closed')
            resultant_df.loc[(resultant_df['IP Address'] == primary_ip) & (resultant_df[column_name] == primary_vname), 'Revalidation Status'] = 'Closed'
            resultant_df.to_excel(result_file, index=False)
    print('++++++++++++++++++++++++++++++++++++++++++++++++++++++')    
print(f"\nThe result file is {result_file}")