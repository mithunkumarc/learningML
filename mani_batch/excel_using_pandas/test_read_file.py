from openpyxl import load_workbook
import xlrd
import pandas as pd

# from Jets_Data loading Customer sheet
dframe = pd.read_excel('Jets_Data.xls',sheet_name='Customer')

# joining Names_first and Names_Second
full_name_list = []
for first,last in zip(list(dframe['Name_First']),list(dframe['Name_Last'])):
    full_name_list.append(first +" "+ last)


# sales_per_seat = total_sales/avg_seats
sales_per_seat = []
for total_sale,avg_seat in zip(list(dframe['Tot_Sales']),list(dframe['Avg_Seats'])):
    sales_per_seat.append((float(total_sale)/avg_seat))


# data to be displayed
required_data_dict = {'CustID':list(dframe['CustID']),
                    'Customer_Name':full_name_list,
                    'Num_Games':list(dframe['Num_Games']),
                    'Avg_Seats':list(dframe['Avg_Seats']),
                    'Tot_Sales':list(dframe['Tot_Sales']),
                    'LastTransYear':list(dframe['LastTransYear']),
                    'Sex':list(dframe['Sex']),
                    'Marital Status':list(dframe['Marital Status']),
                    'Sales_per_seat':sales_per_seat}

# creating data frame for requried data
required_data_frame  = pd.DataFrame(required_data_dict)


#male count
# is_male = required_data_frame['Sex'] == 'Male'
# dframe_male = required_data_frame[is_male]
# print(len(list(dframe_male['CustID'])))

# female count
# is_female = required_data_frame['Sex'] == 'Female'
# dframe_female = required_data_frame[is_female]
# print(len(list(dframe_female['CustID'])))

# Business count
# is_business = required_data_frame['Sex'] == 'Business'
# dframe_business = required_data_frame[is_business]
# print(len(list(dframe_business['CustID'])))

# Finally writing dataframe to output.xlsx
required_data_frame.to_excel("output.xlsx",sheet_name='Customer',index=False)


