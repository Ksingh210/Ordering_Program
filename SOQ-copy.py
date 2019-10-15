def truck_size():
    '''
    Function purpose is to assign a global variable "sizes" which is an empty list that appends 
    values that correspond with company's truck limits. An input question is asked for full 
    truck limits or half truck limits. 
    
    '''

    select = input("What are the truck sizes? (full/half): ").lower()
    
    global sizes

    sizes = []
    
    if select == 'full':
        sizes.append(50)
        sizes.append(60)
        sizes.append(70)
    else:
        sizes.append(25)
        sizes.append(30)
        sizes.append(35)

def main():
    '''
    Function is used to import assortment, sales, and inventory files. Once imported, the sales
    and inventory files are merged onto the assortment file as a new variable "SOQ". Following 
    the merge, the data within SOQ is divided by the truck sizes (copies are made). A while loop
    is executed on each copy. The while loop is dependant on the input from the function "truck_size"
    so if the input was full the while loop runs on full truck quantites and if the input was half
    the while loop runs on half truck quantities. The seperated data is then combined back together
    and exported to an excel file.

    '''
    # Date for report
    report_date = str(input("Enter today's date (Month.Day): "))

    print('\nProcessing....')

    # Import required files
    AMT = pd.read_excel(r'#')
    SALES = pd.read_excel(r'#')
    INVENTORY = pd.read_excel(r'#')

    SOQ = AMT.merge(INVENTORY,
                    how = 'left',
                    left_on = ['STR NBR', 'Sku'],
                    right_on = ['Store Nbr', 'SKU Nbr']).fillna(0)

    SOQ.drop(columns = ['Store Nbr', 'SKU Name', 'SKU Nbr'], inplace = True)

    SOQ = SOQ.merge(SALES,
                    how = 'left',
                    left_on = ['STR NBR', 'Sku'],
                    right_on = ['Store Nbr', 'SKU Nbr']).fillna(0)

    SOQ.drop(columns = ['Store Nbr', 'SKU Nbr'], inplace = True)

    SOQ[['Str OH Units Dly', 'Str OO Units Dly', 'SALES Units']] = SOQ[['Str OH Units Dly', 'Str OO Units Dly', 'SALES Units']].astype('int')

    SOQ['AWS'] = SOQ['SALES Units']/4

    # Lambda expression used to fill replace values less than 1 with 1 for computation
    SOQ['AWS'] = SOQ['AWS'].apply(lambda x: 1 if x < 1 else x)

    SOQ['SOQ PLT'] = 0
    SOQ['SOQ Units'] = SOQ['SOQ PLT']*SOQ['PLT SZ']
    SOQ['AD Units'] = SOQ['Str OH Units Dly'] + SOQ['Str OO Units Dly'] + SOQ['SOQ Units']
    SOQ['WOS'] = SOQ['AD Units'] / SOQ['AWS']

    SOQ2 = SOQ[['STR NBR', 'Sku', 'WOS']]
    SOQ2 = pd.DataFrame.groupby(SOQ2,
                                by = ['STR NBR'],
                                sort = True).rank(method = 'min')

    SOQ = SOQ.join(SOQ2,
                how = 'left',
                rsuffix='x').fillna(0)

    SOQ.drop(columns = 'Skux', inplace = True)

    SOQ['WOSx'] = SOQ['WOSx'].astype('int')
    
    # Renaming columns
    SOQ.columns = ['Regions','BYO', 'MKT',
                'STR', 'Sku', 'Old Sku',
                'Sku Description', 'AMT',
                'PLT SZ', 'WH', 'Truck Limit',
                'STR OH', 'STR OO', 'SALES Units',
                'AWS', 'SOQ PLT', 'SOQ Units',
                'AD Units', 'WOS', 'Rank']

    # Running While Loop Unitl 25 or 50 Pallet Truck Quantities Are Reached
    SOQ_50 = SOQ.loc[SOQ['Truck Limit'] == 50].copy()

    x = 0

    while x < sizes[0]:
        
        SOQ_50['SOQ PLT'] = SOQ_50[['SOQ PLT', 'Rank']].apply((lambda x: x['SOQ PLT']+1 if x['Rank'] == 1 else x['SOQ PLT']),axis =1)

        SOQ_50['SOQ Units'] = SOQ_50['SOQ PLT']*SOQ_50['PLT SZ']
        SOQ_50['AD Units'] = SOQ_50['STR OH'] + SOQ_50['STR OO'] + SOQ_50['SOQ Units']
        SOQ_50['WOS'] = SOQ_50['AD Units'] / SOQ_50['AWS']

        SOQ2 = SOQ_50[['STR', 'Sku', 'WOS']]
        SOQ2 = pd.DataFrame.groupby(SOQ2, by = ['STR'],sort = True).rank(method = 'min')

        SOQ_50 = SOQ_50.join(SOQ2, how = 'left', rsuffix='x').fillna(0)

        SOQ_50.drop(columns = 'Skux', inplace = True)
        SOQ_50['Rank'] = SOQ_50['WOSx'].astype('int')
        SOQ_50.drop(columns = 'WOSx', inplace = True)
        
        x += 1
        
    # Running While Loop Unitl 30 or 60 Pallet Truck Quantities Are Reached
    SOQ_60 = SOQ.loc[SOQ['Truck Limit'] == 60].copy()

    x = 0

    while x < sizes[1]:
        
        SOQ_60['SOQ PLT'] = SOQ_60[['SOQ PLT', 'Rank']].apply((lambda x: x['SOQ PLT']+1 if x['Rank'] == 1 else x['SOQ PLT']),axis =1)

        SOQ_60['SOQ Units'] = SOQ_60['SOQ PLT']*SOQ_60['PLT SZ']
        SOQ_60['AD Units'] = SOQ_60['STR OH'] + SOQ_60['STR OO'] + SOQ_60['SOQ Units']
        SOQ_60['WOS'] = SOQ_60['AD Units'] / SOQ_60['AWS']

        SOQ2 = SOQ_60[['STR', 'Sku', 'WOS']]
        SOQ2 = pd.DataFrame.groupby(SOQ2, by = ['STR'],sort = True).rank(method = 'min')

        SOQ_60 = SOQ_60.join(SOQ2, how = 'left', rsuffix='x').fillna(0)

        SOQ_60.drop(columns = 'Skux', inplace = True)
        SOQ_60['Rank'] = SOQ_60['WOSx'].astype('int')
        SOQ_60.drop(columns = 'WOSx', inplace = True)
        
        x += 1
        
    # Running While Loop Unitl 35 or 70 Pallet Truck Quantities Are Reached
    SOQ_70 = SOQ.loc[SOQ['Truck Limit'] == 70].copy()

    x = 0

    while x < sizes[2]:
        
        SOQ_70['SOQ PLT'] = SOQ_70[['SOQ PLT', 'Rank']].apply((lambda x: x['SOQ PLT']+1 if x['Rank'] == 1 else x['SOQ PLT']),axis =1)

        SOQ_70['SOQ Units'] = SOQ_70['SOQ PLT']*SOQ_70['PLT SZ']
        SOQ_70['AD Units'] = SOQ_70['STR OH'] + SOQ_70['STR OO'] + SOQ_70['SOQ Units']
        SOQ_70['WOS'] = SOQ_70['AD Units'] / SOQ_70['AWS']

        SOQ2 = SOQ_70[['STR', 'Sku', 'WOS']]
        SOQ2 = pd.DataFrame.groupby(SOQ2, by = ['STR'],sort = True).rank(method = 'min')

        SOQ_70 = SOQ_70.join(SOQ2, how = 'left', rsuffix='x').fillna(0)

        SOQ_70.drop(columns = 'Skux', inplace = True)
        SOQ_70['Rank'] = SOQ_70['WOSx'].astype('int')
        SOQ_70.drop(columns = 'WOSx', inplace = True)
        
        x += 1

    # For 50 Pallet Orders
    SOQ3 =  SOQ_50.groupby(by = 'STR')
    Total = SOQ3['SOQ PLT'].sum()

    Total = pd.Series.to_frame(Total)

    SOQ_50 = SOQ_50.merge(Total, how = 'left',
                        on = 'STR', suffixes=('', ' Total'))

    # For 60 Pallet Orders
    SOQ3 =  SOQ_60.groupby(by = 'STR')
    Total = SOQ3['SOQ PLT'].sum()

    Total = pd.Series.to_frame(Total)

    SOQ_60 = SOQ_60.merge(Total, how = 'left',
                        on = 'STR', suffixes=('', ' Total'))

    # For 70 Pallet Orders
    SOQ3 =  SOQ_70.groupby(by = 'STR')
    Total = SOQ3['SOQ PLT'].sum()

    Total = pd.Series.to_frame(Total)

    SOQ_70 = SOQ_70.merge(Total, how = 'left',
                        on = 'STR', suffixes=('', ' Total'))

    # Change File Name Every Time Code Is Run
    export_file = pd.ExcelWriter(r'# '+report_date+' SOQ Reports.xlsx', engine = 'xlsxwriter')

    SOQ_50.to_excel(export_file,
                    sheet_name= 'SOQ_50',
                    index = False)
    SOQ_60.to_excel(export_file,
                    sheet_name= 'SOQ_60',
                    index = False)
    SOQ_70.to_excel(export_file,
                    sheet_name= 'SOQ_70',
                    index = False)

    export_file.save()
    print('\nSOQ Executed! Check SOQ Reports.')

if __name__ == '__main__':
    import pandas as pd
    truck_size()
    main()