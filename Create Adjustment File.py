# Import Libraries
import os
import pandas as pd
import numpy as np

#os.chdir(r'C:\Data\DorenboM\Desktop\test')
#from acc_adj_functions import *

# Declare File Addresses
sFile = r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\Adjusted 403.xlsx'
sFile2 = r'C:\Data\DorenboM\My Documents\Projects\Auto Acc Adj Essbase\Phase2\essbase_data.xlsx'
sFile3 = r'C:\Data\DorenboM\My Documents\Projects\eudw_essb_mapping\mapping_bu_store.xlsx'

PERIODS = ['P01','P02','P03','P04','P05','P06','P07','P08','P09','P10','P11','P12']

period_count = 0

for reporting_period in PERIODS:

    period_count = period_count + 1
    # Import Raw Data
    df_eudw_tot = pd.read_excel(sFile,'BU_Data',converters={'Business Unit Code':str,'Date':str})
    df_eudw_tot = df_eudw_tot[df_eudw_tot['Fiscal Period Formatted'].str.contains(reporting_period)]
    df_essb = pd.read_excel(sFile2,'BU_Data',skiprows=2,parse_cols=("A:D"),header=None,names=['BU','Period','esb_sales','esb_cogs'])
    df_essb = df_essb[df_essb['Period']==reporting_period]
    df_map =pd.read_excel(sFile3,'BU',converters={'eudw_bu_code':str})

    # Adjust Raw Data Files
    df_eudw = df_eudw_tot.groupby(['BU_Code_Merge','Local_Currency_Code'])['Sales_Value'].agg(sum).reset_index()

    df_essb = df_essb[(df_essb.esb_sales <>0) | (df_essb.esb_cogs <>0)]

    df_map =pd.read_excel(sFile3,'BU',converters={'eudw_bu_code':str})
    df_map =df_map[df_map['eudw_channel'].isin(['Advantage','Online'])]

    df_tot = pd.merge(df_eudw,df_essb,left_on='BU_Code_Merge',right_on='BU',how='inner',suffixes=['_eudw','_essb'])

    # Filter mapping file on used BU's
    def in_list(x,data_column):
        data_type = type(x)
        varBool = x in data_column.astype(data_type).tolist()
        return varBool

    bool1 = df_map.eudw_bu_code.apply(in_list,data_column=df_eudw_tot['Business Unit Code'])
    bool2 = df_map.esb_org_code.apply(in_list,data_column=df_essb['BU'])

    df_new_map = df_map[bool1 | bool2]

    # Merge Raw Data with Mapping Info
    ds = pd.merge(df_essb,df_map,how='outer',left_on='BU',right_on = 'esb_org_code')
    ds2 = pd.merge(df_eudw_tot,df_map,how='left',left_on='Business Unit Code',right_on = 'eudw_bu_code')

    #Split Sales & Cogs data for one to many mappings
    def split_sales(x):

        nom = float(ds2[ds2['Business Unit Code']==(x['eudw_bu_code'])]['Sales_Value'].sum())
        denom = float(ds2[ds2['esb_org_code']==x['BU']]['Sales_Value'].sum())

        try:

            return ((nom/denom) * x.esb_sales)  

        except ZeroDivisionError:

            if (x.esb_sales) <> 0:    
                return (x.esb_sales) 
            else:
                return 0
        raise

    def split_cogs(x):

        nom = float(ds2[ds2['Business Unit Code']==(x['eudw_bu_code'])]['Cogs_Value'].sum())
        denom = float(ds2[ds2['esb_org_code']==x['BU']]['Cogs_Value'].sum())

        try:

            return ((nom/denom) * x.esb_cogs) 

        except ZeroDivisionError:

            if (x.esb_cogs) <> 0:

                return (x.esb_cogs) 
            else:
                return 0
        raise

    ds['SALES_VALUE'] = ds.apply(split_sales,axis=1)
    ds['COGS_VALUE'] = ds.apply(split_cogs,axis=1)


    # Group Data and return output
    ds2 = ds2.drop_duplicates(subset=['Business Unit Code','Fiscal Period Formatted','Date'])

    gr_esb_sales = ds.groupby(['eudw_bu_code'])['SALES_VALUE'].sum() 
    gr_esb_cogs = ds.groupby(['eudw_bu_code'])['COGS_VALUE'].sum() 
    gr_esb_count = ds.groupby(['eudw_bu_code'])['esb_org_code'].count() 

    grouped = pd.concat([gr_esb_sales,gr_esb_cogs,gr_esb_count],join='outer',axis=1)
    grouped['SALES_VALUE'] =grouped['SALES_VALUE'] / grouped['esb_org_code']
    grouped['COGS_VALUE'] =grouped['COGS_VALUE'] / grouped['esb_org_code']
    grouped = grouped.drop(['esb_org_code'],axis=1) 

    # Group EuDW Data and calculate differences with essbase
    eudw_group = df_eudw_tot.groupby('Business Unit Code')['Sales_Value','Cogs_Value'].sum()
    eudw_group.columns = ['SALES_VALUE', 'COGS_VALUE']

    Acc_Diff_Grouped = (grouped - eudw_group)
    #Acc_Diff_Grouped = (grouped - eudw_group.ix[grouped.index.tolist()])

    output_class = df_eudw_tot[(df_eudw_tot['Business Unit Code'].isin(df_new_map.eudw_bu_code)) & (df_eudw_tot['Working Day Flag']=='Y')][['Business Unit Code','Fiscal Period Formatted','Date','Local_Currency_Code']].reset_index(drop=True)
    output_class['FILE_TYPE'] = 'Forecast Class'
    output_class['SPGS_CLASS'] = '4-44-419'
    output_class['NR_OF_TRANSACTIONS'] =0
    output_class['COMP_SALES_VALUE'] =0
    output_class['COMP_COGS_VALUE'] =0
    output_class['COMP_TRANS_VALUE'] =0

    output_class = pd.merge(output_class,Acc_Diff_Grouped,how='left',left_on='Business Unit Code',right_index=True)

    def split_by_day_sales(x):
        day_count = len(output_class[output_class['Business Unit Code']==x['Business Unit Code']])

        return x.SALES_VALUE / day_count

    def split_by_day_cogs(x):
        day_count = len(output_class[output_class['Business Unit Code']==x['Business Unit Code']])

        return x.COGS_VALUE / day_count

    output_class['SALES_VALUE'] = output_class.apply(split_by_day_sales,axis=1)
    output_class['COGS_VALUE'] = output_class.apply(split_by_day_cogs,axis=1)

    output_class.rename(columns={'Business Unit Code':'BUSINESS_UNIT_CODE','Fiscal Period Formatted':'FISCAL PERIOD','Local_Currency_Code':'CURRENCY_CODE',},inplace=True)
    output_class = output_class[['BUSINESS_UNIT_CODE','FILE_TYPE','FISCAL PERIOD','Date','SPGS_CLASS','CURRENCY_CODE','SALES_VALUE','COGS_VALUE','NR_OF_TRANSACTIONS','COMP_SALES_VALUE','COMP_COGS_VALUE','COMP_TRANS_VALUE']]
    #output_class.to_excel(r'C:\Data\DorenboM\My Documents\Projects\Auto Acc Adj Essbase\Phase2\edw_forecast_BU_'+reporting_period+'.xlsx','data',index=False)
    
    if period_count ==1:
        total_output_class = output_class
    else:
        total_output_class = total_output_class.append(output_class,ignore_index=True)
total_output_class.to_excel(r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\edw_forecast_BU v2.xlsx','data',index=False)