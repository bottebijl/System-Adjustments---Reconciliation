import pandas as pd
import numpy as np
import xlrd

pd.set_option('precision', 2)
pd.options.display.float_format = '{:.2f}%'.format

mapping = pd.read_excel(r'C:\Data\DorenboM\My Documents\Projects\eudw_essb_mapping\mapping_bu_store.xlsx','BU',parse_cols='A:F')
mapping_eudw = (mapping[['eudw_bu_code','BU Name']]).drop_duplicates()
mapping_esb = (mapping[['esb_org_code','BU Name']]).drop_duplicates()

eudw_trans = pd.merge(pd.read_excel(r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\2016\check\Reconciliation Sales and COGS - Transactional Data.xlsx','data'),mapping_eudw,how='left',left_on='Business Unit Code',right_on='eudw_bu_code')
eudw_trans['Fiscal Period Formatted'] = eudw_trans['Fiscal Period Formatted'].apply(lambda x: str(x)[-3:])
eudw_accadj = pd.merge(pd.read_excel(r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\2016\check\Reconciliation Sales and COGS - Accounting Adjustments.xlsx','data'),mapping_eudw,how='left',left_on='Business Unit Code',right_on='eudw_bu_code')
eudw_accadj['Fiscal Period Formatted'] = eudw_accadj['Fiscal Period Formatted'].apply(lambda x: str(x)[-3:])
esb_tot = pd.merge(pd.melt(pd.read_excel(r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\2016\check\essbase_data.xlsx','ess_data',skiprows=1),id_vars=['esb_bu_code','esb_currency','esb_funccur','esb_acct'],var_name='Fiscal Period'),mapping_esb,how='left',left_on='esb_bu_code',right_on='esb_org_code')

sales_eudw_pt1 = pd.pivot_table(eudw_trans,index='BU Name',columns ='Fiscal Period Formatted',values='Sales',aggfunc=np.sum).fillna(0).astype(int)
sales_eudw_pt2 = pd.pivot_table(eudw_accadj,index=['BU Name'],columns='Fiscal Period Formatted',values='Sales',aggfunc=np.sum).fillna(0).astype(int)
sales_esb_pt = pd.pivot_table(esb_tot[esb_tot['esb_acct']=='IS_Sal_Net'],index='BU Name',columns='Fiscal Period',values='value',aggfunc=np.sum).astype(int)
sales_eudw_pt = sales_eudw_pt1 + sales_eudw_pt2

cogs_eudw_pt1 = pd.pivot_table(eudw_trans,index='BU Name',columns ='Fiscal Period Formatted',values='Cogs',aggfunc=np.sum).fillna(0).astype(int)
cogs_eudw_pt2 = pd.pivot_table(eudw_accadj,index=['BU Name'],columns='Fiscal Period Formatted',values='Cogs',aggfunc=np.sum).fillna(0).astype(int)
cogs_eudw_pt3 = pd.pivot_table(eudw_trans,index='BU Name',columns ='Fiscal Period Formatted',values='VFC-SOA',aggfunc=np.sum).fillna(0).astype(int)
cogs_esb_pt = pd.pivot_table(esb_tot[esb_tot['esb_acct']=='IS_CoGS_Gross_Tot'],index='BU Name',columns='Fiscal Period',values='value',aggfunc=np.sum).astype(int)
cogs_eudw_pt = cogs_eudw_pt1 + cogs_eudw_pt2 + cogs_eudw_pt3

writer = pd.ExcelWriter(r'C:\Data\DorenboM\My Documents\Projects\EuDW Gift Costs Restatement\2016\check\EuDW_vs_Essbase_Rec v2.xlsx')

# Set Workbook
workbook  = writer.book

# Set Formats
format1 = workbook.add_format({'num_format': '#,##0'})
format2 = workbook.add_format({'num_format': '0.00%'})

# Add sales sheets
sales_eudw_pt1.fillna(0).to_excel(writer,'EuDW Sales Excl acc adj')
sales_eudw_pt2.fillna(0).to_excel(writer,'EuDW Sales acc adj')
sales_eudw_pt.fillna(0).to_excel(writer,'EuDW Sales Incl acc adj')
sales_esb_pt.fillna(0).to_excel(writer,'Essbase Sales')
(sales_eudw_pt1 - sales_esb_pt).fillna(0).to_excel(writer,'Total Sales Variance excl adj')
((sales_eudw_pt1 - sales_esb_pt)/sales_eudw_pt1).fillna(0).to_excel(writer,'Total Sales Variance excl adj %')
(sales_eudw_pt - sales_esb_pt).fillna(0).to_excel(writer,'Total Sales Variance incl adj')
((sales_eudw_pt - sales_esb_pt)/sales_eudw_pt).fillna(0).to_excel(writer,'Total Sales Variance incl adj %')

# Add cogs sheets
cogs_eudw_pt1.fillna(0).to_excel(writer,'EuDW CoGS Excl acc adj')
cogs_eudw_pt2.fillna(0).to_excel(writer,'EuDW CoGS acc adj')
cogs_eudw_pt.fillna(0).to_excel(writer,'EuDW CoGS Incl acc adj')
cogs_esb_pt.fillna(0).to_excel(writer,'Essbase CoGS')
(cogs_eudw_pt1 - cogs_esb_pt).fillna(0).to_excel(writer,'Total CoGS Variance excl adj')
((cogs_eudw_pt1 - cogs_esb_pt)/cogs_eudw_pt1).fillna(0).to_excel(writer,'Total CoGS Variance excl adj %')
(cogs_eudw_pt - cogs_esb_pt).fillna(0).to_excel(writer,'Total CoGS Variance incl adj')
((cogs_eudw_pt - cogs_esb_pt)/cogs_eudw_pt).fillna(0).to_excel(writer,'Total CoGS Variance incl adj %')

# Loop through sheets and change format
for sht in writer.sheets:
    if sht.find('%',0) != -1:
        writer.sheets[sht].set_column('A:A',33)
        writer.sheets[sht].set_column('B:M',10,format2)
    else:
        writer.sheets[sht].set_column('A:A',33)
        writer.sheets[sht].set_column('B:M',12,format1)
        
# commit changes and save
writer.save()

