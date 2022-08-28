import pandas as pd
import numpy as np
import re
import openpyxl


def data_frame_from_xlsx(xlsx_file, range_name):
    """ Get a single rectangular region f`rom the specified file.
    range_name can be a standard Excel reference ('Sheet1!A2:B7') or 
    refer to a named region ('my_cells')."""
    wb = openpyxl.load_workbook(xlsx_file, data_only=True, read_only=True)
    if '!' in range_name:
        # passed a worksheet!cell reference
        ws_name, reg = range_name.split('!')
        if ws_name.startswith("'") and ws_name.endswith("'"):
            # optionally strip single quotes around sheet name
            ws_name = ws_name[1:-1]
        region = wb[ws_name][reg]
    else:
        # passed a named range; find the cells in the workbook
        full_range = wb.get_named_range(range_name)
        if full_range is None:
            raise ValueError(
                'Range "{}" not found in workbook "{}".'.format(range_name, xlsx_file)
            )
        # convert to list (openpyxl 2.3 returns a list but 2.4+ returns a generator)
        destinations = list(full_range.destinations) 
        if len(destinations) > 1:
            raise ValueError(
                'Range "{}" in workbook "{}" contains more than one region.'
                .format(range_name, xlsx_file)
            )
        ws, reg = destinations[0]
        # convert to worksheet object (openpyxl 2.3 returns a worksheet object 
        # but 2.4+ returns the name of a worksheet)
        if isinstance(ws, str):
            ws = wb[ws]
        region = ws[reg]
    # an anonymous user suggested this to catch a single-cell range (untested):
    # if not isinstance(region, 'tuple'): df = pd.me2lFrame(region.value)
    df = pd.DataFrame([cell.value for cell in row] for row in region)
    return df

'''ME2L file'''

file = r"EXPORT.xlsx" #this is me2l file 

me2l = pd.read_excel(file,header = 0, dtype={"Purchasing Document":"string","Item":"string","Material Group":"string"}) #dtype prevents auto casting
# imported = imported.astype({"Purchasing Document":"string","Item":"string","Material Group":"string", "Vendor ID":"string"})
# me2l = file_read.iloc[1:,:] #don't take the first row
# me2l = me2l.astype({"Purchasing Document":"string","Item":"string","Material Group":"string"})

#to convert to string
me2l['Purchasing Document'] = me2l['Purchasing Document'].str.replace('\.\d+$','',regex = True)
me2l['Item'] = me2l['Item'].str.replace('\.\d+$','',regex = True)
me2l['Material Group'] = me2l['Material Group'].str.replace('\.\d+$','',regex = True).apply('{:0>8}'.format)



#vendor ID 
me2l.insert(9,'Vendor ID','') #insert at index 9 
me2l[['Vendor ID','Name of Vendor']] = me2l['Name of Vendor'].str.split(" ",1,expand = True)
me2l['Vendor ID'] = me2l['Vendor ID'].str.replace('\.\d+$','',regex = True)

#calculations 
me2l['PO Value'] = me2l['Order Quantity']*me2l['Net price']
me2l['PO Value (£m)'] =  me2l['PO Value']/(10**6)
# me2l['Concatenate of PO number & line item'] = me2l['Purchasing Document'] + me2l['Item']
#IINSERT CONCATENATION OF SAP CATEGORY CODE AND VENDOR ID

#selected cols, note outstanding value = still to delivered 
cols = ['Purchasing Document','Item','Purchasing Doc. Type','Requisitioner Name',
        'Document Date','Name of Vendor','Short Text','Material Group', 'Vendor ID','Still to be delivered (value)','PO Value','PO Value (£m)']
me2l = me2l.loc[:,cols]

me2l.head(5)

'''Data connections'''
'''
Abstract:
1. Table1[B:I] ¬ SpendCategoryCode => used for getting levels + gl account mapping + gl account desc
2. Table2[AN:AS] ¬ Concatenation => new desc + new code (ar:as)
3. MasterContracts[C:E] ¬ gets FA number and cleansed contract description using contract ID
4. FramwWorkLeakage[O:P] ¬ sapcode + vendor ID => FA Number 
'''
SAPCategories = r"lookupTbls\w_SAP_master_mapping.xlsx"
#read master po workbook for - good recepient, good approver, wbs element, contract id 
wbs_costcentre = pd.read_excel("lookupTbls\\w_MASTER PO DATA April 21 onwards.xlsx", usecols = "A:M",skiprows= 1)
wbs_costcentre =  wbs_costcentre.astype({'Purch.Doc.':"string",'Item':"string"})

spend_cat = data_frame_from_xlsx(SAPCategories, "Lookup tables!$B$5:$I$1000") #grey table 
spend_cat.columns = spend_cat.iloc[0]
spend_cat = spend_cat.iloc[1:,]
spend_cat = spend_cat.astype({'Spend Category Code V3.04':"string"})
spend_cat['Spend Category Code V3.04'] = spend_cat['Spend Category Code V3.04'].apply('{:0>8}'.format)

#second talbe in lookup
mapping_cat = data_frame_from_xlsx(SAPCategories, "Lookup tables!$AN$5:$AS$1000") #brown table - keeps changing columns???
mapping_cat.columns = mapping_cat.iloc[0]
mapping_cat = mapping_cat.iloc[1:,:].astype({'SAP Category L4 code':'string'})
mapping_cat.drop_duplicates(subset = ['SAP Category L4 code'], inplace = True)
mapping_cat['SAP Category L4 code'].fillna(method = 'pad', inplace = True)
mapping_cat['SAP Category L4 code'] = mapping_cat['SAP Category L4 code'].apply('{:0>8}'.format)

'''TRANSFORMATIONS'''
opex_capex = pd.merge(me2l,wbs_costcentre, left_on = ['Purchasing Document','Item'],right_on=['Purch.Doc.','Item'],how = 'left')
opex_capex['Capex/Opex'] = np.where(opex_capex['WBS Element'] == "#", 'OPEX', (np.where(opex_capex['WBS Element'] == np.nan, "OPEX", 'CAPEX')))

po_item = pd.merge(opex_capex,spend_cat, left_on = 'Material Group',right_on = 'Spend Category Code V3.04', how = 'left')

po_item[['Spend Category Description (Long 50 char)']].fillna('Uncategorised',inplace = True) #part of guide 
#po_item['Concatenation'] = po_item['Vendor ID'] + po_item['Spend Category Code V3.04'] #created to join on sap category level 4 mapping
#using previous composite key to get code and desc
po_item = po_item.merge(mapping_cat, left_on= 'Material Group',right_on = 'SAP Category L4 code', how = 'left')

#move this line to end as it covers GL + levels 
po_item.rename(columns = {'NEW SAP Category L4 description':'Cleansed SAP L4 category description','SAP Category L4 code':'Cleansed SAP L4 category code',
                          'Spend Category Description (Long 50 char)':'Base_SAP_Category_L4_Description'}, inplace = True)
po_item['Cleansed SAP L4 category code'].fillna(po_item['Material Group'], inplace = True)

#Goods recipient and approver
po_item[['Goods Recipient','Approver']].fillna('not assigned',inplace=True)

#we get list of names to narrow down our relevant column name
fa_1= po_item.merge(ContractID_FA,how='left',left_on='Purchasing Document',right_on='Contract ID')

#Getting fa number using framework leakage tab 
po_item = po_item.merge(Frameworkleakage_FA, left_on=['Cleansed SAP L4 category code','Vendor ID_y'],right_on=['SAP Prodcut category code','VENDOR ID'], how = 'left')


'''FINAL FILE'''


# #Selected columns 
cols = ['Purchasing Document', 'Item', 'Purchasing Doc. Type',
       'Requisitioner Name', 'Document Date','Vendor ID_x', 'Name of Vendor_x',
       'Short Text_x', 'Material Group', 'Still to be delivered (value)',
       'PO Value', 'PO Value (£m)', 'WBS Element', 'Capex/Opex', 'Base_SAP_Category_L4_Description',
       'Cleansed SAP L4 category code', 'Cleansed SAP L4 category description',
       'Level 1 Description', 'Level 2 Description', 'Level 3 Description', 'GL account mapping',
       'GL Account Description','Goods Recipient',
       'Approver','Contract ID','fa_ta_number'
       ]
comp = po_item[cols]

comp['Document Date'] = comp['Document Date'].dt.strftime('%d/%m/%Y')

#final lookups to complete fa number set 
fa_using_contractID = comp.merge(ContractID_FA, on= 'Contract ID', how='left')  #for some reason this join doesn't seem to work
#fa_using_code_vendorID = fa_using_contractID.merge()

#conforming column names 
# renamedCols = {'PO doc type_x': 'PO doc type', 'Vendor ID_x':'Vendor ID','Name of Vendor_x':'Name of Vendor','Short Text_x':'Short Text','Contract ID_y':'Contract ID',
#                'GL account mapping':'GL account code','Approver':'Goods Approver'}

# comp.rename(columns = renamedCols, inplace= True)




