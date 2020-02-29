import pandas as pd
import string

df = pd.read_excel('Fusion_Current_Country_Restrictions__2_28_20.xlsx')
# print(df)

df['Restrictions'] = ""
df['Country_Codes_Restrictions'] = ""


df.rename(columns={'Country Restrictions :  Country/Region' : 'Country_Restriction_Code', 'Country Restrictions :  '
                                                                                          'Restriction' :
    'Shipping_Restriction','Item Status' : 'Item_Status', 'Item Name' : 'Item_Name', 'Item Description' :
    'Item_Description'}
          , inplace=True)

# Removing the nans(from blank cells) from both columns. So that .apply(','.join) can be used against the result columns with no errors.
df['Country_Restriction_Code'] = df.Country_Restriction_Code.fillna('')
df['Shipping_Restriction'] = df.Shipping_Restriction.fillna('')

# If a part number has Item_Status = Obsolete, then filter out all occurances of that PN.
s = 0
obPN = 'l'
pn_Set = set([])
for s in range(len(df)):

    if df.iloc[s, 3] == 'Obsolete':
        obPN = df.iloc[s, 0]
        pn_Set.add(obPN)
        # df = df.loc[df['Item_Name'] != obPN]
        # print(df.head())

p = list(pn_Set)
df = df[~df['Item_Name'].isin(p)]

df.reset_index(drop=True,inplace=True)

ccodes = set()
Index_label = []
codelist = []
restrict_string = set()

for i in range(len(df)):
    ccodes.add(df.iloc[i, 0])
    # restrict_string.add(df.iloc[i, 7])


for setPn in ccodes:
    Index_label = df[df['Item_Name'] == setPn].index.tolist()
    for idx in Index_label:
        codelist.append(df.iloc[idx,5])
        restrict_string.add(df.iloc[idx, 7])
        if idx == Index_label[-1]:
            restrict_string = list(restrict_string)
            df.at[idx, 'Country_Codes_Restrictions'] = codelist
            df.at[idx, 'Restrictions'] = restrict_string
            codelist = []
            restrict_string = set()


# Using .apply(','.join) to join all strings in the result columns and leave out all of the brackets and quote marks.
df['Country_Codes_Restrictions'] = df['Country_Codes_Restrictions'].apply(','.join)
df['Restrictions'] = df['Restrictions'].apply(','.join)

# Remove first character in string for each row in the Restrictions column.
df['Country_Codes_Restrictions'] = df['Country_Codes_Restrictions'].str.strip(',')
df['Restrictions'] = df['Restrictions'].str.strip(',')

df = df.dropna(how='any', subset=['Country_Codes_Restrictions'])

# Send the dataframe to an Excel file.
df.to_excel('Fusion_Country_Restrictions_Results_2_28_20_v2.0.xlsx', index=False)

