import pandas as pd
dil = pd.read_excel('unis_details_dil.xlsx')
ea = pd.read_excel('unis_details_ea.xlsx')
say = pd.read_excel('unis_details_say.xlsx')
söz = pd.read_excel('unis_details_söz.xlsx')

dil.drop(columns=['Unnamed: 0'], inplace=True)
ea.drop(columns=['Unnamed: 0'], inplace=True)
say.drop(columns=['Unnamed: 0'], inplace=True)
söz.drop(columns=['Unnamed: 0'], inplace=True)

joined = pd.concat([dil, ea, say, söz], axis=0)

joined.to_excel('unis_details_AYT.xlsx', index=False)