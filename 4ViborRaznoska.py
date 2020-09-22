import pandas as pd
import xlwings as xw


dp = pd.read_csv('выбранные.csv')
dp = dp[['data', 'dok', 'org',
         'schD', 'cD1', 'cD2', 'cD3',
         'schK', 'cK1', 'cK2', 'cK3',
         'sum', 'abs', 'text',
         'schD2', 'schK2', 'DK',
         'krr', 'vixodnoy', 'notch',
         'dubl', 'otritz', 'vir', 'redko']]


dp1 = dp.loc[dp['krr'] == 1]
sht = xw.Book('TestPr.xlsx').sheets['1']
sht.range('A13').options(index=False, header=False).value = dp1

dp['pp'] = dp['vixodnoy']+dp['notch']
dp2 = dp.loc[dp['pp'].isin([1, 2])]
sht = xw.Book('TestPr.xlsx').sheets['2']
sht.range('A13').options(index=False, header=False).value = dp2

dp3 = dp.loc[dp['dubl'] == 1]
sht = xw.Book('TestPr.xlsx').sheets['3']
sht.range('A13').options(index=False, header=False).value = dp3

dp4 = dp.loc[dp['dubl'] == 0]
dsht = xw.Book('TestPr.xlsx').sheets['4']
sht.range('A13').options(index=False, header=False).value = dp4

dp5 = dp.loc[dp['krr'] == 0]
dp5 = dp5.loc[dp5['vixodnoy'] == 0]
dp5 = dp5.loc[dp5['notch'] == 0]
dp5 = dp5.loc[dp5['dubl'] == 0]
dp5 = dp5.loc[dp5['otritz'] == 0]
dp5 = dp5.loc[dp5['vir'] == 1]
sht = xw.Book('TestPr.xlsx').sheets['5']
sht.range('A13').options(index=False, header=False).value = dp5

dp6 = dp.loc[dp['redko'] == 1]
sht = xw.Book('TestPr.xlsx').sheets['6']
sht.range('A13').options(index=False, header=False).value = dp6

dp7 = dp.loc[dp['krr'] == 0]
dp7 = dp7.loc[dp7['vixodnoy'] == 0]
dp7 = dp7.loc[dp7['notch'] == 0]
dp7 = dp7.loc[dp7['dubl'] == 0]
dp7 = dp7.loc[dp7['otritz'] == 0]
dp7 = dp7.loc[dp7['vir'] == 0]
dp7 = dp7.loc[dp7['redko'] == 0]
sht = xw.Book('TestPr.xlsx').sheets['7']
sht.range('A13').options(index=False, header=False).value = dp7