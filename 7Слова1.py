import re
import pandas as pd
import xlwings as xw
pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# анализ слов при анализе выручки, К60,76, оплате поставщикам

dp = pd.read_csv('PrTest.csv')
dp = dp.fillna(0)
dp['text'] = dp['text'].astype('str')

dp1 = dp.loc[dp['DK'] == '62;90']
dp1 = pd.Series(dp1['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb = pd.pivot_table(dd, values=['x'],
                     index=0,
                     aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb.reset_index(inplace=True)

dp2 = dp.loc[dp['DK'].isin(['20;60', '23;60', '25;60', '26;60', '44;60',
                            '20;76', '23;76', '25;76', '26;76', '44;76', '32;60', '32;76'])]
dp1 = pd.Series(dp2['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb1 = pd.pivot_table(dd, values=['x'],
                     index=0,
                     aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb1.reset_index(inplace=True)

dp3 = dp.loc[dp['DK'].isin(['60;51', '60;52', '76;51', '76;52'])]
dp1 = pd.Series(dp3['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb2 = pd.pivot_table(dd, values=['x'],
                     index=0,
                     aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb2.reset_index(inplace=True)

sht = xw.Book('TestPr.xlsx').sheets['слова1']
sht.range('A14').options(index=False, header=False).value = tb
sht.range('D14').options(index=False, header=False).value = tb1
sht.range('G14').options(index=False, header=False).value = tb2

# print(tb2)
# print(dd[['0']])
# print(dp)
# print(dd.dtypes)