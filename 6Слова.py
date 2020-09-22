import pandas as pd
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# выборка проводок со словами из списка

# dp = pd.read_csv('PrTest.csv')
# dp = dp.fillna(0)
#
# tt1 = pd.read_csv('Slova.csv')
# tt = pd.Series(tt1['слова'].values, name='Value')
#
# dp = dp[['data', 'dok', 'org',
#          'schD', 'cD1', 'cD2','cD3',
#          'schK', 'cK1','cK2', 'cK3',
#          'sum','text', 'schD2', 'schK2', 'DK']]
# dp['text'] = dp['text'].astype('str')
# dp['text'] = dp.text.str.lower()  # в строчные буквы
#
# dd = pd.DataFrame(columns=['data', 'dok', 'org',
#                            'schD', 'cD1', 'cD2','cD3',
#                            'schK', 'cK1','cK2', 'cK3',
#                            'sum', 'text', 'schD2', 'schK2', 'DK'])  # создание пустой заготовки
# dd['slovo'] = str()
# for ind, val in enumerate(tt):
#     dd1 = dp[dp['text'].str.contains(val)]
#     dd1['slovo'] = (val)
#     dd = dd.append(dd1)
# dd['abs'] = dd['sum'].abs()

# dd.to_csv('НайденныеСлова.csv', index=False)

# tb1 = pd.pivot_table(dd, values=['sum', 'data'],
#                      index=['slovo'],
#                      aggfunc={'sum': sum, 'data': 'count'})
# tb1.reset_index(inplace=True)
# tb1 = tb1.fillna(0)

# sht = xw.Book('TestPr.xlsx').sheets['слова']
# sht.range('A13').options(index=False, header=False).value = tb1

dd = pd.read_csv('НайденныеСлова.csv')

tb1 = pd.pivot_table(dd, values=['sum', 'data'],
                     index=['slovo', 'org', 'cD1', 'cK1', 'text', 'schD2', 'schK2', 'DK'],
                     aggfunc={'sum': sum, 'data': 'count'})
tb1.reset_index(inplace=True)
tb1 = tb1.fillna(0)

# sht = xw.Book('TestPr.xlsx').sheets['словаПР']
# sht.range('A6').options(index=False, header=False).value = dd

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = tb1

print(tb1)
# print(dd)
# print(dp.head(15))
# print(dd.shape)