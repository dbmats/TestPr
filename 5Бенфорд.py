import pandas as pd
import xlwings as xw

# pd.options.display.float_format = '{:,.1f}'.format
# pd.set_option('min_rows', 20)
# pd.set_option('max_rows', 100)
# pd.set_option('max_column', 10)
# pd.set_option('max_colwidth', 20)
# pd.set_option('display.width', 1000)

dp = pd.read_csv('PrTest.csv')

# dp = dp.loc[dp['god'] == 2019]
# dp = dp.loc[~dp['x1'].isin([0])]

# tb1 = pd.pivot_table(dp, values=['data'],
#                      index=['x1'],
#                      columns=['god', 'mes'],
#                      aggfunc={'data': 'count'}, margins=True)
# tb1.reset_index(inplace=False)
# tb1 = tb1.fillna(0)
# sht = xw.Book('TestPr.xlsx').sheets['Бенфорд']
# sht.range('B6').options(index=False, header=False).value = tb1
#
# tb2 = pd.pivot_table(dp, values=['data'],
#                    index=['x2'],
#                    columns=['god', 'mes'],
#                    aggfunc={'data': 'count'}, margins=True)
# tb2.reset_index(inplace=True)
# tb2 = tb2.fillna(0)
# sht = xw.Book('TestPr.xlsx').sheets['Бенфорд']
# sht.range('A19').options(index=False, header=False).value = tb2
#
tb = pd.pivot_table(dp, values=['data'],
                    index=['xx'],
                    columns=['god', 'mes'],
                    aggfunc={'data': 'count'}, margins=True)
tb.reset_index(inplace=True)
tb = tb.fillna(0)
# sht = xw.Book('TestPr.xlsx').sheets['Бенфорд']
# sht.range('A33').options(index=False, header=False).value = tb

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = tb

# tb2.to_excel("output2.xlsx")
# print(tb1)
# print(tb1.dtypes)
