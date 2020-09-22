import numpy as np
import pandas as pd
import xlwings as xw
pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# Аналитические процедуры

dp = pd.read_csv('PrTest.csv')
dp = dp.fillna(0)

# # начисление амортизации
dp1 = dp.loc[dp['schK2'] == 2]
tb2 = pd.pivot_table(dp1, values=['sum'],
                    index=['cK1'],
                    aggfunc={'sum': sum},
                    columns=['god', 'mes'], margins=True)
tb2.reset_index(inplace=True)
tb2 = tb2.fillna(0)

# закупки К60,76
dp3 = dp.loc[dp['schK2'].isin([60, 76])]
dp3 = dp3.loc[dp3['schD2'].isin([8, 7, 10, 15, 20, 23, 25, 26, 29, 41, 44, 91, 97, '08', '07'])]
tb4 = pd.pivot_table(dp3, values=['sum'],
                     index=['schD2', 'schK2', 'cD1', 'cD2', 'cD3', 'cK1', 'cK2' , 'god', 'mes'],
                     aggfunc={'sum': sum},)
tb4.reset_index(inplace=True)
tb4 = tb4.fillna(0)

# доходы Д62К90
dp4 = dp.loc[dp['schD2'].isin([62])]
dp4 = dp4.loc[dp4['schK2'].isin([90])]
tb5 = pd.pivot_table(dp4, values=['sum'],
                     index=['cD1', 'cD2', 'cK1', 'god', 'mes'],
                     aggfunc={'sum': sum},)
tb5.reset_index(inplace=True)
tb5 = tb5.fillna(0)

# Выгрузка сделанных анализов в РД
sht = xw.Book('TestPrAn.xlsx').sheets['аморт']
sht.range('A4').options(index=False, header=True).value = tb2

sht = xw.Book('TestPrAn.xlsx').sheets['закуп']
sht.range('A2').options(index=False, header=False).value = tb4

sht = xw.Book('TestPrAn.xlsx').sheets['выр']
sht.range('A2').options(index=False, header=False).value = tb5


# print(tb4)
# print(dp.head(5))
# print(dp.columns)
# print(tb.shape)
# print(tb.dtypes)
