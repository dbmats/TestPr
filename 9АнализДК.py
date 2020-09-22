import pandas as pd
import xlwings as xw
pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# формирование Дт Кт оборотов
dp = pd.read_csv('PrTest.csv')
dp = dp.fillna(0)

# ГК
# tb = pd.pivot_table(dp, values=['sum'],
#                     index=['schD2', 'schK2', 'god', 'mes'],
#                     aggfunc={'sum': sum})
# tb.reset_index(inplace=True)
# tb['sumK'] = 0.0
# tb.columns = ['sch', 'ksch', 'god', 'mes', 'sumD', 'sumK']
#
# tb1 = pd.pivot_table(dp, values=['sum'],
#                      index=['schK2', 'schD2', 'god', 'mes'],
#                      aggfunc={'sum': sum})
# tb1.reset_index(inplace=True)
# tb1['sumD'] = 0.0
# tb1.columns = ['sch', 'ksch', 'god', 'mes', 'sumK', 'sumD']
# tb1 = tb1[['sch', 'ksch', 'god', 'mes', 'sumD', 'sumK']]
#
# fr = [tb, tb1]
# pp = pd.concat(fr)
#
# pp1 = pd.pivot_table(pp, values=['sumD', 'sumK'],
#                      index=['sch', 'ksch'],
#                      columns=['god', 'mes'],
#                      aggfunc={'sumD': sum, 'sumK': sum},
#                      margins= True)
# pp1.reset_index(inplace=True)
# pp1 = pp1.fillna(0)
#
# sht = xw.Book('TestPrAn.xlsx').sheets['ГКдк']
# sht.range('A8').options(index=False, header=True).value = pp1

# сч.91
# dp91 = dp.loc[dp['schD2'] == 91]
# dp91 = dp91.loc[~dp91['schK2'].isin([91, 99])]
# tb = pd.pivot_table(dp91, values=['sum'],
#                     index=['cD1','schK2', 'god'],
#                     aggfunc={'sum': sum})
# tb.reset_index(inplace=True)
# tb['sumK'] = 0.0
# tb.columns = ['cubk', 'Ksch', 'god', 'sumD', 'sumK']
#
# dp91 = dp.loc[dp['schK2'] == 91]
# dp91 = dp91.loc[~dp91['schD2'].isin([91, 99])]
# tb1 = pd.pivot_table(dp91, values=['sum'],
#                      index=['cK1', 'schD2', 'god'],
#                      aggfunc={'sum': sum})
# tb1.reset_index(inplace=True)
# tb1['sumD'] = 0.0
# tb1.columns = ['cubk', 'Ksch', 'god', 'sumK', 'sumD']
# tb1 = tb1[['cubk', 'Ksch', 'god', 'sumD', 'sumK']]
#
# fr = [tb, tb1]
# pp = pd.concat(fr)
# pp.reset_index(inplace=True)
#
# pp3 = pd.pivot_table(pp, values=['sumD', 'sumK'],
#                      index=['cubk'],
#                      columns=['god', 'Ksch'],
#                      aggfunc={'sumD': sum, 'sumK': sum},
#                      margins=True)
# pp3.reset_index(inplace=True)
# pp3 = pp3.fillna(0)
#
# sht = xw.Book('TestPrAn.xlsx').sheets['сч.91']
# sht.range('A8').options(index=False, header=True).value = pp3

# сч.90
# dp90 = dp.loc[dp['schD2'] == 90]
# dp90 = dp90.loc[~dp90['schK2'].isin([90, 99])]
# tb = pd.pivot_table(dp90, values=['sum'],
#                     index=['cD1', 'schD', 'schK2', 'god', 'mes'],
#                     aggfunc={'sum': sum})
# tb.reset_index(inplace=True)
# tb['sumK'] = 0.0
# tb.columns = ['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']
#
# dp90 = dp.loc[dp['schK2'] == 90]
# dp90 = dp90.loc[~dp90['schD2'].isin([90, 99])]
# tb1 = pd.pivot_table(dp90, values=['sum'],
#                      index=['cK1', 'schK', 'schD2', 'god', 'mes'],
#                      aggfunc={'sum': sum})
# tb1.reset_index(inplace=True)
# tb1['sumD'] = 0.0
# tb1.columns = ['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
# tb1 = tb1[['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]
#
# fr = [tb, tb1]
# pp = pd.concat(fr)
# pp.reset_index(inplace=False)
#
# pp1 = pd.pivot_table(pp, values=['sumD', 'sumK'],
#                      index=['cubk'],
#                      columns=['god', 'Ksch'],
#                      aggfunc={'sumD': sum, 'sumK': sum},
#                      margins=True)
# pp1.reset_index(inplace=True)
# pp1 = pp1.fillna(0)
#
# sht = xw.Book('TestPrAn.xlsx').sheets['сч.90']
# sht.range('A2').options(index=False, header=False).value = pp

# # сч.Контрагенты
# dpDK = dp.loc[dp['schD2'].isin([46, 58, 59, 60, 62, 63, 66, 67, 71, 73, 75, 76])]
# tb = pd.pivot_table(dpDK, values=['sum'],
#                     index=['cD1', 'cD2',
#                            'schD2', 'schK2', 'god', 'mes'],
#                     aggfunc={'sum': sum})
# tb.reset_index(inplace=True)
# tb['sumK'] = 0.0
# tb.columns = ['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']
#
# dpDK = dp.loc[dp['schK2'].isin([46, 58, 59, 60, 62, 63, 66, 67, 71, 73, 75, 76])]
# tb1 = pd.pivot_table(dpDK, values=['sum'],
#                      index=['cK1', 'cK2',
#                             'schK2', 'schD2', 'god', 'mes'],
#                      aggfunc={'sum': sum})
# tb1.reset_index(inplace=True)
# tb1['sumD'] = 0.0
# tb1.columns = ['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
# tb1 = tb1[['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]
#
# fr = [tb, tb1]
# pp = pd.concat(fr)
# pp.reset_index(inplace=False)
#
# sht = xw.Book('TestPrAn.xlsx').sheets['ДзКз']
# sht.range('A2').options(index=False, header=False).value = pp

# # сч.20е
dp20 = dp.loc[dp['schD2'].isin([20,
                                21, 23, 25, 26, 28, 29, 43, 44, 8, '08'
                                ])]
tb = pd.pivot_table(dp20, values=['sum'],
                    index=['cD1', 'cD2', 'cD3',
                           'schD2', 'schK2', 'god', 'mes'],
                    aggfunc={'sum': sum})
tb.reset_index(inplace=True)
tb['sumK'] = 0.0
tb.columns = ['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']

dp90 = dp.loc[dp['schK2'].isin([20,
                                21, 23, 25, 26, 28, 29, 43, 44, 8, '08'
                                ])]
tb1 = pd.pivot_table(dp90, values=['sum'],
                     index=['cK1', 'cK2', 'cK3',
                            'schK2', 'schD2', 'god', 'mes'],
                     aggfunc={'sum': sum})
tb1.reset_index(inplace=True)
tb1['sumD'] = 0.0
tb1.columns = ['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
tb1 = tb1[['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]

fr = [tb, tb1]
pp = pd.concat(fr)
pp.reset_index(inplace=False)

sht = xw.Book('TestPrAn.xlsx').sheets['08,20-44']
sht.range('A2').options(index=False, header=False).value = pp

# print(pp3)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(tb.dtypes)
