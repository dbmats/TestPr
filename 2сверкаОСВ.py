import pandas as pd
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# формирование Дт Кт обороты для сверки с ОСВ

dp = pd.read_csv('PrTest.csv')
dp = dp.fillna(0)

tb = pd.pivot_table(dp, values=['sum'],
                    index=['schD2', 'god'],
                    aggfunc={'sum': sum})
tb['sumK'] = 0.0
tb.columns = ['sumD', 'sumK']

tb1 = pd.pivot_table(dp, values=['sum'],
                     index=['schK2', 'god'],
                     aggfunc={'sum': sum})
tb1['sumD'] = 0.0
tb1 = tb1[['sumD', 'sum']]
tb1.columns = ['sumD', 'sumK']

fr = [tb, tb1]
pp = pd.concat(fr)
pp.reset_index(inplace=True)

pp1 = pd.pivot_table(pp, values=['sumD', 'sumK'],
                     columns=['god'],
                     index=['schD2'],
                     aggfunc={'sumD': sum, 'sumK': sum})
pp1.reset_index(inplace=True)
pp1 = pp1.fillna(0)

sht = xw.Book('TestPr.xlsx').sheets['сверка']
sht.range('A10').options(index=False, header=True).value = pp1

# pp1.to_csv('1.csv')
# tb.to_excel("output.xlsx")
print(pp1)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(pp.dtypes)
