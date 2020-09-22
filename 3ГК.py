import pandas as pd
import pickle
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)


# формирование ГК
dp = pd.read_csv('PrTest.csv')
# dp = dp.loc[dp['god'] == 2019]
# dp['schD2'] = (dp['schD2'].astype('str')).str.split('.').str.get(0)
# dp['schK2'] = (dp['schK2'].astype('str')).str.split('.').str.get(0)
# tb = pd.pivot_table(dp, values=['sum', 'data'],
#                     index=['schD2', 'schK2'],
#                     columns=['muk'],
#                     aggfunc={'sum': sum, 'data': 'count'})
# tb.reset_index(inplace=True)
# tb = tb.fillna(0)
# tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')
#
# with open('prVn.pickle', 'rb') as f: prVn = pickle.load(f)
# with open('pr.pickle', 'rb') as f: pr = pickle.load(f)
# tb["vn"] = tb["DK"].map(prVn)
# tb["opis"] = tb["DK"].map(pr)
# tb = tb.fillna(0)
#
# xlapp = xw.apps.active
# rng = xlapp.selection
# rng.options(index=False).value = tb

# делаем помесячную ГК
tb = pd.pivot_table(dp, values=['sum', 'data'],
                    index=['schD2', 'schK2'],
                    columns=['god', 'mes'],
                    aggfunc={'sum': sum, 'data': 'count'}
                    , margins=True
                    )
tb.reset_index(inplace=True)
tb = tb.fillna(0)
tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = tb

# tb.to_excel("output.xlsx")
# print(tb)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(tb.dtypes)
