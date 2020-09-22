import pandas as pd
import pickle
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 25)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# дополнение совокупности столбцами и формирование csv
xlapp = xw.apps.active
rng = xlapp.selection
dp = pd.DataFrame(rng.options(pd.DataFrame, header=1, index=False).value)

# dp = pd.read_csv('PrTest.csv')

# если дата в формате (день.мес.год час:мин:сек)
dp['time'] = (((dp['data'].astype('str')).str.split(' ').str.get(1)).str.split(':').str.get(0)).astype('int64')
dp['god'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(2)
dp['mes'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(1)
dp['denb'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(0)

# если дата в формате (год.мес.день)
# dp['god'] = (dp['data'].astype('str')).str.split('-').str.get(0)
# dp['mes'] = (dp['data'].astype('str')).str.split('-').str.get(1)
# dp['denb'] = (dp['data'].astype('str')).str.split('-').str.get(2)
# dp['time'] = 0

dp['schD2'] = (dp['schD'].astype('str')).str[0:2]
dp['schK2'] = (dp['schK'].astype('str')).str[0:2]
dp['DK'] = dp.schD2.astype(str).str.cat(dp.schK2.astype(str), sep=';')

with open('prVn.pickle', 'rb') as f: prVn = pickle.load(f)
dp["vn"] = dp["DK"].map(prVn)

dp['xxx'] = (dp['sum'].abs() * 1000000).astype('str')
dp['xx'] = dp.xxx.str[0:2]
dp['x1'] = dp.xxx.str[0:1]
dp['x2'] = dp.xx.str[1:2]

dp.to_csv('PrTest.csv', index=False)

# создал словарь с количествами по проводкам
tb = pd.pivot_table(dp, values=['mes'],
                    index=['DK'],
                    aggfunc={'mes': 'count'})
tb.reset_index(inplace=True)
tb = tb.fillna(0)
redko = dict(sorted(tb.values.tolist()))
with open('redko.pickle', 'wb') as f: pickle.dump(redko, f)


# dp.to_excel("output.xlsx")
print(dp)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(dp[['schD', 'schD2']])
# print(dp.dtypes)
