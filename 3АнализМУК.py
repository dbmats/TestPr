import numpy as np
import pandas as pd
import pickle

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 15)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# dp = pd.read_csv('PrTest.csv')
# dp = dp.fillna(0)
# dp = dp[['DK', 'sum', 'vn']]
# dp['abs'] = dp['sum'].abs()
# dp.to_csv('группы8.csv', index=False)

dp = pd.read_csv('группы.csv')
dp = dp.loc[dp['vn'] == 0]
k1 = 100
k2 = 10000
dp.loc[(dp['abs'] > k2), 'mln'] = 2
dp.loc[dp['abs'] <= k1, 'mln'] = 1

dp = dp.fillna(0)
tb = pd.pivot_table(dp, values=['abs', 'sum','DK'],
                    index=['mln'],
                    aggfunc={'abs': sum, 'sum': sum, 'DK': 'count'}
                    , margins=True
                    )
tb.reset_index(inplace=True)
tb = tb.fillna(0)

print(tb)

# tb.plot(x='mln', y='data')
# plt.grid(True)
# plt.show()

# tb.to_excel("output.xlsx")
# print(df[['abs']])
#
# print(dp.head(5))
# print(dp.columns)
# print(tb1.shape)
# print(tb.dtypes)
# dp1 = dp['rdate'].unique()
# dp1=pd.Series(dp1)
