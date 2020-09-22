import pandas as pd
import pickle
import xlwings as xw
pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# выбрал и сохранил выбранные
# dp = pd.read_csv('PrTest.csv')
# dp = dp.fillna(0)
# dp = dp.loc[dp['muk'] == True]
# dp = dp.loc[dp['vn'] == 0]
# dp.to_csv('выбранные0.csv', index=False)

# dp = pd.read_csv('выбранные.csv')

# # добавляем количества по проводкам и меньше 12 раз в год
# with open('redko.pickle', 'rb') as f: redko = pickle.load(f)
# dp["kolPr"] = dp["DK"].map(redko)
# dp['redko'] = (dp['kolPr'] < 12).astype(int)

# dp.to_csv('выбранные.csv', index=False)

# добавляем круглость, одинаковые суммы, отрицательные суммы, выручка
# dp['tzel'] = (((dp['sum'].astype('str')).str.split('.').str.get(1)) == '0') \
#     .astype(int)
# dp['kr1'] = ((((dp['sum']/1000).astype('str')).str.split('.').str.get(1))).str[0:3]
# dp['kr'] = ((dp['kr1'] == '0')|(dp['kr1'] == '999')|(dp['kr1'] == '888')|(dp['kr1'] == '777')
#             |(dp['kr1'] == '666')|(dp['kr1'] == '555')|(dp['kr1'] == '444')|(dp['kr1'] == '333')
#             |(dp['kr1'] == '222')|(dp['kr1'] == '111')).astype('int')
# dp['krr'] = dp['kr']*dp['tzel']
# dp['dubl'] = dp['sum'].abs().duplicated(False).astype(int)
# dp['otritz'] = (dp['sum'] < 0).astype(int)
# dp['vir'] = (dp['DK'] == '62;90').astype(int)
# dp['abs'] = dp['sum'].abs()

# добавляем выходные и ночь
# with open('vixDen.pickle', 'rb') as f: vixDen = pickle.load(f)
# dp['DT'] = dp.god.astype(str)+'#'+dp.mes.astype(str)+'#'+dp.denb.astype(str)
# dp['vixodnoy'] = dp['DT'].map(vixDen)
# dp['notch'] = ((dp['time'] >= 1)&(dp['time'] <= 5)).astype(int)
# dp = dp.fillna(0)
#
# dp.to_csv('выбранные.csv', index=False)

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = dp

# переносим выбранные в РД
# dp = dp[['data', 'dok', 'org',
#          'schD', 'cD1', 'cD2', 'cD3',
#          'schK', 'cK1', 'cK2', 'cK3',
#          'sum', 'abs', 'text',
#          'schD2', 'schK2', 'DK',
#          'krr', 'vixodnoy', 'notch',
#          'dubl', 'otritz', 'vir', 'redko']]
# sht = xw.Book('TestPr.xlsx').sheets['выбранные']
# sht.range('A9').options(index=False, header=False).value = dp

