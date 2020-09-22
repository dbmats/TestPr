import pandas as pd
import xlwings as xw

xlapp = xw.apps.active
rng = xlapp.selection
dp = pd.DataFrame(rng.options(pd.DataFrame, header=1, index=False).value)

df = dp.fillna(0)


print(df.shape)
print(df.dtypes)
