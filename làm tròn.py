import numpy as np
import pandas as pd
from iteround import saferound

[1.0, 2.0, 4.0]
def round_retain_sum(x):
    x=np.array(x)
    x = x*100 # We want 2 decimal precision
    x=np.around(x).astype(int)
    N = np.around(np.sum(x)).astype(int)
    y = np.around(x).astype(np.uint64)
    M = np.sum(y)
    K = N - M
    z = y-x
    if K!=0:
        idx = np.argpartition(z,K)[:K]
        y[idx] += 1
    return y/100.
file_name="Task OS tháng 6.xlsx"
df = pd.read_excel(file_name,sheet_name="Sheet0")
new_groups=[]
groups = df.groupby('Mã story')


for idx, group in groups:
    tong =group["ULNL story"].max()
    group['ULNL task'] = group['ULNL task'] * 100
    group['Task round']=saferound(list(group['ULNL task']),places=0)
    group['Task round'] /= 100
    group['ULNL task'] = group['ULNL task'] / 100
    gap = round((tong - sum(group['Task round'])) /0.01)
    #print(group['ULNL task'].sum())
    if np.abs(gap)>0.01:
        print('check')
        for index,row in group.iterrows():
            if gap>0:
                if row["Task round"] < row["ULNL task"]:
                    row["Task round"]= (row["Task round"]*100+1)/100
                    print("OK")
                    break
            else:
                if row["Task round"] > row["ULNL task"]:
                    row["Task round"]= (row["Task round"]*100-1)/100
                    print("OK")
                    break
    new_groups.append(group)
df_new = pd.concat(new_groups)
df_new.sort_index().to_excel(file_name.split(".xlsx")[0]+"_new.xlsx",index=None)