import pandas as pd
from scipy.interpolate import lagrange
Power_charge='D:/sql/demo/data/电量电费.xlsx'
outputfile='D:/sql/demo/data/电量电费1.xlsx'

data=pd.read_excel(Power_charge,header=None)
data.info()   #检测每一列的缺失值情况
#自定义列向量插值函数
#s为列向量，n为被插值的位置，k为取前后的数据个数，默认为5

def ployinterp_column(s,n,k=5):
    y=s[list(range(n-k,n))+list(range(n+1,n+1+k))] #取数
    y=y[y.notnull()]#剔除空值
    return lagrange(y.index,list(y))(n)    #插值并取回插值结果

#逐个元素判断是否需要插值
for i in data.columns:
    for j in range(len(data)):
        if (data[i].isnull())[j]:
            data[i][j]=ployinterp_column(data[i],j)
data.to_excel(outputfile,header=None,index=False)