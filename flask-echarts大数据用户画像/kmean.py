import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from pylab import *
mpl.rcParams['font.sans-serif'] = ['SimHei']
data='result\\总和.xlsx'
data_c1=pd.read_excel(data)
data_c1=data_c1[['违约次数','办理次数','违约金额','在网时长（天）','总用电量']]

def distEclud(vecA, vecB):
    """
    计算两个向量的欧式距离的平方，并返回
    """
    return np.sum(np.power(vecA - vecB, 2))


def test_Kmeans_nclusters(data_train):
    """
    计算不同的k值时，SSE的大小变化
    """
    data_train = data_train.values
    nums = range(2, 10)
    SSE = []
    for num in nums:
        sse = 0
        kmodel = KMeans(n_clusters=num, n_jobs=4)
        kmodel.fit(data_train)
        # 簇中心
        cluster_ceter_list = kmodel.cluster_centers_
        # 个样本属于的簇序号列表
        cluster_list = kmodel.labels_.tolist()
        for index in range(len(data_c1)):
            cluster_num = cluster_list[index]
            sse += distEclud(data_train[index, :], cluster_ceter_list[cluster_num])
        SSE.append(sse)
    return nums, SSE

# nums, SSE = test_Kmeans_nclusters(data_c1)
# plt.style.use('ggplot')
# ## 绘图观测SSE与簇个数的关系
# fig=plt.figure(figsize=(10, 8))
# ax=fig.add_subplot(1,1,1)
# ax.plot(nums,SSE,marker="+")
# ax.set_xlabel("n_clusters", fontsize=18)
# ax.set_ylabel("SSE", fontsize=18)
# fig.suptitle("KMeans", fontsize=20)
# plt.show()

kmodel = KMeans(n_clusters=5, n_jobs=4)
kmodel.fit(data_c1)
# 简单打印结果
r1 = pd.Series(kmodel.labels_).value_counts()  # 统计各个类别的数目
r2 = pd.DataFrame(kmodel.cluster_centers_)  # 找出聚类中心
# 所有簇中心坐标值中最大值和最小值
max = r2.values.max()
min = r2.values.min()
r = pd.concat([r2, r1], axis=1)  # 横向连接（0是纵向），得到聚类中心对应的类别下的数目
r.columns = list(data_c1.columns) + [u'类别数目']  # 重命名表头

# 绘图
fig = plt.figure(figsize=(10, 8))
ax = fig.add_subplot(111, polar=True)
center_num = r.values
feature = ['违约次数','办理次数','违约金额','在网时长（天）','总用电量']
N = len(feature)
for i, v in enumerate(center_num):
    # 设置雷达图的角度，用于平分切开一个圆面
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False)
    # 为了使雷达图一圈封闭起来，需要下面的步骤
    center = np.concatenate((v[:-1], [v[0]]))
    angles = np.concatenate((angles, [angles[0]]))
    # 绘制折线图
    ax.plot(angles, center, 'o-', linewidth=2, label="第%d簇人群,%d人" % (i + 1, v[-1]))
    # 填充颜色
    ax.fill(angles, center, alpha=0.25)
    # 添加每个特征的标签
    ax.set_thetagrids(angles * 180 / np.pi, feature, fontsize=15)
    # 设置雷达图的范围
    ax.set_ylim(min - 0.1, max + 0.1)
    # 添加标题
    plt.title('客户群特征分析图', fontsize=20)
    # 添加网格线
    ax.grid(True)
    # 设置图例
    plt.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0), ncol=1, fancybox=True, shadow=True)

# 显示图形
plt.show()
