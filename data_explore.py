import pandas as pd
import numpy as np
from datetime import datetime
from pyecharts.charts import Bar,Page,EffectScatter,WordCloud,Line
from pyecharts import options as opts
def dispose():
    userinfo = 'data\\新津用户档案.xlsx'
    Power_charge = 'data\\新津电量电费.xlsx'
    #数据处理
    data_c1=pd.read_excel(userinfo) #data_c1.info()  #检查缺失值情况
    data_c2=pd.read_excel(Power_charge) #data_c2.info()  #检查缺失值情况
    def data_cleaning(df):
        cols = df.columns
        for col in cols:
            if df[col].dtype ==  'object':
                df[col].fillna('缺失数据', inplace = True)
            else:
                df[col].fillna(0, inplace = True)
        return(df)
    data_c3=data_cleaning(data_c2)
    #清除有记录缺失的数据
    for a in data_c3['用户编号']:
        if len(data_c3[data_c3['用户编号']==a])!=9:
            data_c3.drop(data_c3[data_c3['用户编号']==a].index[:], inplace=True)
    for b in data_c3['用户名称']:
        if len(data_c3[data_c3['用户名称'] == b]) != 9:
            data_c3.drop(data_c3[data_c3['用户名称'] == b].index[:], inplace=True)
    data_c4=pd.merge(data_c1,data_c3,how='inner',on=['用户编号','用户名称'])
    #数据提取
    data_c5=data_c4[['用户编号','用户名称','电费年月','总电量','总电费',
                     '用电类别','行业分类','合同容量','峰电量','平电量','谷电量']]
    return data_c5
def total():
    # 计算总用电量、总电费
    total_power=dispose()
    df_q1 = total_power.groupby('用户编号')[['总电量', '总电费']].sum()
    df_q2 = total_power.groupby('用户编号')[['总电量', '总电费']].mean()
    df_q3 = total_power.groupby('用户编号')[['总电量', '总电费']].agg(max)
    df_q1.columns = ['总用电量', '总电费']
    df_q2.columns = ['月均用电量', '月均电费']
    df_q3.columns = ['月最大用电量', '月最大电费']
    qdata1 = pd.merge(df_q1,df_q2,how='inner', on=['用户编号'])
    qdata1 = pd.merge(qdata1, df_q3, how='inner', on=['用户编号'])
    qdata2 = total_power[['用户编号', '用户名称', '用电类别', '行业分类','合同容量']]
    qdata3 = pd.merge(qdata2, qdata1, how='inner', on=['用户编号'])
    qdata3 = qdata3.drop_duplicates()
    qdata3['合同容量']=qdata3['合同容量']*1000
    # qdata3.to_excel('result\\电量电费总和.xlsx', na_rep=None)
    return qdata3
def increase_power(x):
    # 计算电量增长率
    total_power = dispose()
    c1 = total_power[total_power['电费年月'] == x][['用户名称','总电量']]
    c1.columns = ['用户名称',f'{x}月电量']
    c2 = total_power[total_power['电费年月'] == x + 1][['用户名称','总电量']]
    c2.columns = ['用户名称',f'{x + 1}月电量']
    result=pd.merge(c1,c2,how='inner',on='用户名称')
    result[f'{x}月电量增长率']=(result[f'{x + 1}月电量']-result[f'{x}月电量'])/result[f'{x}月电量']
    res=result[['用户名称',f'{x}月电量增长率']]
    return res
def increase_price(x):
    total_power = dispose()
    # 计算电量电费增长率
    c1 = total_power[total_power['电费年月'] == x][['用户名称', '总电费']]
    c1.columns = ['用户名称', f'{x}月电费']
    c2 = total_power[total_power['电费年月'] == x + 1][['用户名称', '总电费']]
    c2.columns = ['用户名称', f'{x + 1}月电费']
    result=pd.merge(c1,c2,how='inner',on='用户名称')
    result[f'{x}月电费增长率'] = (result[f'{x + 1}月电费'] - result[f'{x}月电费']) / result[f'{x}月电费']
    res=result[['用户名称',f'{x}月电费增长率']]
    return res
def avg_power():
    # 计算平均电量增长率
    total_power = dispose()
    datatime = total_power['电费年月'].drop_duplicates()
    data = []
    for d in datatime:
        data.append(d)
    user=total_power[total_power['电费年月'] == datatime[0]][['用户名称','用电类别','行业分类']]
    usernumber=total_power[total_power['电费年月'] == datatime[0]][['用户编号','用户名称']]
    for y in range(len(data)-1):
        y1=increase_power(datatime[y])
        user=pd.merge(user,y1,how='inner',on='用户名称')
    user['电量平均增长率']=round(user.mean(1),2)
    avgpower=pd.merge(usernumber,user[['用户名称','用电类别','行业分类','电量平均增长率']],how='inner',on='用户名称')
    return avgpower
def avg_price():
    # 计算平均电费增长率
    total_power = dispose()
    datatime = total_power['电费年月'].drop_duplicates()
    data = []
    for d in datatime:
        data.append(d)
    user = total_power[total_power['电费年月'] == datatime[0]][['用户名称', '用电类别', '行业分类']]
    usernumber = total_power[total_power['电费年月'] == datatime[0]][['用户编号', '用户名称']]
    for y in range(len(data) - 1):
        y1 = increase_price(datatime[y])
        user = pd.merge(user, y1, how='inner', on='用户名称')
    user['电费平均增长率'] = round(user.mean(1),2)
    avgprice = pd.merge(usernumber, user[['用户名称', '用电类别', '行业分类', '电费平均增长率']], how='inner', on='用户名称')
    return avgprice
    # avgpower_price = pd.merge(avg_power(), avg_price(), how='inner', on=['用户编号', '用户名称', '用电类别', '行业分类'])
    # avgpower_price.to_excel('result\\电量电费平均增长率.xlsx', na_rep=None)
def Default_power():
    """计算用户违约办理次数，用户违约金额"""
    Default_power = 'data\\违约窃电.xlsx'
    Default = pd.read_excel(Default_power)
    Default = Default[['用户号', '用户名', '违约年份', '违窃电性质', '违约使用电费', '用电类别', '行业类别']]

    power = Default.groupby('用户号')[['违窃电性质']].count()
    power.columns = ['违约次数']
    price = Default.groupby('用户号')[['违约使用电费']].sum()
    price.columns = ['违约金额']

    power_price = pd.merge(power, price, how='inner', on='用户号')
    Default = pd.merge(Default, power_price, how='inner', on='用户号')
    Default = Default[['用户号', '用户名', '违约年份', '违约次数', '违约金额', '用电类别', '行业类别']]
    Default = Default.drop_duplicates()
    df_q1 = Default.groupby('用电类别')[['用户号']].count()
    Default1=Default[['用电类别']]
    Default1=Default1.drop_duplicates()

    Default1_df_q1=pd.merge(df_q1,Default1,how='inner',on=['用电类别'])
    Default1_df_q1=Default1_df_q1.head(5)
    bar = Bar(init_opts=opts.InitOpts())
    bar.add_xaxis(Default1_df_q1['用电类别'].values.tolist())
    bar.add_yaxis('2019年', Default1_df_q1['用户号'].values.tolist(), color="blue")
    bar.set_global_opts(title_opts=opts.TitleOpts(title="各行业用电类别窃漏电情况",
                                                  subtitle="窃漏电用户数"))
    bar.render()
Default_power()
    # return Default
def cut_po():
    """计算各个线路的停电次数"""
    userinfo = 'data\\用户档案.xlsx'
    user = pd.read_excel(userinfo)
    power_cut = 'data\\停电信息.xlsx'
    cut = pd.read_excel(power_cut)
    cut_power = cut[['所属供电单位', '工单编号', '停电编号', '停电类型', '停电开始时间', '停电结束时间', '停电范围', '停电区域']]
    cut_power.columns = ['区县公司', '工单编号', '停电编号', '停电类型', '停电开始时间', '停电结束时间', '停电范围', '线路名称']
    cut_power['停电时间间隔'] = pd.to_datetime(cut_power['停电结束时间']) - pd.to_datetime(cut_power['停电开始时间'])
    cut_power1=cut_power[['区县公司', '工单编号', '停电编号', '停电时间间隔','停电范围', '线路名称']]
    cut_power=cut_power[['区县公司', '工单编号', '停电编号', '停电类型','停电时间间隔','停电范围', '线路名称']]
    cut_power_count = cut_power.groupby('线路名称')['停电类型'].count()
    user_cut = pd.merge(cut_power1, cut_power_count, how='inner', on=['线路名称'])
    user=pd.merge(user,user_cut,how='inner', on=['线路名称'])
    user.to_excel('result\\每个用户的停电次数.xlsx', na_rep=None)
    user=user[['用户编号','用户名称','用电类别','行业分类','线路名称','停电类型','停电时间间隔']]
    user.columns = ['用户编号','用户名称','用电类别','行业分类','线路名称','停电次数','停电时间间隔']
    user.to_excel('result\\每个用户的停电次数.xlsx', na_rep=None)
    return user
def transformertime():
    """计算变压器投运时间间隔"""
    transformer = 'data\\变压器数据.xlsx'
    trans = pd.read_excel(transformer)
    trans = trans[['户号', '户名', '线路名称', '投运时间']]
    trans["投运年限"] = trans["投运时间"].apply(
        lambda x: "" if pd.isnull(x) or len(str(x)) <= 10 else round(((datetime.today() - x).days) / 365, 0))
    trans = trans[['户号', '户名', '线路名称', '投运年限']]
    trans.to_excel('result\\电变压器投运年限.xlsx', na_rep=None)
def bus_handing():
    """计算业务办理次数，业务办理时间间隔"""
    business_handling = 'data\\业务办理.xlsx'
    business = pd.read_excel(business_handling)
    business = business[['用户编号', '用户名称', '业务类型', '申请时间',
                         '完成时间', '业务时长（天）', '用电类别', '行业分类']]
    business['业务办理时间间隔'] = pd.to_datetime(business['完成时间']) - pd.to_datetime(business['申请时间'])
    business = business[['用户编号', '用户名称', '业务类型', '业务办理时间间隔',
                         '业务时长（天）', '用电类别', '行业分类']]
    datatime = business.groupby('用户编号')[['业务办理时间间隔']].sum()
    handling = business.groupby('用户编号')[['业务类型']].count()
    handling_datatime=pd.merge(handling, datatime, how='inner', on='用户编号')
    handling_datatime.columns = ['办理次数','业务办理时间间隔']
    handling_datatime.to_excel('result\\每个用户办理次数.xlsx', na_rep=None)

def grid_vertical():
    data = [('老客户',10000),('无违约金额',6181),('无投诉',4386),('业务办理较少',4055),('合同容量极高',2467),('用户规模中等',2244),
            ('变压器投运5年以下',1898),('有备用电源',1484),('锋值用电比例低',1112),('谷值用电比例',965),('用户规模中速增长',847),
            ('平值用电比例低',582),('累计用电量极高',555),('客户月均用电量极大',550),('基本电费极高',462),('无意见建议',366)]
    Word = WordCloud()
    Word.add('', data,word_size_range=[20, 100])
    Word.set_global_opts(title_opts=opts.TitleOpts(title="客户命中标签"))
    power = dispose()
    power = power[power['用户编号'] == 7110207071]
    bar = Bar(init_opts=opts.InitOpts())
    bar.add_xaxis(power['电费年月'].values.tolist())
    bar.add_yaxis('2019年', power['总电量'].values.tolist(),color="blue")
    bar.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年{power['用户名称'].values.tolist()[0]}用户电费",
                                                  subtitle="用电量（万千瓦）"))
    bar1 = Bar(init_opts=opts.InitOpts())
    bar1.add_xaxis(power[power['用户编号'] == 7110207071]['电费年月'].values.tolist())
    bar1.add_yaxis('锋值用电量', power['峰电量'].values.tolist(), color='blue')
    bar1.add_yaxis('平值用电量', power['平电量'].values.tolist(), color='green',gap='0')
    bar1.add_yaxis('古值用电量', power['谷电量'].values.tolist(), color='red',gap='0')
    bar1.set_global_opts(title_opts=opts.TitleOpts(title="2019年该用户用电情况",
                                                  subtitle="用电量（万千瓦）"))
    # avgpower_price = pd.merge(avg_power(), avg_price(), how='inner', on=['用户编号', '用户名称', '用电类别', '行业分类'])
    power_price=pd.read_excel('result\\电量电费平均增长率.xlsx')
    avgpower_price=power_price.loc[power_price['行业分类'].str.contains('（3）机动车、电子产品和日用产品修理业')]
    bar2 = Bar(init_opts=opts.InitOpts())
    bar2.add_xaxis(avgpower_price['用户名称'].values.tolist())
    bar2.add_yaxis('电费量增长率', avgpower_price['电费平均增长率'].values.tolist(), color='blue')
    bar2.add_yaxis('电量增长率', avgpower_price['电量平均增长率'].values.tolist(), color='green',gap='0')
    bar2.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年同行业电费电量增长率图",
                                                   subtitle="百分比"),
                         xaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True),
                                                  axislabel_opts=opts.LabelOpts(rotate=30, font_size=12)),
                         yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True)))
    totalpower = total()
    totalpower = totalpower.loc[totalpower['行业分类'].str.contains('（3）机动车、电子产品和日用产品修理业')]
    eff = (
        EffectScatter(init_opts=opts.InitOpts())
        .set_global_opts(
            title_opts=opts.TitleOpts(title="2019年同行业用电情况",subtitle="总电量  万千瓦"),
            xaxis_opts=opts.AxisOpts(type_="value",
                                     name="合同容量 KVA",
                                     name_gap=30,
                                     splitline_opts=opts.SplitLineOpts(is_show=True)),
            yaxis_opts=opts.AxisOpts(name="总用电量",
                                     name_gap=80,
                                     splitline_opts=opts.SplitLineOpts(is_show=True)),
            legend_opts=opts.LegendOpts(
                orient="vertical",# 设置图例的布局方式：垂直布局
                pos_right="0%",		# 设置图例距离左侧的距离，%
                pos_top="0%"
            ),
        )
            .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    )
    for name, group in totalpower[["用户名称", "合同容量", "总用电量"]].groupby("用户名称"):
        eff = (
            eff
                .add_xaxis(group["合同容量"].tolist())
                .add_yaxis(name,
                           group["总用电量"].tolist())
        )

    page = Page(layout=Page.SimplePageLayout)
    page.add(Word,bar,bar1,bar2,eff)
    page.render("大数据用户画像可视化系统展示.html")
    return page

def page_vertical():
    total='result\\电量电费总和.xlsx'
    power_total=pd.read_excel(total)
    power_total = power_total.sort_values(by='月均用电量', ascending=False)
    power_total['月均用电量'] = power_total['月均用电量'].apply(np.round) / 10
    df=power_total.head(20)
    bar = Bar()
    bar.add_xaxis(df['用户名称'].tolist())
    bar.add_yaxis('', df['月均用电量'].tolist(), color='blue')
    bar.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个用户月均用电量",pos_left='50%'),
                        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                 splitline_opts=opts.SplitLineOpts(is_show=True),name='用户名称'),
                        yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True),name='月均用电量 （M）'))

    avgpower_total = power_total.groupby('用电类别')[['月均电费']].sum()
    avgpower_total['月均电费'] = avgpower_total['月均电费'].apply(np.round) / 10
    type=power_total[['用电类别']]
    avgpower_total=pd.merge(avgpower_total,type,how='inner',on='用电类别')
    avgpower_total = avgpower_total.drop_duplicates()
    avgpower_total = avgpower_total.sort_values(by='月均电费', ascending=False)
    bar1 = Bar()
    bar1.add_xaxis(avgpower_total['用电类别'].tolist())
    bar1.add_yaxis('', avgpower_total['月均电费'].tolist(), color='blue')
    bar1.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个行业月均电费",pos_left='50%'),
                        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                 splitline_opts=opts.SplitLineOpts(is_show=True), name='用电类别'),
                        yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='月均电费'))

    increase = 'result\\电量电费平均增长率.xlsx'
    increase_total = pd.read_excel(increase)
    df = increase_total.head(20)
    bar2 = Bar()
    bar2.add_xaxis(df['用户名称'].tolist())
    bar2.add_yaxis('', df['电量平均增长率'].tolist(), color='blue')
    bar2.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个用户月均用电量增长率", pos_left='50%'),
                        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                 splitline_opts=opts.SplitLineOpts(is_show=True), name='用户名称'),
                        yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='电量平均增长率 （M）'))

    avgincrease_total = increase_total.groupby('行业分类')[['电量平均增长率']].sum()
    type = increase_total[['行业分类']]
    avgincrease_total = pd.merge(avgincrease_total, type, how='inner', on='行业分类')
    avgincrease_total = avgincrease_total.drop_duplicates()
    bar3 = Bar()
    bar3.add_xaxis(avgincrease_total['行业分类'].tolist())
    bar3.add_yaxis('', avgincrease_total['电量平均增长率'].tolist(), color='blue')
    bar3.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个行业电量平均增长率", pos_left='50%'),
                         xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                  splitline_opts=opts.SplitLineOpts(is_show=True), name='行业分类'),
                         yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='电量平均增长率'))
    bar3.set_series_opts(label_opts=opts.LabelOpts(is_show=False))

    Default=Default_power()
    Defaulttype=Default.groupby('行业类别')[['违约次数']].count()
    Defaulttype.columns = ['行业违约次数']
    Defaulttype1 = Default.groupby('行业类别')[['违约金额']].sum()
    Defaulttype1.columns = ['行业违约金额']
    Default = Default[['行业类别']]
    Defaultcount=pd.merge(Default,Defaulttype,how='inner',on='行业类别')
    Defaultcount = Defaultcount.sort_values(by='行业违约次数', ascending=False)
    Defaultcount = Defaultcount.drop_duplicates()
    Defaultcount=Defaultcount.head(20)
    Defaultprice = pd.merge(Default, Defaulttype1, how='inner', on='行业类别')
    Defaultprice = Defaultprice.sort_values(by='行业违约金额', ascending=False)
    Defaultprice = Defaultprice.drop_duplicates()
    Defaultprice = Defaultprice.head(20)
    bar4 = Bar()
    bar4.add_xaxis(Defaultcount['行业类别'].tolist())
    bar4.add_yaxis('', Defaultcount['行业违约次数'].tolist(), color='blue')
    bar4.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个行业内用户的违约次数",pos_left='50%'),
                        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                 splitline_opts=opts.SplitLineOpts(is_show=True),name='行业类别'),
                        yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True),name='行业违约次数 '))
    bar5 = Bar()
    bar5.add_xaxis(Defaultprice['行业类别'].tolist())
    bar5.add_yaxis('', Defaultprice['行业违约金额'].tolist(), color='blue')
    bar5.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年各个行业内用户的违约金额", pos_left='50%'),
                         xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                  splitline_opts=opts.SplitLineOpts(is_show=True), name='行业类别'),
                         yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='行业违约金额 '))
    bar5.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    cut=cut_po()
    cut = cut.sort_values(by='停电次数', ascending=False)
    cut = cut.drop_duplicates()
    cut = cut.head(20)
    bar6 = Bar()
    bar6.add_xaxis(cut['用户名称'].tolist())
    bar6.add_yaxis('', cut['停电次数'].tolist(), color='blue')
    bar6.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年用户的停电次数", pos_left='50%'),
                         xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                  splitline_opts=opts.SplitLineOpts(is_show=True), name='用户名称'),
                         yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='停电次数 '))
    userinfo = 'data\\新津用户档案.xlsx'
    user=pd.read_excel(userinfo)
    user1=user[['用户名称']]
    usergroup = user.groupby('用户名称')[['电源数目']].count()
    usergroup=pd.merge(usergroup,user1,how='inner',on='用户名称')
    usergroup = usergroup.sort_values(by='电源数目', ascending=False)
    usergroup = usergroup.drop_duplicates()
    usergroup = usergroup.head(20)
    bar7 = Bar()
    bar7.add_xaxis(usergroup['用户名称'].tolist())
    bar7.add_yaxis('', usergroup['电源数目'].tolist(), color='blue')
    bar7.set_global_opts(title_opts=opts.TitleOpts(title=f"2019年用户的电源数目", pos_left='50%'),
                         xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=30, font_size=12),
                                                  splitline_opts=opts.SplitLineOpts(is_show=True), name='用户名称'),
                         yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True), name='电源数目 '))
    page = Page(layout=Page.SimplePageLayout)
    page.add(bar, bar1,bar2,bar3,bar4,bar5,bar6,bar7)
    page.render()