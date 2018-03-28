"""
    版本：V1
    功能：
        1.考勤数据清洗、汇总, 写入excel中
        2.原始数据也写到excel不同页中
    优化：
        1.使用到原始数据的列仅有['姓名', '日期', '对应时段', '签到时间', '签退时间']
        2.V0版本中通过日期列产生星期几的列，在V1版本中通过日期列得知该日期是工作日还是休息日，还是节假日
        3.过滤离职员工代码优化
        4.按姓名排序（原始考勤数据未排序）
"""
__author__ = 'YangShiFu'
__date__ = '2017-10-17'


import numpy as np
import pandas as pd
import time
from pylab import *
import re
import xlrd

# 绘图设置中文格式
mpl.rcParams['font.sans-serif'] = ['SimHei']

# 读取原excel
raw_file = './doc/产品测试考勤-1月.xls'
# 写入新的excel
aim_file = './doc/产品测试考勤-1月-汇总.xls'


content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')
df_raw = pd.read_excel(content, engine='xlrd')

# 从excel获取需要的列数据
df = pd.DataFrame(data=df_raw, columns=['姓名', '日期', '对应时段', '签到时间', '签退时间'])
df = df.sort_values('姓名', axis=0)   # 按姓名排序

# 夜班员工
staff_night = ['杨纪华']

# 过滤掉离职员工
df_all = df.groupby('姓名').count()
df_out = df_all[df_all['签退时间'] == 0]
for name in df_out.index:
    df = df[df['姓名'] != name]

# 按姓名日期合并数据
df_am = df[df['对应时段'] == '上午']
df_pm = df[df['对应时段'] == '下午']
df_m = pd.merge(df_am, df_pm, how='outer', on=['姓名', '日期']).drop(columns=['对应时段_x', '对应时段_y'])

# ------------------------------------ 增加日期标志位 ----------------------------------------
# 工作日为值为0，周末值为1， 法定节值为2
def get_week(str_date):
    """
    判断日期是周末还是工作日
    :param str_date: 日期的字符串格式, 如"2018-03-05"
    :return: 工作日则返回0，周末则返回1
    """
    res = re.findall(r'\d+(.*?)\d+', str_date)[0]
    wday = time.strptime(str_date, '%Y'+ res + '%m' + res + '%d').tm_wday
    flag = 0 if wday < 5 else 1
    return flag

day_flag = df_m['日期'].apply(get_week).rename('是否工作日')
df_m = df_m.join(day_flag)

# # 去除周末
# df_m = df_m[df_m['是否工作日'] == 0]

# ------------------------------------ 上班打卡、中午打卡、下班打卡 ----------------------------------------
def get_noon(s):
    if pd.isnull(s['签退时间_x']):
        return s['签到时间_y']
    return s['签退时间_x']
s_noon = df_m.apply(get_noon, axis=1).rename('中午打卡')
df_m = df_m.join(s_noon).rename(columns={'签到时间_x':'上班打卡', '签退时间_y':'下班打卡'})
df_m = df_m.drop(['签退时间_x','签到时间_y'], axis=1)

# ------------------------------------ 是否异常 ----------------------------------------
def leave(s):
    if s['是否工作日'] != 0 or s['姓名'] in staff_night:
        return np.nan
    if pd.isnull(s['上班打卡']) or pd.isnull(s['中午打卡']) or pd.isnull(s['下班打卡']):
        return '异常'
    else:
        return np.nan
s_leave = df_m.apply(leave, axis=1).rename('是否异常')
df_normal = df_m.join(s_leave)


# ------------------------------------ 迟到时间、早退时间 ----------------------------------------
# 迟到、早退、是否旷工
def t2m(t):
    # """ 时分转换成分钟数 """
    h,m = t.strip().split(':')
    return int(h) * 60 + int(m)

def m2t(m):
    # 分钟数转换成时分
    h = m // 60
    m = m % 60
    return "%02d:%02d" % (h, m)

def late_time(s, col_name='上班打卡', expect_t='08:30'):
    # 如果是周末则值位NAN
    if s['是否工作日'] != 0 or s['姓名'] in staff_night:
        return np.nan
    # 计算上班迟到时间
    if pd.notnull(s[col_name]):
        if t2m(expect_t) >= t2m(s[col_name]):
            return np.nan
        diff_t = m2t(t2m(s[col_name]) - t2m(expect_t))
        return diff_t
    return np.nan

def early_time(s, col_name='下班打卡', expect_t='17:30'):
    # 如果是周末则值位NAN
    if s['是否工作日'] != 0 or s['姓名'] in staff_night:
        return np.nan
    # 计算下班早退时间
    if pd.notnull(s[col_name]):
        if t2m(expect_t) <= t2m(s[col_name]):
            return np.nan
        diff_t = m2t(t2m(expect_t) - t2m(s[col_name]))
        return diff_t
    return np.nan

def work_more(s):
    # 如果是周末则值位NAN
    if s['是否工作日'] != 0 or s['姓名'] in staff_night:
        return np.nan
    # 计算加班时间
    if pd.notnull(s['下班打卡']):
        more_t = t2m(s['下班打卡']) - t2m('17:30')
        if more_t >= 150:
            return m2t(more_t)
        else:
            return np.nan
    return np.nan

# 迟到时间或早退时间求和
def add_time(s, cola='上午迟到', colb='下午迟到'):
    if pd.isnull(s[cola]):
        return s[colb]
    elif pd.isnull(s[colb]):
        return s[cola]
    else:
        (s[cola].apply(t2m) + s[colb].apply(t2m)).apply(m2t)

df_late_a = df_normal.apply(late_time, axis=1, args=('上班打卡', '08:30')).rename('上午迟到')
df_late_p = df_normal.apply(late_time, axis=1, args=('中午打卡', '13:00')).rename('下午迟到')
df_early_a = df_normal.apply(early_time, axis=1, args=('中午打卡', '12:00')).rename('上午早退')
df_early_p = df_normal.apply(early_time, axis=1, args=('下班打卡', '17:30')).rename('下午早退')
df_whole = df_normal.join(df_late_a).join(df_late_p).join(df_early_a).join(df_early_p)
# 迟到、早退时间求和
df_late = df_whole.apply(add_time, axis=1, args=('上午迟到', '下午迟到')).rename('迟到时间')
df_early = df_whole.apply(add_time, axis=1, args=('上午早退', '下午早退')).rename('早退时间')
df_whole = df_whole.join(df_late).join(df_early)

# ------------------------------------ 加班时间 ----------------------------------------
df_more = df_whole.apply(work_more, axis=1).rename('加班时间')
df_whole = df_whole.join(df_more)

# ------------------------------------ 格式化待写入excel的数据 ----------------------------------------

df_aim = pd.DataFrame(data=df_whole, columns=['姓名', '日期', '上班打卡', '中午打卡', '下班打卡', '是否异常', '迟到时间', '早退时间', '加班时间'])


# 汇总数据
df_total = df_aim.groupby('姓名').count().drop(['日期', '上班打卡', '中午打卡', '下班打卡'], axis=1).rename(columns={
    '迟到时间':'迟到天数', '早退时间':'早退天数', '加班时间':'加班天数', '是否异常':'异常天数'
})
df_total = pd.DataFrame(data=df_total, columns=['异常天数', '迟到天数', '早退天数', '加班天数'])
df_total = df_total.sort_values(['异常天数', '迟到天数', '早退天数', '加班天数'], ascending=False)
df_total = df_total[df_total > 0]   # 0 改为NaN


# ------------------------------------ 写入excel ----------------------------------------
writer = pd.ExcelWriter(aim_file)
df_aim.to_excel(writer, sheet_name='考勤数据及汇总', freeze_panes=(1,2), index=False)
df_total.to_excel(writer, sheet_name='考勤数据及汇总', startcol=11)
df_raw.to_excel(writer, '原始考勤数据', freeze_panes=(1, 5))
writer.save()
