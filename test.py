import pandas as pd
import time
# TROC 参数
troc_URL = r'..\Data\2GD_NACH_TAKT.XLSX'
troc_tuzhi = ("5Q0 803 091 L",
"5Q0 803 092 L",
"5Q0 803 501 DE",
"5Q0 803 502 DE",
"5Q0 803 203 AB",
"5Q0 813 101 DR",
"5Q1 800 710 EF",
"2GD 800 709",
"2GA 803 147 B",
"2GD 809 039",
"2GD 809 040",
"2GD 809 051",
"2GD 809 052",
"2GD 800 701 A",
"2GD 800 701 B",
"2GA 821 021",
"2GA 821 022",
"2GD 831 051 B",
"2GD 831 052 B",
"2GD 833 051 D",
"2GD 833 052 D",
"2GA 823 031 A",
"2GA 827 025")
# X55 参数
x55_URL = r'..\Data\GAD_NACH_TAKT.XLSX'
x55_tuzhi = ("5Q0 803 091 J",
"5Q0 803 092 J",
"5Q0 803 501 CS",
"5Q0 803 502 CS",
"5Q0 803 203 AC",
"5Q0 813 101 DQ",
"5Q1 800 710 FL",
"81D 800 709 A",
"81B 803 147 B",
"81D 809 039",
"81D 809 040",
"81D 810 075",
"81D 810 076",
"81D 800 415 A",
"81D 800 415 B",
"81D 800 415 C",
"81A 821 021",
"81A 821 022",
"81D 831 051 B",
"81D 831 052 B",
"81D 833 051 B",
"81D 833 052 B",
"81A 823 151",
"81D 827 025 A")
# VW380 参数
vw380_URL = r'..\Data\5HG_NACH_TAKT.XLSX'
vw380_tuzhi = ("5WD 803 091",
"5WD 803 092",
"5Q0 803 501 CS",
"5Q0 803 502 CS",
"5WA 803 203 E",
"5WA 813 101 BD",
"5WB 800 710 BE",
"5HG 800 709",
"5WB 803 147",
"5HG 809 039",
"5HG 809 040",
"5HG 809 051",
"5HG 809 052",
"5HG 800 701 B",
"5HG 800 701 A",
"5HG 821 021",
"5HG 821 022",
"5HG 831 051 B",
"5HG 831 052 B",
"5HG 833 051",
"5HG 833 052",
"5HG 823 031",
"5HG 827 025")
# 格式化时间函数
def str_int(nok_Time):
    nok_Time = str(nok_Time)
    ok_Time = nok_Time[4:] + nok_Time[2:4] + nok_Time[:2]
    return ok_Time
# 实际操作函数
def get_Excel(open_URL,tuzhi):
    # 打开
    wd = pd.read_excel(open_URL,usecols=[4,12,13,23],converters={'EINDAT': str, 'ENTDAT': str, 'ZEIDAT': str})
    # 刷选
    fc = wd[wd['TEILNR'].isin(tuzhi)].copy()
    # 排序
    fc['TEILNR'] = fc['TEILNR'].astype('category')
    fc['TEILNR'].cat.reorder_categories(tuzhi, inplace=True)
    fc.sort_values('TEILNR', inplace=True)
    # 日期转换  301021  转换为 211030
    fc.iloc[:, 1] = tuple(map(str_int, fc['EINDAT']))       # 图纸启用日期时间转换
    fc.iloc[:, 2] = tuple(map(str_int, fc['ENTDAT']))       # 图纸停用日期时间转换
    # fc.iloc[:, 3] = tuple(map(str_int, fc['ZEIDAT']))     # 图纸使用日期时间转换
    # 按条件刷选，起始日期小于当前时间，停用日期大于当前时间。
    day = time.strftime('%y%m%d')
    fc.replace('nna', '999999')
    fd = fc[(fc.EINDAT < '%s' % (day)) & (fc.ENTDAT > '%s' % (day))]
    # 从新设置索引
    fd.index = range(len(fd))
    return fd
troc = get_Excel(troc_URL, troc_tuzhi)
x55 = get_Excel(x55_URL, x55_tuzhi)
vw380 = get_Excel(vw380_URL, vw380_tuzhi)
# 合并表参数
df = pd.concat([troc, x55, vw380], axis=1, ignore_index=False, join="outer")
# 插入一行
df.loc[-1] = ['TROC零件号', '启用日期', '停用日期', '使用日期', 'X55零件号', '启用日期', '停用日期', '使用日期', 'VW380零件号', '启用日期', '停用日期', '使用日期']
df.index = df.index + 1
df = df.sort_index()
# 保存路径
writer = pd.ExcelWriter('图纸自动更新.xlsx')
df.to_excel(writer, sheet_name='图纸刷选')
# 设置列宽
worksheet = writer.sheets['图纸刷选']
worksheet.set_column('B:M',15)
writer.save()


