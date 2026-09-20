import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ==================== 第1步：读取两张表 ====================
# 注意：请把 '系统表.xlsx' 和 '手工表.xlsx' 放在本代码同级目录下
try:
    df_sys = pd.read_excel('系统表.xlsx')
    df_hand = pd.read_excel('手工表.xlsx')
    print(" 文件读取成功！")
except FileNotFoundError:
    print(" 报错：请检查文件名是否为 '系统表.xlsx' 和 '手工表.xlsx'，并确保它们在当前文件夹下。")
    exit()

# ==================== 第2步：数据预处理（避坑重点！） ====================
# 把订单号转为字符串，防止“数字+文本”匹配不上
df_sys['订单号'] = df_sys['订单号'].astype(str).str.strip()
df_hand['订单号'] = df_hand['订单号'].astype(str).str.strip()

# 确保金额是数字（如果是带￥符号的文本会报错，这里强转数字）
df_sys['金额'] = pd.to_numeric(df_sys['金额'], errors='coerce')
df_hand['金额'] = pd.to_numeric(df_hand['金额'], errors='coerce')

# ==================== 第3步：核心比对逻辑（用 merge） ====================
# 内连接，只比对两边都存在的订单号（如果只想查系统有但手工没有的，可以改 how='outer'）
compare_df = pd.merge(df_sys, df_hand, on='订单号', how='inner', suffixes=('_系统', '_手工'))

# 如果合并后为空，说明两张表没有共同订单号
if compare_df.empty:
    print("警告：两张表没有匹配的订单号，请检查数据。")
    exit()

# 判断金额是否一致（注意：浮点数比较建议用 round 防止精度误差）
compare_df['金额_系统'] = compare_df['金额_系统'].round(2)
compare_df['金额_手工'] = compare_df['金额_手工'].round(2)
compare_df['差异状态'] = compare_df.apply(
    lambda row: '不一致' if row['金额_系统'] != row['金额_手工'] else '一致', axis=1
)

# 筛选出有差异的记录
result_df = compare_df[compare_df['差异状态'] == '不一致'].copy()

# 整理输出列（只保留关键信息，避免表格太乱）
output_cols = ['订单号', '金额_系统', '金额_手工', '差异状态']
result_df = result_df[output_cols]

# 保存为Excel（不带颜色先）
result_df.to_excel('差异明细.xlsx', index=False)
print(f" 已找到 {len(result_df)} 条差异数据，已生成 '差异明细.xlsx'")

# ==================== 第4步：给差异行标红（装逼加分项） ====================
if len(result_df) > 0:
    wb = load_workbook('差异明细.xlsx')
    ws = wb.active
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    # 遍历行（从第2行开始，因为第1行是表头）
    # 差异状态在第4列（D列）
    for row in range(2, ws.max_row + 1):
        status_cell = ws.cell(row, 4)  # D列
        if status_cell.value == '不一致':
            # 把这一行的A到D列都标红
            for col in range(1, 5):  # 1(A),2(B),3(C),4(D)
                ws.cell(row, col).fill = red_fill

    wb.save('差异明细_标红.xlsx')
    print(" 标红文件已生成：差异明细_标红.xlsx (红色行为异常数据)")
else:
    print(" 恭喜，所有数据核对一致，没有差异！")