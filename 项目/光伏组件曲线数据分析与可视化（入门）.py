import xlrd
xlsx= xlrd.open_workbook('d://python/工作.xlsx')
biaoqian =xlsx.sheet_by_index(0)
hang = biaoqian.nrows
for i in range(1,hang):
    I = biaoqian.cell_value(i,0)
    U = biaoqian.cell_value(i,1)
    Pw = I*U
    print(Pw)

#加载两个模块
import xlrd
import matplotlib.pyplot as plt
#打开文件,获取标签页
xlsx = xlrd.open_workbook('d:/python/工作.xlsx')
biaoqian = xlsx.sheet_by_index(0)
hang = biaoqian.nrows
current = []
voltage = []
power = []
for i in range(1, hang):
    I = biaoqian.cell_value(i, colx=0)
    U = biaoqian.cell_value(i, colx=1)
    P = I * U
    current.append(I)
    voltage.append(U)
    power.append(P)
    #把数据都放入相应的列表
# 绘制IV曲线
plt.figure(figsize=(10, 5))
plt.subplot(1, 2, 1)
plt.plot(voltage, current, 'b-o')
plt.title(f'I-V数据分析可视化')
plt.xlabel('电压U (V)')
plt.ylabel('电流I (A)')
#绘制PV曲线
plt.subplot(1, 2, 2)
plt.plot(voltage, power, 'r-o')
plt.title('PV Curve')
plt.xlabel('Voltage (V)')
plt.ylabel('Power (W)')
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.tight_layout()
plt.show()







# import openpyxl
# import matplotlib.pyplot as plt
#
# # 设置中文字体，解决中文显示乱码
# plt.rcParams['font.sans-serif'] = ['SimHei']
# plt.rcParams['axes.unicode_minus'] = False
#
# # 读取Excel光伏数据
# wb = openpyxl.load_workbook('d:/python/工作.xlsx')
# sheet = wb.active
# voltage = []
# current = []
# power = []
#
# # 提取数据并计算功率
# for row in sheet.iter_rows(min_row=2, values_only=True):
#     u = row[1]  # 第2列电压
#     i = row[0]  # 第1列电流
#     if u is not None and i is not None:
#         voltage.append(u)
#         current.append(i)
#         power.append(u * i)

# 创建画布和主坐标轴（左侧y轴：电流）
fig, ax1 = plt.subplots(figsize=(10, 6))

# 绘制I-V曲线（主坐标轴）
# ax1.plot(voltage, current, 'b-o', linewidth=1.5, markersize=4, label='I-V曲线')
# ax1.set_xlabel('电压 (V)', fontsize=12)
# ax1.set_ylabel('电流 (A)', fontsize=12, color='b')
# ax1.tick_params(axis='y', labelcolor='b')  # 左侧y轴刻度标为蓝色
# ax1.grid(True, linestyle='--', alpha=0.7)
#
# # 创建次坐标轴（右侧y轴：功率）
# ax2 = ax1.twinx()
# # 绘制P-V曲线（次坐标轴）
# ax2.plot(voltage, power, 'r-o', linewidth=1.5, markersize=4, label='P-V曲线')
# ax2.set_ylabel('功率 (W)', fontsize=12, color='r')
# ax2.tick_params(axis='y', labelcolor='r')  # 右侧y轴刻度标为红色
#
# # 添加总标题和图例
# plt.title('光伏组件I-V/P-V曲线对比', fontsize=14)
# # 合并两个坐标轴的图例
# lines1, labels1 = ax1.get_legend_handles_labels()
# lines2, labels2 = ax2.get_legend_handles_labels()
# ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
#
# plt.tight_layout()
# plt.show()
#
# wb.close()