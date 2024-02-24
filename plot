import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 加载Excel文件
file_path = 'C:/HVSource_Eload/XFMR700uH_CS0.7.xlsx'  # 替换为你的Excel文件路径
# sheet_name = 'XFMR700uH_CS0.7'  # 替换为你要处理的工作表名称
df = pd.read_excel(file_path)

# 定义X轴和Y轴的数据段

x_segments = [(1, 12), (14, 26), (28, 40), (42, 54), (56, 68), (70, 82), (84, 96), (98, 110), (112, 124)]
y_segments = [(1, 12), (14, 26), (28, 40), (42, 54), (56, 68), (70, 82), (84, 96), (98, 110), (112, 124)]

# 确保X轴和Y轴的数据段数量相同
assert len(x_segments) == len(y_segments), "X and Y segments must have the same number of data segments."
markers = ['*', 'o', 'x', '+', '*', '*', '*', '*', '*']

# 绘制折线图
# plt.figure(1)
fig1, ax = plt.subplots()
for i, (x_segments, y_segments) in enumerate(zip(x_segments, y_segments)):
    x_values = df['Load current(A)'].iloc[x_segments[0]:x_segments[1]+1]
    y_values = df['Efficiency'].iloc[y_segments[0]:y_segments[1]+1]
    # x_values_rounded = x_values.round(1)
    # y_values_rounded = y_values.round(2)
    marker = markers[i % len(markers)]
    ax.plot(x_values, y_values, marker=marker)
    for x, y in zip(x_values.round(1),y_values.round(2)):
        ax.text(x, y, f'{y}', ha='left', va='center', fontsize=8)
ax.set_xlabel('load current')
ax.set_ylabel('Efficiency')
ax.set_title('XFMR700uH')
ax.tick_params(axis='x', width=2)
ax.tick_params(axis='y', width=2)
ax.legend(['70', '120', '170', '220', '270', '320', '370', '420', '470'], loc='lower right')
plt.grid(True)
x_ticks = np.arange(0,1.4,0.1)
ax.set_xticks(x_ticks)
y_ticks = np.arange(60,85,2)
ax.set_yticks(y_ticks)
plt.show()

# plt.figure(2)
fig2, ax1 = plt.subplots()
for i, (x_segments, y_segments) in enumerate(zip(x_segments, y_segments)):
    x_values = df['Load current(A)'].iloc[x_segments[0]:x_segments[1]+1]
    y_values = df['Load power(W)'].iloc[y_segments[0]:y_segments[1]+1]
    # x_values_rounded = x_values.round(1)
    # y_values_rounded = y_values.round(2)
    marker = markers[i % len(markers)]
    ax1.plot(x_values, y_values, marker=marker)
    for x, y in zip(x_values.round(1),y_values.round(2)):
        ax1.text(x, y, f'{y}', ha='left', va='center', fontsize=8)
ax1.set_xlabel('load current')
ax1.set_ylabel('Load power(W)')
ax1.set_title('XFMR700uH')
ax1.tick_params(axis='x', width=2)
ax1.tick_params(axis='y', width=2)
ax1.legend(['70', '120', '170', '220', '270', '320', '370', '420', '470'], loc='lower right')
plt.grid(True)
x_ticks = np.arange(0,1.4,0.1)
ax1.set_xticks(x_ticks)
y_ticks = np.arange(0,30,2)
ax1.set_yticks(y_ticks)
plt.show()
