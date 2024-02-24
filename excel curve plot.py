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
x_segments_2 = [(1, 12), (14, 26), (28, 40), (42, 54), (56, 68), (70, 82), (84, 96), (98, 110), (112, 124)]
y_segments_2 = [(1, 12), (14, 26), (28, 40), (42, 54), (56, 68), (70, 82), (84, 96), (98, 110), (112, 124)]
# 确保X轴和Y轴的数据段数量相同
assert len(x_segments) == len(y_segments), "X and Y segments must have the same number of data segments."
markers = ['*', 'o', 'x', '+', '*', '*', '*', '*', '*']

# 绘制折线图
plt.figure(1)
for i, (x_segments, y_segments) in enumerate(zip(x_segments, y_segments)):
    x_values = df['Load current(A)'].iloc[x_segments[0]:x_segments[1]+1]
    y_values = df['Efficiency'].iloc[y_segments[0]:y_segments[1]+1]
    marker = markers[i % len(markers)]
    plt.plot(x_values, y_values, marker=marker)
    for x, y in zip(x_values.round(1), y_values.round(2)):
        plt.text(x, y, f'{y}', ha='left', va='center', fontsize=8)
plt.xlabel('load current')
plt.ylabel('Efficiency')
plt.title('XFMR700uH - Efficiency')
plt.tick_params(axis='x', width=2)
plt.tick_params(axis='y', width=2)
plt.legend(['70', '120', '170', '220', '270', '320', '370', '420', '470'], loc='lower right')
plt.grid(True)
x_ticks = np.arange(0, 1.4, 0.1)
plt.xticks(x_ticks)
y_ticks = np.arange(60, 85, 2)
plt.yticks(y_ticks)

plt.figure(2)
for i, (x_segments_2, y_segments_2) in enumerate(zip(x_segments_2, y_segments_2)):
    x_values_2 = df['Load current(A)'].iloc[x_segments_2[0]:x_segments_2[1]+1]
    y_values_2 = df['Load power(W)'].iloc[y_segments_2[0]:y_segments_2[1]+1]
    marker = markers[i % len(markers)]
    plt.plot(x_values_2, y_values_2, marker=marker)
    for x, y in zip(x_values_2.round(1), y_values_2.round(2)):
        plt.text(x, y, f'{y}', ha='left', va='center', fontsize=8)
plt.xlabel('load current')
plt.ylabel('Load power(W)')
plt.title('XFMR700uH - Load Power')
plt.tick_params(axis='x', width=2)
plt.tick_params(axis='y', width=2)
plt.legend(['70', '120', '170', '220', '270', '320', '370', '420', '470'], loc='lower right')
plt.grid(True)
x_ticks = np.arange(0, 1.4, 0.1)
plt.xticks(x_ticks)
y_ticks = np.arange(0, 30, 2)
plt.yticks(y_ticks)

plt.show()

