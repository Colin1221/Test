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
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 10))

for i, (x_segments, y_segments) in enumerate(zip(x_segments, y_segments)):
    x_values = df['Load current(A)'].iloc[x_segments[0]:x_segments[1]+1]
    y_values_efficiency = df['Efficiency'].iloc[y_segments[0]:y_segments[1]+1]
    y_values_load_power = df['Load power(W)'].iloc[y_segments[0]:y_segments[1]+1]
    marker = markers[i % len(markers)]
    ax1.plot(x_values, y_values_efficiency, marker=marker, label=f'Segment {i+1}')
    ax2.plot(x_values, y_values_load_power, marker=marker, label=f'Segment {i+1}')

    for x, y in zip(x_values.round(1), y_values_efficiency.round(2)):
        ax1.text(x, y, f'{y}', ha='left', va='center', fontsize=8)
    for x, y in zip(x_values.round(1), y_values_load_power.round(2)):
        ax2.text(x, y, f'{y}', ha='left', va='center', fontsize=8)

ax1.set_xlabel('load current')
ax1.set_ylabel('Efficiency')
ax1.set_title('XFMR700uH - Efficiency')
ax1.tick_params(axis='x', width=2)
ax1.tick_params(axis='y', width=2)
ax1.legend(loc='lower right')
ax1.grid(True)
x_ticks = np.arange(0, 1.4, 0.1)
ax1.set_xticks(x_ticks)
y_ticks = np.arange(60, 85, 2)
ax1.set_yticks(y_ticks)

ax2.set_xlabel('load current')
ax2.set_ylabel('Load power(W)')
ax2.set_title('XFMR700uH - Load Power')
ax2.tick_params(axis='x', width=2)
ax2.tick_params(axis='y', width=2)
ax2.legend(loc='lower right')
ax2.grid(True)
x_ticks = np.arange(0, 1.4, 0.1)
ax2.set_xticks(x_ticks)
y_ticks = np.arange(0, 30, 2)
ax2.set_yticks(y_ticks)

plt.tight_layout()
plt.show()



