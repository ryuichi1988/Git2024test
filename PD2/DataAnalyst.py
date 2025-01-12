# 什么时候能把这些内容都弄明白那也厉害啊。

import pandas as pd
import matplotlib.pyplot as plt

# Data preparation
data = {
    "全体オーダー差異": [0, 0, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 3, 1, 0, 0, 5, 14, 8, 8, 3, 13, 10, 13, 13, 5, -2],
    "全体オーダー人数": [32, 16, 12, 23, 22, 18, 19, 19, 13, 13, 18, 17, 15, 23, 27, 17, 11, 17, 17, 23, 18, 18, 15, 11, 23, 11, 11, 16, 22, 22, 22],
    "全体出勤人数": [42, 38, 36, 43, 43, 38, 33, 32, 33, 39, 36, 40, 35, 34, 35, 34, 35, 36, 39, 42, 36, 41, 41, 41, 42, 49, 44, 47, 49, 45, 40]
}

df = pd.DataFrame(data)

# Display basic statistics
basic_stats = df.describe()

# Visualize the data trends
df.plot(title="全体オーダー差異, 全体オーダー人数, 全体出勤人数のトレンド", marker='o', figsize=(10, 6))
plt.xlabel("Time Index")
plt.ylabel("Values")
plt.legend(["全体オーダー差異", "全体オーダー人数", "全体出勤人数"])
plt.grid(True)
plt.show()

# Displaying the statistics
import ace_tools as tools; tools.display_dataframe_to_user(name="基本統計量", dataframe=basic_stats)
