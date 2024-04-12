import openpyxl
import jieba
from collections import Counter

# 打开Excel文件
workbook = openpyxl.load_workbook(r'C:\Users\86138\Desktop\京东数据抓取.xlsx')
worksheet = workbook.active

# 选择要处理的列
column = 'A'  # 假设要处理的列是A列

# 获取指定列的值，并进行分词
comments = []
for cell in worksheet[column]:
    comments.append(cell.value)

# 分词
words_list = []
for comment in comments:
    words = jieba.lcut(str(comment))
    words_list.extend(words)

# 统计词频
word_counts = Counter(words_list)

# 输出词频统计结果
for word, count in word_counts.items():
    print(word, count)
