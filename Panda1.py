import jieba
import pandas as pd
import numpy as np
import xlwt
import wordcloud
from collections import Counter
import matplotlib.pyplot as matplo
pd.set_option('display.max_columns', 5)
pd.set_option('display.max_rows', 250)
pd.set_option('display.width', 50)
class DataProcess:
    def __init__(self, xls, sheet):#初始化类型
        self.xls = xls
        self.data = pd.read_excel(xls, sheet_name=sheet)
    def drop_blank(self,coloum_name):#删除存在的空行
        self.data = self.data.dropna(subset=coloum_name)
        self.data.reset_index(drop=True, inplace=True)
        print("已完成过滤！")
        return
    def average_score(self):#计算评分平均数和有用的平均数
        gradelist = ['很差', '较差', '还行', '推荐', '力荐']
        sample = list(self.data['score'].values)
        sum = 0.0
        sum2 = 0.0
        for i in sample:
            for j in gradelist:
                if i == j:
                    sum += (gradelist.index(j)+1)
                else:
                    continue
        avg1 = sum/len(sample)
        for i in list(self.data['usefulNum'].values):
            sum2 += i
        avg2 = sum2/len(sample)
        print("评分的平均值为:",avg1)
        print("有用的平均值为:", avg2)
        return

    def find_useful(self):#最有用的
        self.data.sort_values(by='usefulNum')
        s = self.data.loc[0]
        print("有用数最高的数据是:\n{:}".format(s))
        return

    def relation_analysis(self):#分析相关性

        l1 = list(self.data.score)
        l2 = ['很差', '较差', '还行', '推荐', '力荐']
        for i in range(len(l1)):
            for j in range(len(l2)):
                if l1[i] == l2[j]:
                    self.data.at[i, 'score'] = j + 1#data.at定位在i行的某列的值
        data = pd.DataFrame({"score": list(self.data.score), "usefulNum": list(self.data.usefulNum)})
        print("相关性:")
        print(data.corr())

    def word_cloud(self):#绘制词云
        list2 = list(self.data['comment'])
        ls = []
        for i in list2:
            i2 = jieba.lcut(str(i))
            for j in i2:
                if len(j) > 1:
                    ls.append(j)
        top3 = Counter(ls).most_common(3)  # 计数
        print(top3)
        print("出现最多的三个词是：")
        for i in top3:
            print(i[0] + ":出现了{}次".format(i[1]))
            txt = ' '.join(ls)
            input("按任意键绘制词云")
            w = wordcloud.WordCloud(font_path="msyh.ttc",
                                    width=500, height=350,
                                    background_color='white',
                                    max_words=40
                                    )
            w.generate(txt)
            matplo.figure("词云")
            matplo.imshow(w)
            matplo.axis("off")
            matplo.show()
            return
    def save_to_excel(self):
        self.data.to_excel('daoban.xls', sheet_name='0')
        print('数据已保存')
        return
obj = DataProcess(r'C:\Users\wujun\PycharmProjects\SpiderTest\text.xls',sheet=0)#创建处理对象
obj.drop_blank('score')#清除'score'中的空值行
obj.drop_blank('usefulNum')#清除'usefulNum'中的空值行
obj.drop_blank('time')#清除'time'中的空值行
obj.drop_blank('comment')#清除'comment'中的空值行
obj.average_score()#计算评分平均值和有用平均值
obj.find_useful()#找出最有用的 评论
obj.relation_analysis()#相关性分析
obj.word_cloud()#词云
obj.save_to_excel()#保存至excel






