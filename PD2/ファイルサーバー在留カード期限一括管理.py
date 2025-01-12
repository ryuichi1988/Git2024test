# 读取本社FileServer的在留文件自动进行分析.并输出至csv文件

import openpyxl
import datetime
import csv

date = datetime.datetime.now() #日付処理
date = date.strftime('%Y%m%d%H') #フォーマットする。ファイル名

path1 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\グループ外派遣先（湘南）\ダイエットクック小田原在留カード確認リスト.xlsx")
path2 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\グループ外派遣先（神奈川）\銀座コージーコーナー関係\銀座コージーコーナー在留カード確認リスト.xlsx")
path3 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\グループ外派遣先（神奈川）\プライムデリカ第一工場在留カード確認リスト.xlsx")
path4 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\グループ外派遣先（神奈川）\プラィムデリカ第二工場在留カード確認リスト.xlsx")
path5 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\グループ外派遣先（神奈川）\日本フルハーフ在留カード確認リスト.xlsx")
path6 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\ニッセーデリカ湘南工場関係\ニッセーデリカ湘南工場在留カード確認リスト.xlsx")
path7 = (r"\\192.168.1.102\人財開発\★外国籍既存就労スタッフ在留カード関係\ニッセーデリカ神奈川工場関係\ニッセーデリカ神奈川工場在留カード確認リスト.xlsx")
list = [path1,path2,path3,path4,path5,path6,path7] #リストを作る

# csvファイルを作成。
with open("在留カード期限{}.csv".format(date),"w",newline="") as f:
    writer = csv.writer(f) #書き込み
    writer.writerow(["所属","名前","在留期限"])

    for path in list: #遍历list
        file = openpyxl.load_workbook(path)
        workbook = file.active
        for a,b,c in zip(workbook["B"],workbook["E"],workbook["J"]): #把列表写入csv时需要这样操作
            try:
                if c.value < datetime.datetime(year=2024,month=10,day=31): #这个日期前的人PICKUP
                    c = datetime.date(year=c.value.year,month=c.value.month,day=c.value.day)
                    writer.writerow([a.value,b.value,c])
            # 如果是空白cell,则跳过此行.进行下一行
            except TypeError:
                continue
        file.close()
f.close()
    