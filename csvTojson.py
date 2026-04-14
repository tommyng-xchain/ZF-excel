import csv

def getjson(data):    
    items = []
    count = 0

    for key in data:
        item = {}
        # print(key)
        if key:
            item['label'] = key
            item['id'] = id[count]
            items.append(item)
        count+=1
    print(items)

# 開啟CSV檔案
with open(r"G:\Users\zzho8\Downloads\benz_ui.csv", "r") as file:
    # 建立CSV讀取器
    reader = csv.reader(file)

    # 讀取標題列
    empty = next(reader)
    mbhk = next(reader)
    # print(mbhk)
    MBZF = next(reader)
    MBZF = next(reader)
    MBZF = next(reader)
    # print(MBZF)
    MBFS = next(reader)
    MBFS = next(reader)
    # print(MBFS)
    id = next(reader)
    id = next(reader)
    id = next(reader)
    # print(id)
    getjson(mbhk)
    getjson(MBZF)
    getjson(MBFS)
    