# encoding:utf-8

import requests
import base64
import pandas as pd
import os
import json



request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/table"
#搜索文件
os.chdir("D:\\dong")
catalog = os.listdir()
for each in catalog:
    # 二进制方式打开图片文件
    f = open(each, 'rb')
    img = base64.b64encode(f.read())
    print(img)
    params = {"image":img}
    #通过下面注释的程序可以获得access token
    access_token = 'access_token'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print(response.json())

        #把文件以json格式保存起来
        name = each.split('.')[0]
        with open(f'{name}.json', 'w', encoding = 'utf-8') as z:
            z.write(response.text)

input_folder = "D:\\dong"
output_folder = "C:\\Users\\mthmu\\OneDrive\\桌面\\工行\\python\\大作业"

# 遍历文件夹中的 JSON 文件
for filename in os.listdir(input_folder):
    if filename.endswith(".json"):
        # 读取 JSON 文件内容并转换为 Python 对象
        with open(os.path.join(input_folder, filename), "r", encoding='utf-8') as f:
            data = json.load(f)
        table_data = data['tables_result'][0]['body']
        # 将表格数据转换成pandas DataFrame对象
        df = pd.DataFrame()
        for row in table_data:
            # 获取单元格位置和内容
            row_start = row['row_start']
            row_end = row['row_end']
            col_start = row['col_start']
            col_end = row['col_end']
            words = row['words']
    
            # 将单元格内容添加到DataFrame中
            for r in range(row_start, row_end+1):
                for c in range(col_start, col_end+1):
                    df.loc[r, c] = words


        # 保存为 Excel 文件
        output_file = os.path.splitext(filename)[0] + ".xlsx"
        output_path = os.path.join(output_folder, output_file)
        df.to_excel(output_path, index=False)

#下面的代码用来获取access_token
'''
import requests
import json


def main():
        
    url = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=xx&client_secret=xx  "
    
    payload = ""
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    
    response = requests.request("POST", url, headers=headers, data=payload)
    
    print(response.text)
'''
'''
if __name__ == '__main__':
    main()
'''