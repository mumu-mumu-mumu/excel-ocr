#encoding:utf-8

import requests
import base64
import pandas as pd
import os
import json
import configparser
import tkinter as tk
from tkinter import filedialog

# 创建一个窗口
root = tk.Tk()
root.title("输入API Key和Secret Key")

# 创建两个标签和两个输入框
api_key_label = tk.Label(root, text="API Key:")
api_key_entry = tk.Entry(root)
secret_key_label = tk.Label(root, text="Secret Key:")
secret_key_entry = tk.Entry(root)

# 检查配置文件是否存在，如果不存在就创建一个
if not os.path.exists('config.ini'):
    config = configparser.ConfigParser()
    config['DEFAULT'] = {'API_Key': '', 'Secret_Key': ''}
    with open('config.ini', 'w') as f:
        config.write(f)

# 读取配置文件，如果已经有保存的API Key和Secret Key，就显示在输入框中
config = configparser.ConfigParser()
config.read('config.ini')
API_Key = config.get('DEFAULT', 'API_Key')
Secret_Key = config.get('DEFAULT', 'Secret_Key')
if API_Key:
    api_key_entry.insert(0, API_Key)
if Secret_Key:
    secret_key_entry.insert(0, Secret_Key)

# 定义一个函数，保存用户输入的API Key和Secret Key
def save_keys():
    API_Key = api_key_entry.get()
    Secret_Key = secret_key_entry.get()
    config.set('DEFAULT', 'API_Key', API_Key)
    config.set('DEFAULT', 'Secret_Key', Secret_Key)
    with open('config.ini', 'w') as f:
        config.write(f)
    root.destroy()

# 创建一个保存按钮
save_button = tk.Button(root, text="保存", command=save_keys)

# 使用网格布局来排列标签、输入框和按钮
api_key_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
api_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
secret_key_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
secret_key_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
save_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# 进入消息循环，等待用户输入
root.mainloop()

        
url = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=%s&client_secret=%s"%(API_Key,Secret_Key)

payload = ""
headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json'
    }
    
response = requests.request("POST", url, headers=headers, data=payload)
response_str = response.content.decode('utf-8')
response_dict = json.loads(response_str)
access_token = response_dict['access_token']
print(access_token)

# 弹出选择文件夹路径的界面

folder_path = filedialog.askdirectory()

# 如果用户点击了取消按钮，返回的路径为''
if folder_path != '':
    print("您选择的文件夹路径为：", folder_path)
else:
    print("您没有选择文件夹路径")



request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/table"
#搜索文件
input_folder = folder_path
os.chdir(input_folder)
catalog = os.listdir()
for each in catalog:
    # 二进制方式打开图片文件
    f = open(each, 'rb')
    img = base64.b64encode(f.read())
    print(img)
    params = {"image":img}
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print(response.json())
        print(response.text)

        #把文件以json格式保存起来
        name = each.split('.')[0]
        with open(f'{name}.json', 'w', encoding = 'utf-8') as z:
            z.write(response.text)
input_folder = 'D:\\dong\\tu'
output_folder = "%s\\excel"%(folder_path)
#如果无output路径不存在则自动生成excel文件夹
file_path = output_folder
if os.path.exists(file_path) is False:
    os.makedirs(file_path)
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





