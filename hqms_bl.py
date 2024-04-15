import argparse
import requests
import pandas as pd
from openpyxl import Workbook
import os
import json

# 设置命令行参数
parser = argparse.ArgumentParser(description="处理和上传数据")
parser.add_argument('command', choices=['data', 'upload'], help='运行模式: "data" 或 "upload"')
args = parser.parse_args()

# 配置变量
province = "山东省"
hospital = "临沂某医院"
hqms_cookie = "JSESSIONID=xxxxxxx"
url = "https://blzk3.hqms.org.cn/blzk/diecaseindex/list" #三级、二级、民营按需 blzk3为民营，blzk为三级，blzk2为二级
content_type = "application/x-www-form-urlencoded; charset=UTF-8"

def download_data():
    """下载数据并保存到Excel"""
    data = "pageNum=1&pageSize=5000&a48="
    headers = {
        "Content-Type": content_type,
        "Cookie": hqms_cookie
    }
    response = requests.post(url, data=data, headers=headers)
    response_data = response.json()

    # 保存响应数据到JSON文档
    with open('data.json', 'w') as f:
        json.dump(response_data, f, indent=4)

    print("响应数据已保存到 data.json 文档中。")

    list_data = response_data.get('rows', [])
    
    df = pd.DataFrame(list_data)
    if 'id' in df and 'a48' in df:
        df['path'] = df.apply(lambda x: f"pdf\\{province}_{hospital}_{x['a48']}_{x['b15'].replace('-', '')}.pdf", axis=1)
        #df['path'] = df.apply(lambda x: f"pdf\\{province}_{hospital}_{x['a48'][:-3]}_{x['b15'].replace('-', '')}.pdf", axis=1)
        df['valid'] = df['path'].apply(lambda x: 1 if os.path.exists(x) else 0)

        df.to_excel("output.xlsx", index=False)
    print("数据下载和Excel导出完成！")

def upload_files():
    """上传valid为1的文档，使用预设的hqms_cookie"""
    # 读取之前保存的Excel文档
    df = pd.read_excel("output.xlsx")
    valid_data = df[df['valid'] == 1]

    # 设置请求会话
    session = requests.Session()
    session.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'

    # 将cookie添加到session
    session.cookies.set('JSESSIONID', hqms_cookie.split('=')[1], domain='blzk3.hqms.org.cn', path='/')

    # 遍历有效数据并上传
    for index, row in valid_data.iterrows():
        file_path = row['path']
        if os.path.exists(file_path):
            with open(file_path, 'rb') as file:
                file_content = file.read()
            upload_url = f"https://blzk3.hqms.org.cn/blzk/diecasefile/add?indexid={row['id']}"
            files = {
                'file': (os.path.basename(file_path), file_content, 'application/pdf')
            }
            # 发送文档上传请求
            response = session.post(upload_url, files=files)
            print(f"文档 {file_path} 上传结果: {response.status_code}")

# 根据命令行参数选择运行模式
if args.command == 'data':
    download_data()
elif args.command == 'upload':
    upload_files()
