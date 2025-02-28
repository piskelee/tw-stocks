import base64
import os
import requests
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from config import token, owner, repo


def fetch_stock_data(stock_code: str, start_date: str, end_date: str) -> pd.DataFrame:
    stock_data = yf.Ticker(stock_code)
    historical_data = stock_data.history(start=start_date, end=end_date)
    print(historical_data)
    return historical_data


def calculate_bollinger_bands(data: pd.DataFrame, window: int = 20) -> pd.DataFrame:
    """
    计算布林带
    """
    data['Middle Band'] = data['Close'].rolling(window=window).mean()
    data['Upper Band'] = data['Middle Band'] + (data['Close'].rolling(window=window).std() * 2)
    data['Lower Band'] = data['Middle Band'] - (data['Close'].rolling(window=window).std() * 2)
    return data


def plot_bollinger_bands(data: pd.DataFrame, stock_code: str, save_path: str) -> None:
    """
    绘制布林带图，当 Close 价格低于 Lower Band 时显示 Close 数值，并在图表上标注最后一笔数据的日期。
    """
    plt.figure(figsize=(8, 5))  # 800x500 像素

    # 绘制布林带和收盘价
    plt.plot(data['Close'], label='Close Price', color='blue', linewidth=1.5)
    plt.plot(data['Middle Band'], label='Middle Band', linestyle='--', color='orange')
    plt.plot(data['Upper Band'], label='Upper Band', linestyle='--', color='green')
    plt.plot(data['Lower Band'], label='Lower Band', linestyle='--', color='red')
    plt.fill_between(data.index, data['Lower Band'], data['Upper Band'], color='gray', alpha=0.1)

    # 筛选 Close 低于 Lower Band 的点
    below_lower_band = data[data['Close'] < data['Lower Band']]
    plt.scatter(below_lower_band.index, below_lower_band['Close'], color='darkred', s=25, label='Below Lower Band')

    # 显示 Close 值，调整字体和位置
    for idx, row in below_lower_band.iterrows():
        plt.text(idx, row['Close'], f"{row['Close']:.2f}", color='darkred', fontsize=8,
                 ha='center', va='top', bbox=dict(facecolor='white', alpha=0.5, edgecolor='none'))

    # 获取最后一笔数据的日期
    last_date = data.index[-1].strftime('%Y-%m-%d')

    # 设置标题和标签
    plt.title(f'Bollinger Bands for {stock_code} ({last_date})')
    plt.xlabel('Date')
    plt.ylabel('Price')
    plt.legend()
    plt.grid(True, linestyle='--', alpha=0.5)

    # 保存图像
    plt.savefig(save_path, format='png', dpi=300, bbox_inches='tight')
    print(f"图像已保存为 {save_path}")

    # 关闭图表
    plt.close()


def generate_html_with_images(image_paths: list, stock_codes: list) -> str:
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>TW ETF BAND</title>
    </head>
    <body>
        <h1>TW ETF BAND</h1>
    """
    # Add each image to the HTML
    for stock_code, image_path in zip(stock_codes, image_paths):
        html_content += f"""
        <img src="{image_path}" style="width:600px; height:300px;" alt="Bollinger Bands Chart for {stock_code}">
        """

    html_content += """
    </body>
    </html>
    """
    return html_content


def upload_files_github(local_directory) -> None:
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
    }
    file_paths = []
    for root, directories, files in os.walk(local_directory):
        for filename in files:
            file_paths.append(os.path.join(root, filename))
    for local_file_path in file_paths:
        with open(local_file_path, 'rb') as file:
            file_content = base64.b64encode(file.read()).decode('utf-8')
        file_name = os.path.basename(local_file_path)
        url = f'https://api.github.com/repos/{owner}/{repo}/contents/{file_name}'
        data = {
            'message': f'Upload {file_name}',
            'content': file_content,
        }
        response = requests.put(url, headers=headers, json=data)
        if response.status_code == 201:
            print(f'文件 {file_name} 上傳成功！')
        else:
            print(f'文件 {file_name} 上傳失敗。CODE：{response.status_code}')
            print(response.text)


def del_github_files() -> None:
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
    }
    url = f'https://api.github.com/repos/{owner}/{repo}/contents/'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        files = response.json()
        for file_info in files:
            if file_info['type'] == 'file':
                file_name = file_info['name']
                delete_url = f'https://api.github.com/repos/{owner}/{repo}/contents/{file_name}'
                delete_data = {
                    'message': f'Delete {file_name}',
                    'sha': file_info['sha']
                }
                delete_response = requests.delete(delete_url, headers=headers, json=delete_data)
                if delete_response.status_code == 200:
                    print(f'github文件 {file_name} 刪除成功！')
                else:
                    print(f'github文件 {file_name} 刪除失敗！！ code:{delete_response.status_code}')
                    print(delete_response.text)
    else:
        print(f'get ng code:{response.status_code}')


if __name__ == "__main__":
    # 设置文件保存路径
    local_directory = 'D:\\StocKPlay\\StocKPlay'
    stock_ids = ["0050.TW", "0056.TW", "006208.TW", "00713.TW", "00878.TW", "00919.TW" ,"2812.TW"]
    start_date = (datetime.today() - timedelta(days=365)).strftime('%Y-%m-%d')
    end_date = datetime.today().strftime('%Y-%m-%d')

    # github delete
    del_github_files()
    image_paths = []

    # 遍历股票代码列表
    for stock_id in stock_ids:
        print(f"正在处理股票代码: {stock_id}")
        stock_data = fetch_stock_data(stock_id, start_date, end_date)
        if stock_data.empty:
            print(f"无法获取 {stock_id} 在 {start_date} 至 {end_date} 期间的数据。")
        else:
            stock_data_with_bands = calculate_bollinger_bands(stock_data)
            save_path = os.path.join(local_directory, f"{stock_id}_bollinger_bands.png")
            plot_bollinger_bands(stock_data_with_bands, stock_id, save_path)

            # Add image path to list for HTML generation
            image_paths.append(f"{stock_id}_bollinger_bands.png")

    # 生成HTML内容并保存
    html_content = generate_html_with_images(image_paths, stock_ids)
    html_filename = os.path.join(local_directory, "index.html")
    with open(html_filename, "w") as html_file:
        html_file.write(html_content)
    print(f"HTML文件已保存为 {html_filename}")
    upload_files_github(local_directory)
