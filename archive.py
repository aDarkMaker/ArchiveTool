import os
import re
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from docx import Document
from urllib.parse import urlparse

def sanitize_filename(filename):
    """移除文件名中的非法字符"""
    return re.sub(r'[\\/:*?"<>|]', '', filename)

def get_web_content(url):
    """获取网页内容（强制使用UTF-8编码）"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept-Language': 'zh-CN,zh;q=0.9'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    # 强制使用UTF-8编码解决中文乱码
    response.encoding = 'utf-8'  
    return response.text

def parse_wechat_article(html):
    """解析微信公众号文章内容"""
    # 指定解析编码解决中文乱码
    soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')  
    
    # 提取标题
    title_tag = soup.find('h1', {'class': 'rich_media_title'})
    title = title_tag.get_text().strip() if title_tag else '无标题'
    
    # 提取正文
    content_div = soup.find('div', {'class': 'rich_media_content'})
    content = '\n'.join([p.get_text().strip() for p in content_div.find_all('p')]) if content_div else ''
    
    # 提取图片链接
    imgs = []
    for img in (content_div.find_all('img') if content_div else []):
        img_url = img.get('data-src') or img.get('src')
        if img_url and not img_url.startswith('data:'):
            imgs.append(img_url)
    
    return title, content, imgs

def download_image(img_url, save_path):
    """下载单张图片"""
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Referer': 'https://mp.weixin.qq.com/'
    }
    try:
        response = requests.get(img_url, headers=headers, stream=True, timeout=10)
        if response.status_code == 200:
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            return True
    except Exception as e:
        print(f"图片下载失败: {str(e)}")
    return False

def main():
    # 用户输入
    url = input("请输入微信公众号文章URL: ").strip()
    base_dir = input("请输入存储根目录路径（留空则默认为当前目录）: ").strip() or '.'
    
    # 获取并解析内容
    html = get_web_content(url)
    title, content, imgs = parse_wechat_article(html)
    
    # 创建文件夹
    date_str = datetime.now().strftime("%Y%m%d")
    folder_name = f"{date_str}_{sanitize_filename(title)}"
    folder_path = os.path.join(base_dir, folder_name)
    img_dir = os.path.join(folder_path, '图片')
    
    os.makedirs(img_dir, exist_ok=True)
    
    # 保存文字内容（直接保存在主目录）
    doc = Document()
    doc.add_paragraph(content)
    doc.save(os.path.join(folder_path, '文字.docx'))  # 修改保存路径
    
    # 下载图片
    success_count = 0
    for idx, img_url in enumerate(imgs, 1):
        ext = os.path.splitext(urlparse(img_url).path)[1] or '.jpg'
        save_path = os.path.join(img_dir, f'图片{idx}{ext}')
        if download_image(img_url, save_path):
            success_count += 1
            print(f"已下载图片 {idx}/{len(imgs)}")
    
    print(f"处理完成！文档保存在：{os.path.abspath(folder_path)}")
    print(f"成功下载 {success_count}/{len(imgs)} 张图片")

if __name__ == "__main__":
    main()