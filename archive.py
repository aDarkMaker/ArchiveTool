import os
import re
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from urllib.parse import urlparse

def sanitize_filename(filename):
    """移除文件名中的非法字符"""
    return re.sub(r'[\\/:*?"<>|]', '', filename)

def get_web_content(url):
    """获取网页内容"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept-Language': 'zh-CN,zh;q=0.9'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    response.encoding = 'utf-8'
    return response.text

def parse_date(publish_time_str):
    """解析微信发布日期"""
    try:
        return datetime.strptime(publish_time_str.strip(), "%Y年%m月%d日 %H:%M").strftime("%Y%m%d")
    except:
        return datetime.now().strftime("%Y%m%d")

def parse_wechat_article(html):
    """解析微信公众号文章内容"""
    soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
    
    # 提取发布日期
    publish_time_tag = soup.find('em', {'id': 'publish_time'})
    date_str = parse_date(publish_time_tag.get_text()) if publish_time_tag else datetime.now().strftime("%Y%m%d")
    
    # 提取标题
    title_tag = soup.find('h1', {'class': 'rich_media_title'})
    title = title_tag.get_text().strip() if title_tag else '无标题'
    
    # 提取内容
    content_div = soup.find('div', {'class': 'rich_media_content'})
    content_paragraphs = []
    imgs = []
    cutoff_flag = False
    
    if content_div:
        for element in content_div.children:
            if cutoff_flag:
                break
                
            if element.name == 'p':
                text = element.get_text().strip()
                if '审核' in text and ('|' in text or '｜' in text):
                    cutoff_flag = True
                    break
                
                content_paragraphs.append(element)
                # 提取当前段落中的图片
                for img in element.find_all('img'):
                    img_url = img.get('data-src') or img.get('src')
                    if img_url and not img_url.startswith('data:'):
                        imgs.append(img_url)
                        
            elif element.name == 'img':
                img_url = element.get('data-src') or element.get('src')
                if img_url and not img_url.startswith('data:'):
                    imgs.append(img_url)
    
    return date_str, title, content_paragraphs, imgs

def set_doc_style(doc):
    """设置文档默认样式"""
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def add_formatted_paragraph(doc, paragraph):
    """带格式添加段落"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)  # 段后间距为0
    
    for content in paragraph.contents:
        if isinstance(content, str):
            run = p.add_run(content.strip())
        elif content.name == 'strong':
            run = p.add_run(content.get_text().strip())
            run.bold = True
        elif content.name == 'span':
            run = p.add_run(content.get_text().strip())
            if 'color:' in content.get('style', ''):
                color_str = re.findall(r'color:(#[0-9a-fA-F]+)', content['style'])
                if color_str:
                    rgb = tuple(int(color_str[0][i:i+2], 16) for i in (1, 3, 5))
                    run.font.color.rgb = RGBColor(*rgb)
    
    # 过滤空段落
    if not p.text.strip():
        doc.paragraphs.remove(p)

def download_image(img_url, save_path):
    """下载单张图片（增加重试机制）"""
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Referer': 'https://mp.weixin.qq.com/'
    }
    for _ in range(3):  # 最多重试3次
        try:
            response = requests.get(img_url, headers=headers, stream=True, timeout=15)
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    for chunk in response.iter_content(1024):
                        f.write(chunk)
                return True
        except Exception as e:
            print(f"图片下载失败（重试中）: {str(e)}")
    return False

def main():
    # 用户输入
    url = input("请输入微信公众号文章URL: ").strip()
    base_dir = input("请输入存储根目录路径（留空则默认为当前目录）: ").strip() or '.'
    
    # 获取并解析内容
    html = get_web_content(url)
    date_str, title, content_paragraphs, imgs = parse_wechat_article(html)
    
    # 创建文件夹
    folder_name = f"{date_str}_{sanitize_filename(title)}"
    folder_path = os.path.join(base_dir, folder_name)
    img_dir = os.path.join(folder_path, '图片')
    os.makedirs(img_dir, exist_ok=True)
    
    # 创建带格式的Word文档
    doc = Document()
    set_doc_style(doc)
    
    # 添加内容段落
    prev_has_content = False
    for p in content_paragraphs:
        current_text = p.get_text().strip()
        if current_text:
            add_formatted_paragraph(doc, p)
            prev_has_content = True
        elif prev_has_content:  # 只允许保留一个空行
            doc.add_paragraph()
            prev_has_content = False
    
    # 保存文档
    doc.save(os.path.join(folder_path, '文字.docx'))
    
    # 下载图片
    success_count = 0
    for idx, img_url in enumerate(imgs, 1):
        ext = os.path.splitext(urlparse(img_url).path)[1] or '.jpg'
        save_path = os.path.join(img_dir, f'图片{idx}{ext}')
        if download_image(img_url, save_path):
            success_count += 1
            print(f"已下载图片 {idx}/{len(imgs)}")
        else:
            print(f"图片下载失败：{img_url}")
    
    print(f"处理完成！文档保存在：{os.path.abspath(folder_path)}")
    print(f"成功下载 {success_count}/{len(imgs)} 张图片")

if __name__ == "__main__":
    main()