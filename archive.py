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
        # 改进的日期匹配正则表达式
        match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', publish_time_str)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            day = int(match.group(3))
            return f"{year:04}{month:02}{day:02}"
        return datetime.now().strftime("%Y%m%d")
    except Exception as e:
        print(f"日期解析失败: {str(e)}")
        return datetime.now().strftime("%Y%m%d")

def parse_wechat_article(html):
    """解析微信公众号文章内容"""
    # 移除from_encoding参数解决警告
    soup = BeautifulSoup(html, 'html.parser')
    
    # 改进的日期提取方法
    date_str = datetime.now().strftime("%Y%m%d")
    date_pattern = re.compile(r'\d{4}年\d{1,2}月\d{1,2}日')
    
    # 尝试多种选择器查找日期
    publish_time_tag = soup.find('em', id='publish_time') or \
                      soup.find('div', class_='rich_media_meta_text', string=date_pattern) or \
                      soup.find('em', class_='rich_media_meta', string=date_pattern)
    
    if publish_time_tag:
        date_str = parse_date(publish_time_tag.get_text())
    
    # 提取标题
    title_tag = soup.find('h1', {'class': 'rich_media_title'}) or \
               soup.find('h1', id='activity-name')
    title = title_tag.get_text().strip() if title_tag else '无标题'
    
    # 提取内容
    content_div = soup.find('div', {'class': 'rich_media_content'}) or \
                 soup.find('div', id='js_content')
    content_paragraphs = []
    imgs = []
    
    if content_div:
        for element in content_div.descendants:
            if element.name == 'p':
                content_paragraphs.append(element)
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
    """带格式添加段落（修复删除错误）"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    
    try:
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
        
        # 修复空段落删除逻辑
        if not p.text.strip():
            p._element.getparent().remove(p._element)
    except Exception as e:
        print(f"段落处理出错: {str(e)}")
        p._element.getparent().remove(p._element)

def download_image(img_url, save_path):
    """下载单张图片（增加超时和重试）"""
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Referer': 'https://mp.weixin.qq.com/'
    }
    for retry in range(3):
        try:
            response = requests.get(img_url, headers=headers, stream=True, timeout=20)
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    for chunk in response.iter_content(1024):
                        f.write(chunk)
                return True
        except Exception as e:
            if retry == 2:
                print(f"图片下载最终失败: {img_url}")
                return False
            print(f"图片下载重试中({retry+1}/3): {img_url}")
    return False

def main():
    # 用户输入
    url = input("请输入微信公众号文章URL: ").strip()
    base_dir = input("请输入存储根目录路径（留空则默认为当前目录）: ").strip() or '.'
    
    try:
        html = get_web_content(url)
        date_str, title, content_paragraphs, imgs = parse_wechat_article(html)
        
        # 调试输出
        print(f"解析结果：日期={date_str}，标题={title}，段落数={len(content_paragraphs)}，图片数={len(imgs)}")
        
        # 创建文件夹
        folder_name = f"{date_str}_{sanitize_filename(title)}"
        folder_path = os.path.join(base_dir, folder_name)
        img_dir = os.path.join(folder_path, '图片')
        os.makedirs(img_dir, exist_ok=True)
        
        # 创建文档
        doc = Document()
        set_doc_style(doc)
        
        # 添加段落
        for p in content_paragraphs:
            try:
                add_formatted_paragraph(doc, p)
            except Exception as e:
                print(f"跳过无效段落: {str(e)}")
        
        # 保存文档
        doc_path = os.path.join(folder_path, '文字.docx')
        doc.save(doc_path)
        print(f"文档已保存至：{doc_path}")
        
        # 下载图片
        success_count = 0
        for idx, img_url in enumerate(imgs, 1):
            ext = os.path.splitext(urlparse(img_url).path)[1] or '.jpg'
            save_path = os.path.join(img_dir, f'图片{idx}{ext}')
            if download_image(img_url, save_path):
                success_count += 1
        
        print(f"图片下载完成：成功 {success_count}/{len(imgs)}")
        print(f"处理完成！保存路径：{os.path.abspath(folder_path)}")
    
    except Exception as e:
        print(f"程序运行出错: {str(e)}")

if __name__ == "__main__":
    main()