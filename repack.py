import sys
import os
import shutil
import subprocess
import zipfile
import re
from PIL import Image, ImageStat

try:
    from mobi.extract import extract as mobi_extract
except ImportError:
    print("================================================================")
    print("致命错误：缺少核心解包模块 'mobi'！")
    print("请在终端或 CMD 中运行以下命令进行安装：")
    print("pip install mobi")
    print("================================================================")
    input("按回车键退出...")
    sys.exit(1)

# --- Kindle 2022 入门款 硬件配置 ---
KINDLE_WIDTH = 1072
KINDLE_HEIGHT = 1448
BLANK_THRESHOLD = 5.0 

def is_blank_page(img):
    try:
        stat = ImageStat.Stat(img)
        if stat.stddev[0] < BLANK_THRESHOLD: return True
        if img.width < 10 or img.height < 10: return True
    except:
        pass
    return False

def get_ordered_images_from_extracted_mobi(extract_dir):
    print("正在破译 MOBI 排版说明书 (OPF/HTML Spine)...")
    opf_file = None
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.lower().endswith('.opf'):
                opf_file = os.path.join(root, file)
                break
        if opf_file: break
        
    if not opf_file:
        print("错误：无法在解包数据中找到 OPF 目录树文件。")
        return []
        
    with open(opf_file, 'r', encoding='utf-8', errors='ignore') as f:
        opf_data = f.read()
        
    manifest_items = {}
    for match in re.finditer(r'<item\s+[^>]*id=["\']([^"\']+)["\'][^>]*href=["\']([^"\']+)["\']', opf_data, re.IGNORECASE):
        manifest_items[match.group(1)] = match.group(2)
        
    ordered_images = []
    opf_base_dir = os.path.dirname(opf_file)
    
    for match in re.finditer(r'<itemref\s+[^>]*idref=["\']([^"\']+)["\']', opf_data, re.IGNORECASE):
        idref = match.group(1)
        if idref in manifest_items:
            html_rel_path = manifest_items[idref]
            html_full_path = os.path.join(opf_base_dir, html_rel_path)
            
            if os.path.exists(html_full_path):
                 with open(html_full_path, 'r', encoding='utf-8', errors='ignore') as f:
                     html_data = f.read()
                     # --- 核心修复点：将 search 改为 finditer，扫出同一个 HTML 内的所有图片！ ---
                     for img_match in re.finditer(r'<img[^>]+src=["\']([^"\']+)["\']', html_data, re.IGNORECASE):
                         img_rel_path = img_match.group(1)
                         img_full_path = os.path.normpath(os.path.join(os.path.dirname(html_full_path), img_rel_path))
                         if os.path.exists(img_full_path):
                             if img_full_path not in ordered_images:
                                 ordered_images.append(img_full_path)
                                 
    print(f"成功破译！共锁定 {len(ordered_images)} 张高清原图的正确阅读顺序。")
    return ordered_images

def process_ordered_images(ordered_files, output_name, kindlegen_path, base_dir, target_dir):
    temp_dir = os.path.join(base_dir, "mobi_build_temp")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    print("开始清洗空白页并精准适配 Kindle 2022 分辨率...")
    Image.MAX_IMAGE_PIXELS = None 

    valid_images = []
    for img_path in ordered_files:
        filename = os.path.basename(img_path)
        try:
            with Image.open(img_path) as img:
                img = img.convert('L')
                
                if is_blank_page(img):
                    print(f"[-] 智能剔除空白页: {filename}")
                    continue

                width_ratio = KINDLE_WIDTH / img.width
                height_ratio = KINDLE_HEIGHT / img.height
                scale = min(width_ratio, height_ratio)
                
                new_width = int(img.width * scale)
                new_height = int(img.height * scale)
                
                img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

                bg = Image.new('L', (KINDLE_WIDTH, KINDLE_HEIGHT), 255)
                offset_x = (KINDLE_WIDTH - new_width) // 2
                offset_y = (KINDLE_HEIGHT - new_height) // 2
                bg.paste(img_resized, (offset_x, offset_y))
                
                out_filename = f"page_{len(valid_images):04d}.jpg"
                out_path = os.path.join(temp_dir, out_filename)
                bg.save(out_path, 'JPEG', quality=90) 
                valid_images.append(out_filename)
        except Exception as e:
             pass

    if len(valid_images) == 0:
         print("错误：没有提取到有效图片。")
         return False

    print("生成 Kindle 原生全屏漫画元数据 (强制锁定视口与零边距)...")
    html_files = []
    
    for img_file in valid_images:
        html_name = img_file.replace('.jpg', '.html')
        html_files.append(html_name)
        with open(os.path.join(temp_dir, html_name), 'w', encoding='utf-8') as f:
            f.write(f'<?xml version="1.0" encoding="UTF-8"?>\n'
                    f'<!DOCTYPE html>\n'
                    f'<html xmlns="http://www.w3.org/1999/xhtml">\n'
                    f'<head><title>Page</title>\n'
                    f'<meta name="viewport" content="width={KINDLE_WIDTH}, height={KINDLE_HEIGHT}" />\n'
                    f'<style>body, div, img {{ margin: 0; padding: 0; border: 0; }} body {{ width: {KINDLE_WIDTH}px; height: {KINDLE_HEIGHT}px; overflow: hidden; }} img {{ width: {KINDLE_WIDTH}px; height: {KINDLE_HEIGHT}px; display: block; }}</style>\n'
                    f'</head>\n'
                    f'<body><img src="{img_file}" alt="comic page"/></body></html>')

    book_title = output_name.replace("_重构版.mobi", "").replace(".mobi", "")
    
    opf_content = ['<?xml version="1.0" encoding="utf-8"?>',
                   '<package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="uid">',
                   f'<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">',
                   f'<dc:title>{book_title}</dc:title><dc:language>zh</dc:language>',
                   f'<meta name="fixed-layout" content="true"/>',
                   f'<meta name="original-resolution" content="{KINDLE_WIDTH}x{KINDLE_HEIGHT}"/>',
                   f'<meta name="book-type" content="comic"/>',
                   f'<meta name="zero-gutter" content="true"/>',
                   f'<meta name="zero-margin" content="true"/>',
                   f'</metadata>',
                   '<manifest>']
    
    opf_content.append('<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>')

    for html_file, img_file in zip(html_files, valid_images):
        img_id = img_file.replace('.', '_')
        html_id = html_file.replace('.', '_')
        opf_content.append(f'<item id="{html_id}" href="{html_file}" media-type="application/xhtml+xml"/>')
        opf_content.append(f'<item id="{img_id}" href="{img_file}" media-type="image/jpeg"/>')
    
    opf_content.append('</manifest><spine toc="ncx">')
    
    for html_file in html_files:
        html_id = html_file.replace('.', '_')
        opf_content.append(f'<itemref idref="{html_id}"/>')
    opf_content.append('</spine></package>')

    opf_path = os.path.join(temp_dir, "content.opf")
    with open(opf_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(opf_content))
    
    ncx_path = os.path.join(temp_dir, "toc.ncx")
    with open(ncx_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?><ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="uid"/></head><docTitle><text>Comic</text></docTitle><navMap><navPoint id="navPoint-1" playOrder="1"><navLabel><text>Start</text></navLabel><content src="' + html_files[0] + '"/></navPoint></navMap></ncx>')

    print("开始调用 kindlegen 编译 (这可能需要几分钟)...")
    process = subprocess.Popen([kindlegen_path, opf_path, "-c1", "-o", output_name, "-dont_append_source"],
                     stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()

    output_target_path = os.path.join(target_dir, output_name)
    if os.path.exists(output_target_path):
        os.remove(output_target_path)
    shutil.move(os.path.join(temp_dir, output_name), output_target_path)
    shutil.rmtree(temp_dir) 
    print(f"\n[大功告成] 完美重构版已生成: {output_target_path}")
    return True

if __name__ == "__main__":
    import traceback
    try:
        if sys.platform.startswith('win'):
            os.system('chcp 65001 >nul')

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        KINDLEGEN_EXE_PATH = os.path.join(BASE_DIR, "kindlegen.exe")

        if not os.path.exists(KINDLEGEN_EXE_PATH):
            print("错误：找不到 kindlegen.exe！")
            input("按回车键退出...")
            sys.exit(1)

        if len(sys.argv) < 2:
            print("用法：请将【MOBI文件】拖拽到此脚本上！")
            input("按回车键退出...")
            sys.exit(1)

        input_path = sys.argv[1]
        target_output_dir = os.path.dirname(os.path.abspath(input_path))
        processing_success = False

        if input_path.lower().endswith(('.mobi', '.azw3')):
            print(f"--------------------------------------------------")
            print(f"检测到已编译的 Kindle 文件！")
            print(f"正在启动 mobi 核心引擎进行标准解包，请耐心等待...")
            print(f"--------------------------------------------------")
            
            # 使用 mobi 库进行极其标准的解包
            temp_extract_dir, _ = mobi_extract(input_path)
            
            try:
                file_name = os.path.basename(input_path)
                name_without_ext = os.path.splitext(file_name)[0]
                
                ordered_images = get_ordered_images_from_extracted_mobi(temp_extract_dir)
                
                if ordered_images:
                    processing_success = process_ordered_images(
                        ordered_images, 
                        f"{name_without_ext}_重构版.mobi", 
                        KINDLEGEN_EXE_PATH, 
                        BASE_DIR, 
                        target_output_dir
                    )
            finally:
                if os.path.exists(temp_extract_dir):
                    shutil.rmtree(temp_extract_dir, ignore_errors=True)
            
        else:
            print("目前此终极版本专为破解 MOBI/AZW3 乱序问题设计，请拖入电子书文件。")

        if processing_success:
             print("\n处理成功！窗口将在 5 秒后关闭。")
             import time
             time.sleep(5)
        else:
             input("\n处理失败，请查看上方错误信息。按回车键退出...")

    except Exception as e:
        error_msg = traceback.format_exc()
        with open(os.path.join(BASE_DIR, "error_log.txt"), "w", encoding="utf-8") as f:
            f.write(error_msg)
        print("\n[崩溃] 程序出错，详情已写入 error_log.txt")
        input("按回车键退出...")