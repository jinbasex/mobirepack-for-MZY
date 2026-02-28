import sys
import os
import shutil
import subprocess
import zipfile
import re
import time
from PIL import Image, ImageStat

# 尝试导入核心解包引擎
try:
    from mobi.extract import extract as mobi_extract
except ImportError:
    print("================================================================")
    print("错误：缺少核心模块 'mobi'！请在终端运行：pip install mobi")
    print("================================================================")
    input("按回车键退出...")
    sys.exit(1)

# --- Kindle 2022 入门款 硬件配置 ---
KINDLE_WIDTH = 1072
KINDLE_HEIGHT = 1448

def is_blank_page(img):
    """
    终极去残算法 V12：融合阴影宽容、中心采样与 BFS 连续黑块侦测
    """
    try:
        if img.width < 10 or img.height < 10: 
            return True
            
        grayscale_img = img.convert('L')
        w, h = grayscale_img.width, grayscale_img.height
        total_pixels = w * h
        
        hist = grayscale_img.histogram()
        white_pixels = sum(hist[210:256])
        if (white_pixels / total_pixels) >= 0.70:
            return True

        center_crop = grayscale_img.crop((int(w*0.2), int(h*0.2), int(w*0.8), int(h*0.8)))
        center_stat = ImageStat.Stat(center_crop)
        if center_stat.stddev[0] < 8.0:
            return True

        thumb_w, thumb_h = 50, 50
        thumb = grayscale_img.resize((thumb_w, thumb_h), Image.Resampling.BILINEAR)
        pixels = thumb.load()
        
        visited = set()
        max_black_blob = 0
        
        for x in range(thumb_w):
            for y in range(thumb_h):
                if pixels[x, y] < 40 and (x, y) not in visited:
                    queue = [(x, y)]
                    visited.add((x, y))
                    head = 0
                    
                    while head < len(queue):
                        cx, cy = queue[head]
                        head += 1
                        
                        for dx, dy in [(-1,0), (1,0), (0,-1), (0,1)]:
                            nx, ny = cx + dx, cy + dy
                            if 0 <= nx < thumb_w and 0 <= ny < thumb_h:
                                if (nx, ny) not in visited and pixels[nx, ny] < 40:
                                    visited.add((nx, ny))
                                    queue.append((nx, ny))
                    
                    if len(queue) > max_black_blob:
                        max_black_blob = len(queue)
                        
        blob_ratio = max_black_blob / (thumb_w * thumb_h)
        if blob_ratio >= 0.40:
            return True

    except Exception as e:
        pass
    return False

def parse_opf_for_images_and_meta(extract_dir):
    """解析 OPF，还原阅读顺序，并提取原作者与出版社元数据"""
    print("正在破译 MOBI 排版说明书 (OPF/HTML)...")
    opf_file = None
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.lower().endswith('.opf'):
                opf_file = os.path.join(root, file)
                break
        if opf_file: break
        
    if not opf_file:
        print("错误：无法在解包数据中找到目录树文件。")
        return [], "", ""
        
    with open(opf_file, 'r', encoding='utf-8', errors='ignore') as f:
        opf_data = f.read()
        
    creator_match = re.search(r'<dc:creator[^>]*>(.*?)</dc:creator>', opf_data, re.IGNORECASE)
    author = creator_match.group(1).strip() if creator_match else ""
    
    publisher_match = re.search(r'<dc:publisher[^>]*>(.*?)</dc:publisher>', opf_data, re.IGNORECASE)
    publisher = publisher_match.group(1).strip() if publisher_match else ""

    manifest_items = {}
    for match in re.finditer(r'<item\s+[^>]*id=["\']([^"\']+)["\'][^>]*href=["\']([^"\']+)["\']', opf_data, re.IGNORECASE):
        manifest_items[match.group(1)] = match.group(2)
        
    ordered_images = []
    opf_base_dir = os.path.dirname(opf_file)
    
    for match in re.finditer(r'<itemref\s+[^>]*idref=["\']([^"\']+)["\']', opf_data, re.IGNORECASE):
        idref = match.group(1)
        if idref in manifest_items:
            html_rel_path = manifest_items[idref]
            html_full_path = os.path.normpath(os.path.join(opf_base_dir, html_rel_path))
            
            if os.path.exists(html_full_path):
                 with open(html_full_path, 'r', encoding='utf-8', errors='ignore') as f:
                     html_data = f.read()
                     for img_match in re.finditer(r'<img[^>]+src=["\']([^"\']+)["\']', html_data, re.IGNORECASE):
                         img_rel_path = img_match.group(1)
                         img_full_path = os.path.normpath(os.path.join(os.path.dirname(html_full_path), img_rel_path))
                         if os.path.exists(img_full_path) and img_full_path not in ordered_images:
                             ordered_images.append(img_full_path)
                                 
    print(f"成功破译！锁定 {len(ordered_images)} 张图片" + (f" | 作者: {author}" if author else ""))
    return ordered_images, author, publisher

def process_single_book(input_path, kindlegen_path, base_dir):
    target_dir = os.path.dirname(os.path.abspath(input_path))
    file_name = os.path.basename(input_path)
    name_without_ext = os.path.splitext(file_name)[0]
    output_name = f"{name_without_ext}_重构版.mobi"
    
    # --- 核心新增：自动创建并接管 remake 收纳文件夹 ---
    remake_dir = os.path.join(target_dir, "remake")
    if not os.path.exists(remake_dir):
        os.makedirs(remake_dir, exist_ok=True)
        print(f"  [+] 探测到新据点，已创建收纳目录: {remake_dir}")
    
    temp_dir = os.path.join(base_dir, f"temp_{int(time.time())}_{name_without_ext[:5]}")
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    print(f"\n>>> 开始处理: {file_name}")
    
    try:
        print("正在解包...")
        temp_extract_dir, _ = mobi_extract(input_path)
        
        try:
            ordered_images, author, publisher = parse_opf_for_images_and_meta(temp_extract_dir)
            if not ordered_images:
                print("[-] 错误：未能提取到有效图片序列。")
                return False

            print("正在深度清洗空白废页并适配分辨率...")
            Image.MAX_IMAGE_PIXELS = None 
            valid_images = []
            
            for img_path in ordered_images:
                try:
                    with Image.open(img_path) as img:
                        img = img.convert('L')
                        
                        if is_blank_page(img):
                            print(f"  [✂️ 拦截废页] {os.path.basename(img_path)}")
                            continue 

                        width_ratio = KINDLE_WIDTH / img.width
                        height_ratio = KINDLE_HEIGHT / img.height
                        scale = min(width_ratio, height_ratio)
                        new_width, new_height = int(img.width * scale), int(img.height * scale)
                        
                        img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        bg = Image.new('L', (KINDLE_WIDTH, KINDLE_HEIGHT), 255)
                        bg.paste(img_resized, ((KINDLE_WIDTH - new_width) // 2, (KINDLE_HEIGHT - new_height) // 2))
                        
                        out_filename = f"page_{len(valid_images):04d}.jpg"
                        bg.save(os.path.join(temp_dir, out_filename), 'JPEG', quality=90) 
                        valid_images.append(out_filename)
                except: pass

            if not valid_images: return False

            print(f"实际保留有效页数: {len(valid_images)}。正在生成元数据...")
            html_files = []
            for img_file in valid_images:
                html_name = img_file.replace('.jpg', '.html')
                html_files.append(html_name)
                with open(os.path.join(temp_dir, html_name), 'w', encoding='utf-8') as f:
                    f.write(f'<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml"><head><title>Page</title><meta name="viewport" content="width={KINDLE_WIDTH}, height={KINDLE_HEIGHT}" /><style>body, div, img {{ margin: 0; padding: 0; border: 0; }} body {{ width: {KINDLE_WIDTH}px; height: {KINDLE_HEIGHT}px; overflow: hidden; }} img {{ width: {KINDLE_WIDTH}px; height: {KINDLE_HEIGHT}px; display: block; }}</style></head><body><img src="{img_file}" alt="comic page"/></body></html>')

            book_title = name_without_ext
            author_tag = f'<dc:creator>{author}</dc:creator>' if author else ''
            publisher_tag = f'<dc:publisher>{publisher}</dc:publisher>' if publisher else ''
            
            opf_content = ['<?xml version="1.0" encoding="utf-8"?>',
                           '<package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="uid">',
                           f'<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">',
                           f'<dc:title>{book_title}</dc:title><dc:language>zh</dc:language>',
                           author_tag,
                           publisher_tag,
                           f'<meta name="fixed-layout" content="true"/><meta name="original-resolution" content="{KINDLE_WIDTH}x{KINDLE_HEIGHT}"/><meta name="book-type" content="comic"/><meta name="zero-margin" content="true"/></metadata><manifest>',
                           f'<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>']
            
            opf_content = [line for line in opf_content if line]
            
            for h, i in zip(html_files, valid_images):
                safe_id = h[:9].replace(".", "")
                opf_content.append(f'<item id="h_{safe_id}" href="{h}" media-type="application/xhtml+xml"/>')
                opf_content.append(f'<item id="i_{safe_id}" href="{i}" media-type="image/jpeg"/>')
            
            opf_content.append('</manifest><spine toc="ncx">')
            for h in html_files:
                safe_id = h[:9].replace(".", "")
                opf_content.append(f'<itemref idref="h_{safe_id}"/>')
            opf_content.append('</spine></package>')

            with open(os.path.join(temp_dir, "content.opf"), 'w', encoding='utf-8') as f: f.write("\n".join(opf_content))
            with open(os.path.join(temp_dir, "toc.ncx"), 'w', encoding='utf-8') as f:
                f.write('<?xml version="1.0" encoding="UTF-8"?><ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><navMap><navPoint id="p1" playOrder="1"><navLabel><text>Start</text></navLabel><content src="'+html_files[0]+'"/></navPoint></navMap></ncx>')

            print("正在调用 KindleGen 编译底层...")
            subprocess.run([kindlegen_path, os.path.join(temp_dir, "content.opf"), "-c1", "-o", output_name, "-dont_append_source"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # --- 核心修改：将最终成品移动到 remake 文件夹中 ---
            output_target = os.path.join(remake_dir, output_name)
            if os.path.exists(output_target): os.remove(output_target)
            shutil.move(os.path.join(temp_dir, output_name), output_target)
            print(f"[√] 成功收纳至: {os.path.join('remake', output_name)}")
            return True

        finally:
            if os.path.exists(temp_extract_dir): shutil.rmtree(temp_extract_dir, ignore_errors=True)

    except Exception as e:
        print(f"[X] 处理失败: {file_name} - {str(e)}")
        return False
    finally:
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir, ignore_errors=True)

def collect_files(inputs):
    targets = []
    for path in inputs:
        if os.path.isfile(path):
            if path.lower().endswith(('.mobi', '.azw3')):
                targets.append(path)
        elif os.path.isdir(path):
            print(f"正在扫描文件夹: {path} ...")
            for root, dirs, files in os.walk(path):
                for f in files:
                    if f.lower().endswith(('.mobi', '.azw3')):
                        targets.append(os.path.join(root, f))
    return targets

if __name__ == "__main__":
    try:
        if sys.platform.startswith('win'): os.system('chcp 65001 >nul')
        
        if getattr(sys, 'frozen', False) or '__compiled__' in globals():
            BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
        else:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        KINDLEGEN_EXE_PATH = os.path.join(BASE_DIR, "kindlegen.exe")

        if not os.path.exists(KINDLEGEN_EXE_PATH):
            print("---------------------------------------------------------")
            print(f"致命错误：找不到 kindlegen.exe")
            print(f"请将其放在目录: {BASE_DIR}")
            print("---------------------------------------------------------")
            input("按回车键退出..."); sys.exit(1)

        if len(sys.argv) < 2:
            print("用法：请将 文件、文件夹 拖拽到此程序上（支持多选）！")
            input("按回车键退出..."); sys.exit(1)

        all_books = collect_files(sys.argv[1:])
        
        if not all_books:
            print("错误：未找到任何 .mobi 或 .azw3 文件。")
            input("按回车键退出..."); sys.exit(1)

        print(f"==========================================")
        print(f"共发现 {len(all_books)} 本书，准备开始批量处理...")
        print(f"==========================================")

        success_count = 0
        fail_count = 0

        for index, book_path in enumerate(all_books):
            print(f"\n[{index+1}/{len(all_books)}] 正在处理...")
            if process_single_book(book_path, KINDLEGEN_EXE_PATH, BASE_DIR):
                success_count += 1
            else:
                fail_count += 1

        print(f"\n==========================================")
        print(f"全部任务结束！")
        print(f"成功: {success_count} 本 | 失败: {fail_count} 本")
        print(f"==========================================")
        
        input("按回车键退出...")

    except Exception as e:
        import traceback
        print(f"\n[严重崩溃]:\n{traceback.format_exc()}")
        input("按回车键退出...")
