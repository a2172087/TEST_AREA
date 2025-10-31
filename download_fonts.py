"""
Font Downloader for AutoZ Wafer4P Aligner
下載 Font Awesome 和 Google Fonts 到本地網路路徑

用途：
此腳本會下載所需的字體資源到網路共享位置，
使 AutoZ Wafer4P Aligner 可以在內網環境中使用字體。

目標路徑：
- M:\BI_Database\Apps\Database\Apps_Database\RD_All\AutoZ Wafer4P Aligner\Font_Awesome
- M:\BI_Database\Apps\Database\Apps_Database\RD_All\AutoZ Wafer4P Aligner\Google_Fonts

使用方法：
    python download_fonts.py
"""

import os
import requests
import zipfile
import shutil
from pathlib import Path

# 目標路徑
FONT_AWESOME_PATH = r"M:\BI_Database\Apps\Database\Apps_Database\RD_All\AutoZ Wafer4P Aligner\Font_Awesome"
GOOGLE_FONTS_PATH = r"M:\BI_Database\Apps\Database\Apps_Database\RD_All\AutoZ Wafer4P Aligner\Google_Fonts"

# Font Awesome 下載網址（使用免費版）
FONT_AWESOME_URL = "https://use.fontawesome.com/releases/v6.4.0/fontawesome-free-6.4.0-web.zip"

# Google Fonts Noto Sans TC CSS
NOTO_SANS_TC_CSS_URL = "https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap"


def download_file(url, local_path):
    """下載檔案"""
    print(f"正在下載: {url}")
    response = requests.get(url, stream=True)
    response.raise_for_status()

    with open(local_path, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)

    print(f"下載完成: {local_path}")


def extract_zip(zip_path, extract_to):
    """解壓縮 ZIP 檔案"""
    print(f"正在解壓縮: {zip_path}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"解壓縮完成: {extract_to}")


def download_font_awesome():
    """下載 Font Awesome"""
    print("\n" + "="*60)
    print("開始下載 Font Awesome 6.4.0")
    print("="*60)

    # 創建目標目錄
    os.makedirs(FONT_AWESOME_PATH, exist_ok=True)

    # 下載 ZIP 檔案
    temp_zip = os.path.join(FONT_AWESOME_PATH, "fontawesome.zip")
    download_file(FONT_AWESOME_URL, temp_zip)

    # 解壓縮
    temp_extract = os.path.join(FONT_AWESOME_PATH, "temp")
    extract_zip(temp_zip, temp_extract)

    # 移動檔案到正確位置
    extracted_folder = os.path.join(temp_extract, "fontawesome-free-6.4.0-web")

    # 複製 css 和 webfonts 資料夾
    for folder in ['css', 'webfonts']:
        src = os.path.join(extracted_folder, folder)
        dst = os.path.join(FONT_AWESOME_PATH, folder)

        if os.path.exists(dst):
            shutil.rmtree(dst)

        shutil.copytree(src, dst)
        print(f"已複製: {folder} -> {dst}")

    # 清理臨時檔案
    os.remove(temp_zip)
    shutil.rmtree(temp_extract)

    print("\n✅ Font Awesome 下載完成!")
    print(f"   路徑: {FONT_AWESOME_PATH}")


def download_google_fonts():
    """下載 Google Fonts - Noto Sans TC"""
    print("\n" + "="*60)
    print("開始下載 Google Fonts - Noto Sans TC")
    print("="*60)

    # 創建目標目錄
    css_dir = os.path.join(GOOGLE_FONTS_PATH, "css")
    fonts_dir = os.path.join(GOOGLE_FONTS_PATH, "fonts")
    os.makedirs(css_dir, exist_ok=True)
    os.makedirs(fonts_dir, exist_ok=True)

    # 下載 CSS
    print("正在下載 CSS...")
    response = requests.get(NOTO_SANS_TC_CSS_URL)
    response.raise_for_status()

    css_content = response.text

    # 解析 CSS 中的字體 URL
    import re
    font_urls = re.findall(r'url\((https://[^)]+)\)', css_content)

    print(f"找到 {len(font_urls)} 個字體檔案")

    # 下載每個字體檔案並修改 CSS
    for i, font_url in enumerate(font_urls):
        # 取得檔案名稱
        font_filename = f"noto-sans-tc-{i}.woff2"
        font_path = os.path.join(fonts_dir, font_filename)

        # 下載字體檔案
        print(f"  下載字體 {i+1}/{len(font_urls)}: {font_filename}")
        font_response = requests.get(font_url)
        font_response.raise_for_status()

        with open(font_path, 'wb') as f:
            f.write(font_response.content)

        # 修改 CSS 中的 URL 為相對路徑
        css_content = css_content.replace(font_url, f"../fonts/{font_filename}")

    # 儲存修改後的 CSS
    css_file_path = os.path.join(css_dir, "noto-sans-tc.css")
    with open(css_file_path, 'w', encoding='utf-8') as f:
        f.write(css_content)

    print(f"\n已儲存 CSS: {css_file_path}")
    print(f"已儲存字體檔案到: {fonts_dir}")

    print("\n✅ Google Fonts 下載完成!")
    print(f"   路徑: {GOOGLE_FONTS_PATH}")


def main():
    """主程式"""
    print("\n" + "="*60)
    print("Font Downloader for AutoZ Wafer4P Aligner")
    print("="*60)

    try:
        # 下載 Font Awesome
        download_font_awesome()

        # 下載 Google Fonts
        download_google_fonts()

        print("\n" + "="*60)
        print("✅ 所有字體下載完成!")
        print("="*60)
        print(f"\nFont Awesome: {FONT_AWESOME_PATH}")
        print(f"Google Fonts: {GOOGLE_FONTS_PATH}")
        print("\n現在您可以在內網環境中使用 AutoZ Wafer4P Aligner 了。")

    except Exception as e:
        print(f"\n❌ 錯誤: {str(e)}")
        print("\n請檢查:")
        print("1. 網路連線是否正常")
        print("2. 目標路徑是否可存取")
        print("3. 是否有足夠的權限")
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
