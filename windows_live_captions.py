import cv2
import numpy as np
import pyautogui
import pytesseract
import os
import sys
from googletrans import Translator
import time
import logging
import asyncio
from concurrent.futures import ThreadPoolExecutor

# 设置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 初始化翻译缓存
translation_cache = {}

def check_tesseract():
    try:
        pytesseract.get_tesseract_version()
        print("Tesseract OCR 已正确安装。")
    except pytesseract.TesseractNotFoundError:
        print("错误：未找到 Tesseract OCR。")
        print("请确保已安装 Tesseract OCR 并将其添加到系统 PATH 中。")
        print("安装说明：")
        print("1. 访问 https://github.com/UB-Mannheim/tesseract/wiki 下载并安装 Tesseract")
        print("2. 将安装路径（如 C:\\Program Files\\Tesseract-OCR）添加到系统 PATH 中")
        print("3. 确保已安装日语语言包")
        print("4. 重启程序")
        sys.exit(1)

def get_caption_area():
    print("请按照以下步骤指定字幕区域：")
    print("1. 将鼠标移动到字幕区域的左上角")
    print("2. 按下回车键")
    input("准备好后按回车...")
    x1, y1 = pyautogui.position()
    print(f"左上角坐标：({x1}, {y1})")
    
    print("现在将鼠标移动到字幕区域的右下角")
    input("准备好后按回车...")
    x2, y2 = pyautogui.position()
    print(f"右下角坐标：({x2}, {y2})")
    
    return x1, y1, x2 - x1, y2 - y1

async def translate_text(text, src='ja', dest='zh-cn', max_retries=3):
    if text in translation_cache:
        return translation_cache[text]
    
    translator = Translator()
    for attempt in range(max_retries):
        try:
            result = await asyncio.to_thread(translator.translate, text, src=src, dest=dest)
            translation_cache[text] = result.text
            return result.text
        except Exception as e:
            if attempt < max_retries - 1:
                logging.warning(f"翻译失败，正在重试... (尝试 {attempt + 1}/{max_retries})")
                await asyncio.sleep(1)
            else:
                logging.error(f"翻译失败: {e}")
                return None

def split_sentences(text):
    # 简单的句子分割，可以根据日语的特点进行优化
    return text.replace('。', '。\n').split('\n')

def preprocess_image(image, options):
    if options['grayscale']:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    if options['denoise']:
        image = cv2.fastNlMeansDenoising(image)
    if options['threshold']:
        _, image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    if options['deskew']:
        coords = np.column_stack(np.where(image > 0))
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle
        (h, w) = image.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        image = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return image

async def capture_and_process_captions(x, y, width, height, last_sentences, last_translation, preprocess_options, ocr_lang, translate_src, translate_dest):
    try:
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        frame = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        # 图像预处理
        frame = preprocess_image(frame, preprocess_options)
        
        # 使用自定义配置
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(frame, config=custom_config, lang=ocr_lang)
        
        new_original = ""
        new_translation = ""
        
        if text.strip():
            new_sentences = split_sentences(text)
            
            # 找出新的句子
            if last_sentences:
                for sentence in new_sentences:
                    if sentence not in last_sentences:
                        new_original += sentence + " "
            else:
                new_original = ' '.join(new_sentences)
            
            if new_original.strip():
                new_translation = await translate_text(new_original, src=translate_src, dest=translate_dest)
                
                if new_translation:
                    logging.info(f"新增原文: {new_original}")
                    logging.info(f"新增翻译: {new_translation}")
                    logging.info("-" * 50)
                    
                    # 更新last_sentences和last_translation
                    last_sentences = (last_sentences + new_sentences)[-5:]
                    last_translation = last_translation + ' ' + new_translation if last_translation else new_translation
        
        return last_sentences, last_translation, new_original, new_translation
        
    except Exception as e:
        logging.error(f"发生错误: {e}")
        raise  # 将异常传播到上层
