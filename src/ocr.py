import cv2
import pytesseract
import os
import re
import numpy as np
from PIL import Image
import google.generativeai as genai
import config
import configparser
import tkinter as tk
from tkinter import simpledialog

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PRIVATE_CONFIG_PATH = os.path.join(SCRIPT_DIR, 'private_config.ini')

def load_or_prompt_api_key():
    """private_config.iniからAPIキーを読み込む。なければユーザーに尋ねて保存する。"""
    config_parser = configparser.ConfigParser()
    
    # ファイルからキーを読み込む試み
    if os.path.exists(PRIVATE_CONFIG_PATH):
        config_parser.read(PRIVATE_CONFIG_PATH)
        if 'Gemini' in config_parser and 'api_key' in config_parser['Gemini']:
            key = config_parser['Gemini']['api_key']
            if key: # キーが空でなければ返す
                return key

    # ファイルまたはキーが存在しない場合、ダイアログでユーザーに尋ねる
    root = tk.Tk()
    root.withdraw() # メインウィンドウを非表示にする
    api_key = simpledialog.askstring("Gemini API Key", "Google AI Studioから取得したAPIキーを入力してください:", show='*')
    root.destroy()

    if api_key:
        # 入力されたキーをファイルに保存
        config_parser['Gemini'] = {'api_key': api_key}
        with open(PRIVATE_CONFIG_PATH, 'w') as f:
            config_parser.write(f)
        return api_key
    else:
        # 入力がなければNoneを返す
        return None

# --- Gemini APIの初期設定 ---
API_KEY = load_or_prompt_api_key()

if API_KEY:
    try:
        genai.configure(api_key=API_KEY)
    except Exception as e:
        print(f"[ocr] ERROR: Gemini APIキーの設定に失敗しました: {e}")
        API_KEY = None # 設定に失敗したらキーをNoneに戻す

# (analyze_rank_ocr などの他の関数は変更なし)
def analyze_rank_ocr(image_path, tesseract_path=None):
    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    
    roi = cv2.imread(image_path)
    if roi is None: return None
    
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    config_str = f'--oem 1 --psm 7 -c tessedit_char_whitelist="0123456789"'
    text = pytesseract.image_to_string(binary, lang='eng', config=config_str).strip()
    
    print(f"[ocr] DEBUG: OCR Raw Text ('{os.path.basename(image_path)}') = '{text}'")
    return int(text) if text and text.isdigit() else None

def analyze_rate_ocr(image_path, tesseract_path=None):
    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    roi = cv2.imread(image_path)
    if roi is None: return None
    
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    config_str = f'--oem 1 --psm 7 -c tessedit_char_whitelist="0123456789"'
    text = pytesseract.image_to_string(binary, lang='eng', config=config_str).strip()
    
    print(f"[ocr] DEBUG: OCR Raw Text ('{os.path.basename(image_path)}') = '{text}'")
    return int(text) if text and text.isdigit() else None

def analyze_rate_change_ocr(image_path, tesseract_path=None):
    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    roi = cv2.imread(image_path)
    if roi is None: return 0
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    
    inverted = cv2.bitwise_not(gray)
    _, binary = cv2.threshold(inverted, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    config_str = f'--oem 1 --psm 7 -c tessedit_char_whitelist="+-0123456789"'
    text = pytesseract.image_to_string(binary, lang='eng', config=config_str).strip()
    
    print(f"[ocr] DEBUG: OCR Raw Text ('{os.path.basename(image_path)}') = '{text}'")
    if text and re.match(r'^[+\-]\d+$', text): return int(text)
    return 0 

def analyze_course_ocr(image_path):
    if not API_KEY:
        print("[ocr] WARNING: Gemini APIキーが設定されていません。コース名認識をスキップします。")
        return "コース不明 (APIキー未設定)"
        
    if not os.path.exists(image_path):
        print(f"[ocr] ERROR: 画像ファイルが見つかりません: {image_path}")
        return "コース不明 (ファイルなし)"

    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        img = Image.open(image_path)
        
        course_list_str = ", ".join(config.COURSE_NAMES)

        prompt = (
            "これはレースゲームのコース選択画面の画像です。"
            "以下の『コース名リスト』の中から、画像に表示されているコース名を**一つだけ**正確に選び出し、そのテキストだけを返してください。"
            "リストにない名前は絶対に回答しないでください。\n\n"
            "--- コース名リスト ---\n"
            f"{course_list_str}"
        )
        
        response = model.generate_content([prompt, img])
        text = response.text.strip().replace(" ", "").replace("\n", "")
        
        print(f"[ocr] DEBUG: Gemini API Raw Text ('{os.path.basename(image_path)}') = '{text}'")
        
        return text if text else "コース不明"

    except Exception as e:
        print(f"[ocr] ERROR: Gemini APIの呼び出し中にエラーが発生しました: {e}")
        return "コース不明 (APIエラー)"