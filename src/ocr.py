import cv2
import pytesseract
import os
import re
import numpy as np
from difflib import SequenceMatcher # 文字列の類似度を計算するためにインポート

# --- コース名辞書とホワイトリストの作成 ---
COURSE_NAMES = [
    "マリオブラザーズサーキット", "トロフィーシティ", "シュポポコースター", "DKうちゅうセンター",
    "サンサンさばく", "ヘイホーカーニバル", "ワリオスタジアム", "キラーシップ", "DKスノーマウンテン",
    "ロゼッタてんもんだい", "アイスビルディング", "ワリオシップ", "ノコノコビーチ",
    "リバーサイドサファリ", "ピーチスタジアム", "バナナカップピーチビーチ", "ソルティータウン",
    "ディノディノジャングル", "ハテナしんでん", "プクプクフォールズ", "ショーニューロード", "おばけシネマ",
    "ホネホネツイスター", "モーモーカントリー", "チョコマウンテン", "キノピオファクトリー",
    "クッパキャッスル", "どんぐりツリーハウス", "マリオサーキット", "レインボーロード"
]

all_chars = set()
for name in COURSE_NAMES:
    all_chars.update(list(name))
COURSE_WHITELIST = "".join(sorted(list(all_chars)))

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
USER_WORDS_PATH = os.path.join(SCRIPT_DIR, 'user-words.txt')
with open(USER_WORDS_PATH, 'w', encoding='utf-8') as f:
    for name in set(COURSE_NAMES):
        f.write(name + '\n')


def analyze_rank_ocr(image_path, tesseract_path=None):
    """Analyzes the rank (numbers) using OCR"""
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
    """Analyzes the rate (numbers) using OCR with contour correction"""
    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    roi = cv2.imread(image_path)
    if roi is None: return None
    
    height, width, _ = roi.shape
    upscaled = cv2.resize(roi, (width * 5, height * 5), interpolation=cv2.INTER_LANCZOS4)
    gray = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
    
    inverted = cv2.bitwise_not(gray)
    kernel = np.ones((2, 2), np.uint8)
    opening = cv2.morphologyEx(inverted, cv2.MORPH_OPEN, kernel, iterations=1)
    final_img = cv2.bitwise_not(opening)

    config_str = f'--oem 1 --psm 7 -c tessedit_char_whitelist="0123456789"'
    text = pytesseract.image_to_string(final_img, lang='eng', config=config_str).strip()
    
    print(f"[ocr] DEBUG: OCR Raw Text ('{os.path.basename(image_path)}') = '{text}'")
    return int(text) if text and text.isdigit() else None

def analyze_rate_change_ocr(image_path, tesseract_path=None):
    """Analyzes the rate change (+/- and numbers) using OCR"""
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

def find_closest_course_name(ocr_text, course_list):
    """
    Finds the closest course name from a list based on the OCR result.
    """
    if not ocr_text: return "コース不明"
    
    best_match = None
    best_score = 0
    
    for course in course_list:
        score = SequenceMatcher(None, ocr_text, course).ratio()
        if score > best_score:
            best_score = score
            best_match = course
    
    # If similarity is above 60%, consider it a match.
    if best_score > 0.6:
        print(f"[ocr] Matched: '{ocr_text}' => '{best_match}' (score: {best_score:.2f})")
        return best_match
    
    print(f"[ocr] No confident match found for '{ocr_text}'. Best was '{best_match}' with score {best_score:.2f}.")
    return "コース不明"

def analyze_course_ocr(image_path, tesseract_path=None):
    """
    Analyzes the course name using an improved OCR pipeline.
    """
    if not os.path.exists(image_path): return "コース不明"
    
    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    roi = cv2.imread(image_path)
    if roi is None: return "コース不明"

    # --- Improved Preprocessing Pipeline ---
    height, width, _ = roi.shape
    
    # 1. Upscale for better detail
    upscaled = cv2.resize(roi, (width * 3, height * 3), interpolation=cv2.INTER_CUBIC)
    
    # 2. Convert to grayscale
    gray = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
    
    # 3. Use adaptive thresholding to handle varying local contrast
    binary = cv2.adaptiveThreshold(
        gray, 255, 
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 
        35, 15
    )
    
    # 4. Invert colors if the original text is white
    if np.mean(gray) > 127:
        binary = cv2.bitwise_not(binary)
    
    # 5. Use morphological closing to connect broken parts of characters
    kernel = np.ones((2,2), np.uint8)
    final_image = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel, iterations=1)
    
    debug_steps_dir = os.path.join(os.path.dirname(image_path), "debug_steps")
    os.makedirs(debug_steps_dir, exist_ok=True)
    cv2.imwrite(os.path.join(debug_steps_dir, f"{os.path.basename(image_path)}_final_ocr_input.png"), final_image)
    
    # Execute OCR with improved configuration
    config_str = f'--oem 1 --psm 7 --user-words "{USER_WORDS_PATH}" -c tessedit_char_whitelist="{COURSE_WHITELIST}"'
    text = pytesseract.image_to_string(final_image, lang='jpn', config=config_str).strip()
    
    print(f"[ocr] DEBUG: Course OCR Raw Text ('{os.path.basename(image_path)}') = '{text}'")
    
    # Correct the OCR result using the dictionary
    return find_closest_course_name(text.replace("\n", "").replace(" ", ""), COURSE_NAMES)

