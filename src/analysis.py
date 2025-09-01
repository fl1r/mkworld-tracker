import os
import csv
from datetime import datetime
import shutil
from difflib import SequenceMatcher

import imaging
import ocr
import config

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CROPPED_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'temp', 'cropped')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'output')
DEBUG_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'debug')
OUTPUT_CSV_PATH = os.path.join(OUTPUT_DIR, 'race_data.csv')

# --- 設定 ---
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
MAX_VALID_RATE = 10000


def find_closest_course_name(ocr_text, course_list):
    """
    OCRで読み取ったテキストと、既知のコース名リストを比較し、
    最も類似度が高いコース名を返す。
    """
    if not ocr_text or ocr_text == "コース不明":
        return "コース不明"
    
    best_match = "コース不明"
    best_score = 0.6
    
    for course in course_list:
        score = SequenceMatcher(None, ocr_text, course).ratio()
        if score > best_score:
            best_score = score
            best_match = course
            
    if best_match != "コース不明":
        print(f"[analysis] Matched: '{ocr_text}' => '{best_match}' (score: {best_score:.2f})")
    
    return best_match

def get_last_race_rate(csv_path):
    """CSVファイルから最後のレースの最終レートを取得する。"""
    if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
        return None
    with open(csv_path, 'r', newline='', encoding='utf-8-sig') as f:
        reader = list(csv.reader(f))
        if len(reader) < 2: return None
        header = reader[0]; last_log = reader[-1]
        try:
            rate_idx = header.index('Rate')
            return int(last_log[rate_idx])
        except (ValueError, IndexError): return None

def get_last_race_course(csv_path):
    """CSVファイルから最後のレースのコース名を取得する"""
    if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
        return None
    with open(csv_path, 'r', newline='', encoding='utf-8-sig') as f:
        reader = list(csv.reader(f))
        if len(reader) < 2: return None
        header = reader[0]
        last_log = reader[-1]
        try:
            course_idx = header.index('Course')
            return last_log[course_idx]
        except (ValueError, IndexError):
            return None

def get_course_and_pre_race_rate(image_path):
    """
    コース決定画面から、レート、コース名、参加人数を取得する。
    2連続レースの場合、CSVの最後のコースを始点とする。
    """
    os.makedirs(CROPPED_DIR, exist_ok=True)
    
    rate_path, course_path, participant_count, is_single_course = imaging.analyze_course_decision_screen(image_path)
    
    pre_race_rate = None
    if rate_path:
        pre_race_rate = ocr.analyze_rate_ocr(rate_path, TESSERACT_PATH)
    
    raw_course_name = "コース不明"
    if course_path:
        raw_course_name = ocr.analyze_course_ocr(course_path)
        
    corrected_course_name = find_closest_course_name(raw_course_name, config.COURSE_NAMES)
    
    final_course_name = "コース不明"
    if corrected_course_name != "コース不明":
        if is_single_course:
            final_course_name = corrected_course_name
        else:
            start_point = get_last_race_course(OUTPUT_CSV_PATH)
            if start_point is None:
                start_point = "不明"
            
            end_point = corrected_course_name
            final_course_name = f"{start_point} → {end_point}"

    return pre_race_rate, final_course_name, participant_count

def process_result_image(image_path, course_name, pre_race_rate, participant_count, is_debug_mode=False):
    """リザルト画面の画像を解析し、最終的なレース結果をCSVに保存する。"""
    if not os.path.exists(image_path): return None
    
    os.makedirs(CROPPED_DIR, exist_ok=True); os.makedirs(OUTPUT_DIR, exist_ok=True); os.makedirs(DEBUG_DIR, exist_ok=True)

    detected_pos = imaging.crop_image_for_result(image_path)
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    final_result = None

    if detected_pos:
        rank_path = os.path.join(CROPPED_DIR, f"{base_filename}_rank.png")
        rate_path = os.path.join(CROPPED_DIR, f"{base_filename}_rate.png")
        race_points_path = os.path.join(CROPPED_DIR, f"{base_filename}_rate_change.png")
        
        final_rank = detected_pos
        if detected_pos >= 13: 
             final_rank = ocr.analyze_rank_ocr(rank_path, TESSERACT_PATH)
        
        final_rate = ocr.analyze_rate_ocr(rate_path, TESSERACT_PATH)
        race_points = ocr.analyze_rate_change_ocr(race_points_path, TESSERACT_PATH)
        
        if final_rate is not None and final_rate > MAX_VALID_RATE:
            print(f"[analysis] WARNING: 異常なレート値({final_rate})を検出したため、この結果を破棄します。")
            return None
            
        net_rate_change = 0
        if final_rate is not None:
            if pre_race_rate is not None and pre_race_rate > 0:
                net_rate_change = final_rate - pre_race_rate
            else:
                last_final_rate = get_last_race_rate(OUTPUT_CSV_PATH)
                if last_final_rate is not None:
                    net_rate_change = final_rate - last_final_rate

        if final_rank and final_rate is not None:
            timestamp_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            is_first_record = not os.path.exists(OUTPUT_CSV_PATH) or os.path.getsize(OUTPUT_CSV_PATH) == 0
            if is_first_record:
                net_rate_change = 0

            with open(OUTPUT_CSV_PATH, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                if is_first_record:
                    writer.writerow(['Filename', 'Timestamp', 'Course', 'Rank', 'Participants', 'Rate', 'Rate Change'])
                row_data = [os.path.basename(image_path), timestamp_str, course_name, final_rank, participant_count, final_rate, net_rate_change]
                writer.writerow(row_data)
            
            print(f"[analysis] SUCCESS: 結果をCSVに保存しました -> Course:{course_name}, Rank:{final_rank}/{participant_count}, Rate:{final_rate}, Change:{net_rate_change:+}")
            final_result = row_data
    else:
        print(f"[analysis] WARNING: '{base_filename}'からハイライトが見つからなかったため、解析をスキップしました。")

    if is_debug_mode:
        debug_subfolder = os.path.join(DEBUG_DIR, base_filename)
        os.makedirs(debug_subfolder, exist_ok=True)
        print(f"[analysis] DEBUG: デバッグファイルを '{debug_subfolder}' に保存します。")
        try:
            shutil.move(image_path, os.path.join(debug_subfolder, os.path.basename(image_path)))
        except (FileNotFoundError, PermissionError) as e:
            print(f"[analysis] DEBUG WARNING: 元ファイルの移動に失敗しました: {e}")
        for region_name in ['rank', 'rate', 'rate_change', 'prerace_rate', 'course_name', 'course_gemini_input']:
            cropped_path = os.path.join(CROPPED_DIR, f"{base_filename}_{region_name}.png")
            if os.path.exists(cropped_path):
                try:
                    shutil.move(cropped_path, os.path.join(debug_subfolder, f"{base_filename}_{region_name}.png"))
                except (FileNotFoundError, PermissionError) as e:
                    print(f"[analysis] DEBUG WARNING: 切り抜きファイルの移動に失敗しました: {e}")

    return final_result