import os
import csv
from datetime import datetime
import shutil
import imaging
import ocr

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CROPPED_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'temp', 'cropped')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'output')
DEBUG_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'debug')
OUTPUT_CSV_PATH = os.path.join(OUTPUT_DIR, 'race_data.csv')

# --- 設定 ---
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
MAX_VALID_RATE = 10000


def get_last_race_rate(csv_path):
    """CSVファイルから最後のレースの最終レートを取得する"""
    if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
        return None
    
    with open(csv_path, 'r', newline='', encoding='utf-8-sig') as f:
        reader = list(csv.reader(f))
        if len(reader) < 2: # ヘッダーのみ、または空の場合
            return None
        
        header = reader[0]
        last_log = reader[-1]
        try:
            rate_idx = header.index('Rate')
            return int(last_log[rate_idx])
        except (ValueError, IndexError):
            # ヘッダーや列が見つからない場合
            return None

def get_course_and_pre_race_rate(image_path, previous_course_name=None):
    """
    コース決定画面の画像から、コース名、ユーザーの開始前レート、参加人数を取得する。
    """
    os.makedirs(CROPPED_DIR, exist_ok=True)
    
    rate_path, course_path, participant_count = imaging.analyze_course_decision_screen(image_path)
    
    pre_race_rate = None
    if rate_path:
        pre_race_rate = ocr.analyze_rate_ocr(rate_path, TESSERACT_PATH)
        
    course_name = "コース不明"
    if course_path:
        current_course_name = ocr.analyze_course_ocr(course_path, TESSERACT_PATH)
        if current_course_name != "コース不明" and previous_course_name:
            course_name = f"{previous_course_name} → {current_course_name}"
        else:
            course_name = current_course_name
    
    return pre_race_rate, course_name, participant_count

def process_result_image(image_path, course_name, pre_race_rate, participant_count, is_debug_mode=False):
    """リザルト画面を解析し、レート差分と参加人数をCSVに保存する"""
    if not os.path.exists(image_path): return None
    
    os.makedirs(CROPPED_DIR, exist_ok=True); os.makedirs(OUTPUT_DIR, exist_ok=True); os.makedirs(DEBUG_DIR, exist_ok=True)

    detected_pos = imaging.crop_image_for_result(image_path)
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    final_result = None

    if detected_pos:
        rank_path = os.path.join(CROPPED_DIR, f"{base_filename}_rank.png")
        rate_path = os.path.join(CROPPED_DIR, f"{base_filename}_rate.png")
        race_points_path = os.path.join(CROPPED_DIR, f"{base_filename}_rate_change.png")
        
        final_rank = None
        if 1 <= detected_pos <= 12: final_rank = detected_pos
        elif detected_pos == 13: final_rank = ocr.analyze_rank_ocr(rank_path, TESSERACT_PATH)
        
        final_rate = ocr.analyze_rate_ocr(rate_path, TESSERACT_PATH)
        
        if final_rate is not None and final_rate > MAX_VALID_RATE:
            print(f"[analysis] WARNING: 異常なレート値({final_rate})を検出したため、この結果を破棄します。")
            if is_debug_mode:
                debug_subfolder = os.path.join(DEBUG_DIR, f"error_{base_filename}")
                os.makedirs(debug_subfolder, exist_ok=True)
                try: shutil.copy(image_path, os.path.join(debug_subfolder, os.path.basename(image_path)))
                except Exception as e: print(f"[analysis] DEBUG WARNING: 異常画像のコピーに失敗: {e}")
            return None

        race_points = ocr.analyze_rate_change_ocr(race_points_path, TESSERACT_PATH)
        
        # --- ★レート変動計算のロジックを刷新★ ---
        net_rate_change = 0
        baseline_rate = None

        # 優先度1: コース決定画面で取得したレート
        if pre_race_rate is not None and pre_race_rate > 0:
            baseline_rate = pre_race_rate
            print(f"[analysis] INFO: 開始前レート ({pre_race_rate}) を使用して変動を計算します。")
        # 優先度2: 前回のレース結果の最終レート
        else:
            last_final_rate = get_last_race_rate(OUTPUT_CSV_PATH)
            if last_final_rate is not None:
                baseline_rate = last_final_rate
                print(f"[analysis] WARNING: 開始前レートが不明です。前回の最終レート ({last_final_rate}) を参照します。")
            else:
                # 最初の記録
                print("[analysis] INFO: 初回記録のため、レート変動は0になります。")
                baseline_rate = final_rate # 差分が0になるように設定

        if final_rate is not None and baseline_rate is not None:
            net_rate_change = final_rate - baseline_rate

        if final_rank and final_rate is not None:
            timestamp_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            final_result = [image_path, timestamp_str, course_name, final_rank, participant_count, final_rate, net_rate_change, race_points]
            
            is_first_record = not os.path.exists(OUTPUT_CSV_PATH) or os.path.getsize(OUTPUT_CSV_PATH) == 0
            with open(OUTPUT_CSV_PATH, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                if is_first_record: writer.writerow(['Filename', 'Timestamp', 'Course', 'Rank', 'Participants', 'Rate', 'Rate Change', 'Points'])
                writer.writerow([os.path.basename(image_path), timestamp_str, course_name, final_rank, participant_count, final_rate, net_rate_change, race_points])
            print(f"[analysis] SUCCESS: 結果をCSVに保存しました -> Course:{course_name}, Rank:{final_rank}/{participant_count}, Rate:{final_rate}, Change:{net_rate_change:+}")
    else:
        print(f"[analysis] WARNING: '{base_filename}'からハイライトが見つからなかったため、解析をスキップしました。")

    if is_debug_mode:
        debug_subfolder = os.path.join(DEBUG_DIR, base_filename)
        os.makedirs(debug_subfolder, exist_ok=True)
        print(f"[analysis] DEBUG: デバッグファイルを '{debug_subfolder}' に保存します。")
        try: shutil.move(image_path, os.path.join(debug_subfolder, os.path.basename(image_path)))
        except (FileNotFoundError, PermissionError) as e: print(f"[analysis] DEBUG WARNING: 元ファイルの移動に失敗しました: {e}")
        for region_name in ['rank', 'rate', 'rate_change', 'prerace_rate', 'course_name']:
            cropped_path = os.path.join(CROPPED_DIR, f"{base_filename}_{region_name}.png")
            if os.path.exists(cropped_path):
                try: shutil.move(cropped_path, os.path.join(debug_subfolder, f"{base_filename}_{region_name}.png"))
                except (FileNotFoundError, PermissionError) as e: print(f"[analysis] DEBUG WARNING: 切り抜きファイルの移動に失敗しました: {e}")

    return final_result

