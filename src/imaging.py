import cv2
import numpy as np
import os
import config

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'temp', 'cropped')
DEBUG_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'debug')

def crop_image_for_result(image_path):
    """リザルト画面の画像を解析し、ハイライト位置と関連領域を切り抜く"""
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    img = cv2.imread(image_path)
    if img is None: return None
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    lower_highlight = np.array(config.LOWER_HIGHLIGHT)
    upper_highlight = np.array(config.UPPER_HIGHLIGHT)
    player_rank_found = None
    for i in range(1, 14):
        x1, y1, x2, y2 = config.RESULT_COORDS[f'rank_{i}']
        if y2 > img.shape[0] or x2 > img.shape[1]: continue
        check_roi = hsv[y1:y2, x1:x2]
        mask = cv2.inRange(check_roi, lower_highlight, upper_highlight)
        if cv2.countNonZero(mask) > ((x2 - x1) * (y2 - y1) * 0.2):
            player_rank_found = i; break
    if player_rank_found:
        regions_to_crop = {
            'rank': config.RESULT_COORDS[f'rank_{player_rank_found}'],
            'rate': config.RESULT_COORDS[f'rate_{player_rank_found}'],
            'rate_change': config.RESULT_COORDS[f'rate_change_{player_rank_found}'],
        }
        for region_name, (x1, y1, x2, y2) in regions_to_crop.items():
            if y2 > img.shape[0] or x2 > img.shape[1]: continue
            roi = img[y1:y2, x1:x2]
            save_path = os.path.join(OUTPUT_DIR, f"{base_filename}_{region_name}.png")
            cv2.imwrite(save_path, roi)
        return player_rank_found
    return None

def analyze_course_decision_screen(image_path):
    """
    コース決定画面を解析し、レート・コース・参加人数に加え、
    単独レースかどうかも判別する。
    """
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    img = cv2.imread(image_path)
    if img is None:
        return None, None, 0, False

    # --- 参加人数とプレイヤーのレート特定処理 (変更なし) ---
    brightest_rate_roi = None
    max_brightness = 0
    participant_count = 0
    GRAY_THRESHOLD = 50

    for coords in config.ALL_PLAYER_SLOTS:
        x1, y1, x2, y2 = coords['x1'], coords['y1'], coords['x2'], coords['y2']
        if y2 > img.shape[0] or x2 > img.shape[1]:
            continue
        roi = img[y1:y2, x1:x2]
        gray_roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        
        if np.mean(gray_roi) > GRAY_THRESHOLD:
            participant_count += 1
            brightness = np.max(gray_roi)
            if brightness > max_brightness:
                max_brightness = brightness
                brightest_rate_roi = roi
    
    rate_save_path = None
    if brightest_rate_roi is not None:
        rate_save_path = os.path.join(OUTPUT_DIR, f"{base_filename}_prerace_rate.png")
        cv2.imwrite(rate_save_path, brightest_rate_roi)

    course_gemini_input_path = None
    search_coords = config.COURSE_SEARCH_AREA
    ax1, ay1, ax2, ay2 = search_coords['x1'], search_coords['y1'], search_coords['x2'], search_coords['y2']
    ax1, ay1 = max(0, ax1), max(0, ay1)
    ax2, ay2 = min(img.shape[1], ax2), min(img.shape[0], ay2)
    course_search_area_img = img[ay1:ay2, ax1:ax2]
    
    if course_search_area_img.size > 0:
        course_gemini_input_path = os.path.join(OUTPUT_DIR, f"{base_filename}_course_gemini_input.png")
        cv2.imwrite(course_gemini_input_path, course_search_area_img)

    # --- 単独レースかの判別ロジック ---
    is_single_course = False
    single_coords = config.SINGLE_COURSE_NAME_AREA
    
    print(f"[DEBUG] SINGLE_COURSE_NAME_AREA: {single_coords}")
    
    sx1, sy1, sx2, sy2 = single_coords['x1'], single_coords['y1'], single_coords['x2'], single_coords['y2']
    sx1, sy1 = max(0, sx1), max(0, sy1)
    sx2, sy2 = min(img.shape[1], sx2), min(img.shape[0], sy2)
    center_area = img[sy1:sy2, sx1:sx2]

    if center_area.size > 0:
        ### 追加 ###
        # デバッグフォルダの存在を確認
        os.makedirs(DEBUG_DIR, exist_ok=True)
        # デバッグ用の画像として保存
        debug_save_path = os.path.join(DEBUG_DIR, f"{base_filename}_single_course_check_area.png")
        cv2.imwrite(debug_save_path, center_area)
        print(f"[DEBUG] 確認エリアの画像を保存しました: {os.path.basename(debug_save_path)}")
        #############

        hsv = cv2.cvtColor(center_area, cv2.COLOR_BGR2HSV)
        lower_black = np.array([0, 0, 0])
        upper_black = np.array([180, 255, 70])
        mask = cv2.inRange(hsv, lower_black, upper_black)
        
        pixel_count = cv2.countNonZero(mask)
        print(f"[DEBUG] 比較対象 (黒ピクセル数): {pixel_count}")
        
        if pixel_count > 5000:
            is_single_course = True
            print("[imaging] INFO: 単独コースレースを検出しました。")

    return rate_save_path, course_gemini_input_path, participant_count, is_single_course

def draw_debug_overlay(image, state, course_name=None, pre_race_rate=None):
    """指定された監視状態に基づいて、画像にデバッグ用の枠線を描画する"""
    debug_img = image.copy()
    
    if state == "course_decision":
        state_text = "State: Waiting for Course Decision"
        for coords in config.ALL_PLAYER_SLOTS:
            cv2.rectangle(debug_img, (coords['x1'], coords['y1']), (coords['x2'], coords['y2']), (0, 255, 255), 2)
        search_coords = config.COURSE_SEARCH_AREA
        cv2.rectangle(debug_img, (search_coords['x1'], search_coords['y1']), (search_coords['x2'], search_coords['y2']), (255, 0, 255), 2)
        single_coords = config.SINGLE_COURSE_NAME_AREA
        cv2.rectangle(debug_img, (single_coords['x1'], single_coords['y1']), (single_coords['x2'], single_coords['y2']), (0, 0, 255), 2)


    elif state == "result":
        state_text = f"State: Waiting for Result (Course: {course_name}, Pre-Rate: {pre_race_rate})"
        for i in range(1, 14):
            x1, y1, x2, y2 = config.RESULT_COORDS[f'rate_{i}']
            cv2.rectangle(debug_img, (x1, y1), (x2, y2), (0, 255, 0), 2)
    
    else: state_text = "State: Unknown"

    cv2.putText(debug_img, state_text, (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2, cv2.LINE_AA)
    return debug_img

def check_for_highlight(image):
    """
    リザルト画面にプレイヤーのハイライト（黄色い背景）が存在するかをチェックする。
    """
    lower_highlight = np.array(config.LOWER_HIGHLIGHT)
    upper_highlight = np.array(config.UPPER_HIGHLIGHT)

    try:
        y1 = config.RESULT_COORDS['rank_1'][1]
        y2 = config.RESULT_COORDS['rank_13'][3]
        x1 = config.RESULT_COORDS['rank_1'][0]
        x2 = config.RESULT_COORDS['rate_1'][2]
        
        search_roi = image[y1:y2, x1:x2]
    except (KeyError, IndexError):
        print("[imaging] ERROR: config.pyのRESULT_COORDSの定義が不完全です。")
        return False

    hsv_roi = cv2.cvtColor(search_roi, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv_roi, lower_highlight, upper_highlight)
    
    highlight_pixel_count = cv2.countNonZero(mask)
    if highlight_pixel_count > 500:
        return True
    else:
        return False