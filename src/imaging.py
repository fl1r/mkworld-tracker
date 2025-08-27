import cv2
import numpy as np
import os
import config # 設定ファイルをインポート

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'temp', 'cropped')

def crop_image_for_result(image_path):
    """リザルト画面の画像を解析し、ハイライト位置と関連領域を切り抜く"""
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    img = cv2.imread(image_path)
    if img is None: return None
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    lower_highlight = np.array([15, 100, 100])
    upper_highlight = np.array([40, 255, 255])
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
    """コース決定画面を解析し、レート・コース部分を切り抜き、参加人数を数える"""
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    base_filename = os.path.splitext(os.path.basename(image_path))[0]
    img = cv2.imread(image_path)
    if img is None: return None, None, 0

    brightest_rate_roi = None
    max_brightness = 0
    participant_count = 0
    GRAY_THRESHOLD = 50 

    for coords in config.ALL_PLAYER_SLOTS:
        x1, y1, x2, y2 = coords['x1'], coords['y1'], coords['x2'], coords['y2']
        if y2 > img.shape[0] or x2 > img.shape[1]: continue
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

    # --- ★コース名検出ロジック (水平直線検出)★ ---
    course_save_path = None
    search_coords = config.COURSE_SEARCH_AREA
    ax1, ay1, ax2, ay2 = search_coords['x1'], search_coords['y1'], search_coords['x2'], search_coords['y2']
    search_area = img[ay1:ay2, ax1:ax2]

    # Cannyエッジ検出で画像のエッジ（輪郭）を見つける
    edges = cv2.Canny(search_area, 50, 150, apertureSize=3)
    
    # 確率的ハフ変換で直線を検出
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=100, minLineLength=100, maxLineGap=10)
    
    horizontal_lines = []
    if lines is not None:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            # ほぼ水平な直線（Y座標の変化が小さい）のみを抽出
            if abs(y1 - y2) < 5:
                horizontal_lines.append(line[0])

    debug_steps_dir = os.path.join(OUTPUT_DIR, "debug_steps")
    os.makedirs(debug_steps_dir, exist_ok=True)
    debug_img = search_area.copy()
    if lines is not None:
        for x1, y1, x2, y2 in horizontal_lines:
            cv2.line(debug_img, (x1, y1), (x2, y2), (0, 255, 0), 2)
    cv2.imwrite(os.path.join(debug_steps_dir, f"{base_filename}_1_detected_lines.png"), debug_img)

    # 2本以上の水平線が見つかった場合
    if len(horizontal_lines) >= 2:
        # Y座標でソートして、最も近い2本の線を見つける
        horizontal_lines.sort(key=lambda line: line[1])
        
        min_dist = float('inf')
        best_pair = None
        for i in range(len(horizontal_lines) - 1):
            dist = abs(horizontal_lines[i][1] - horizontal_lines[i+1][1])
            if 20 < dist < 100: # 適切な距離のペアを探す
                if dist < min_dist:
                    min_dist = dist
                    best_pair = (horizontal_lines[i], horizontal_lines[i+1])

        if best_pair:
            line1, line2 = best_pair
            # 2本の線の平均的なX座標の範囲と、Y座標の範囲を計算
            crop_x1 = min(line1[0], line2[0])
            crop_x2 = max(line1[2], line2[2])
            crop_y1 = min(line1[1], line2[1])
            crop_y2 = max(line1[3], line2[3])

            padding = 5
            roi = search_area[max(0, crop_y1-padding):min(crop_y2+padding, search_area.shape[0]), 
                              max(0, crop_x1-padding):min(crop_x2+padding, search_area.shape[1])]
            
            course_save_path = os.path.join(OUTPUT_DIR, f"{base_filename}_course_name.png")
            cv2.imwrite(course_save_path, roi)
            cv2.imwrite(os.path.join(debug_steps_dir, f"{base_filename}_4_final_crop.png"), roi)

    if not course_save_path:
        print("[imaging] WARNING: コース名が見つかりませんでした。")

    return rate_save_path, course_save_path, participant_count

def draw_debug_overlay(image, state, course_name=None, pre_race_rate=None):
    """指定された監視状態に基づいて、画像にデバッグ用の枠線を描画する"""
    debug_img = image.copy()
    
    if state == "course_decision":
        state_text = "State: Waiting for Course Decision"
        for coords in config.ALL_PLAYER_SLOTS:
            cv2.rectangle(debug_img, (coords['x1'], coords['y1']), (coords['x2'], coords['y2']), (0, 255, 255), 2)
        search_coords = config.COURSE_SEARCH_AREA
        cv2.rectangle(debug_img, (search_coords['x1'], search_coords['y1']), (search_coords['x2'], search_coords['y2']), (255, 0, 255), 2)

    elif state == "result":
        state_text = f"State: Waiting for Result (Course: {course_name}, Pre-Rate: {pre_race_rate})"
        for i in range(1, 14):
            x1, y1, x2, y2 = config.RESULT_COORDS[f'rate_{i}']
            cv2.rectangle(debug_img, (x1, y1), (x2, y2), (0, 255, 0), 2)
    
    else: state_text = "State: Unknown"

    cv2.putText(debug_img, state_text, (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2, cv2.LINE_AA)
    return debug_img

