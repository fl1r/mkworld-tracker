# --- 座標設定ファイル ---
# このファイル内の値を変更することで、解析領域を調整できます。

# --- レース前の「コース決定画面」関連 ---

# 画面左側のプレイヤーリストのレートが表示される領域
COURSE_DECISION_RATE_COORDS_LEFT = [
    {'x1': 440, 'y1': 100 + i*76, 'x2': 520, 'y2': 140 + i*76} for i in range(12)
]
# 画面右側のプレイヤーリストのレートが表示される領域
COURSE_DECISION_RATE_COORDS_RIGHT = [
    {'x1': 945, 'y1': 100 + i*76, 'x2': 1025, 'y2': 140 + i*76} for i in range(12)
]
# 上記2つのリストを結合した、全24人分の座標リスト
ALL_PLAYER_SLOTS = COURSE_DECISION_RATE_COORDS_LEFT + COURSE_DECISION_RATE_COORDS_RIGHT

# コース名を探すための、画面右側の大きな探索範囲
COURSE_SEARCH_AREA = {'x1': 1100, 'y1': 100, 'x2': 1900, 'y2': 900}


# --- レース後の「リザルト画面」関連 ---

# 1位のプレイヤーの「順位」「レート」「得点」の基準座標
BASE_COORDS = {
    'rank':   {'x1': 1070, 'y1': 45, 'x2': 1135, 'y2': 110},
    'rate':   {'x1': 1730, 'y1': 45, 'x2': 1860, 'y2': 110},
    'rate_change': {'x1': 1630, 'y1': 45, 'x2': 1730, 'y2': 110},
}
# 各順位間のY軸方向の間隔
BASE_Y_STEP = 77

# 1位の座標を基準に、13位までの各順位の座標を自動計算
RESULT_COORDS = {}
for i in range(1, 14):
    y_offset = (i - 1) * BASE_Y_STEP
    RESULT_COORDS[f'rank_{i}'] = (BASE_COORDS['rank']['x1'], BASE_COORDS['rank']['y1'] + y_offset, BASE_COORDS['rank']['x2'], BASE_COORDS['rank']['y2'] + y_offset)
    RESULT_COORDS[f'rate_{i}'] = (BASE_COORDS['rate']['x1'], BASE_COORDS['rate']['y1'] + y_offset, BASE_COORDS['rate']['x2'], BASE_COORDS['rate']['y2'] + y_offset)
    RESULT_COORDS[f'rate_change_{i}'] = (BASE_COORDS['rate_change']['x1'], BASE_COORDS['rate_change']['y1'] + y_offset, BASE_COORDS['rate_change']['x2'], BASE_COORDS['rate_change']['y2'] + y_offset)
