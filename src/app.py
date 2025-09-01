import cv2
import numpy as np
import os
from datetime import datetime
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import shutil
from pygrabber.dshow_graph import FilterGraph
import pygetwindow as gw
import pytesseract
import win32gui
import win32ui
import win32con
from ctypes import windll
import csv
import configparser

import analysis
import config 

# --- パス設定 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'temp')
DEBUG_DIR = os.path.join(SCRIPT_DIR, '..', 'data', 'debug')
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.ini')

# --- 設定 ---
MONITORING_INTERVAL = 2 
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

monitoring_active = False 
request_debug_capture = False

def capture_win_bg(hwnd):
    left, top, right, bot = win32gui.GetClientRect(hwnd)
    w, h = right - left, bot - top
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
    saveDC.SelectObject(saveBitMap)
    windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)
    bmpstr = saveBitMap.GetBitmapBits(True)
    img = np.frombuffer(bmpstr, dtype='uint8').reshape(h, w, 4)
    win32gui.DeleteObject(saveBitMap.GetHandle()); saveDC.DeleteDC(); mfcDC.DeleteDC(); win32gui.ReleaseDC(hwnd, hwndDC)
    return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

def monitor_loop(target_name, mode, app_instance):
    global monitoring_active, request_debug_capture
    monitoring_active = True 
    if not os.path.exists(TESSERACT_PATH):
        app_instance.update_status("エラー: Tesseract-OCRが見つかりません。"); app_instance.reset_gui_state(); return
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    cap = None; target_hwnd = None
    if mode == "device":
        cap = cv2.VideoCapture(target_name, cv2.CAP_DSHOW)
        if not cap.isOpened(): app_instance.update_status(f"エラー: デバイス {target_name} を開けません"); app_instance.reset_gui_state(); return
    elif mode == "window":
        try: target_hwnd = gw.getWindowsWithTitle(target_name)[0]._hWnd
        except IndexError: app_instance.update_status("エラー: ウィンドウが見つかりません。"); app_instance.reset_gui_state(); return
    
    app_instance.update_status("監視中 (コース決定画面を待っています)...")
    
    while monitoring_active:
        raw_frame = None 
        if mode == "device":
            ret, raw_frame = cap.read()
            if not ret: time.sleep(0.1); continue
        elif mode == "window":
            try:
                if not win32gui.IsWindow(target_hwnd): app_instance.update_status("エラー: ウィンドウが閉じられました。"); break
                raw_frame = capture_win_bg(target_hwnd)
            except Exception: app_instance.update_status("エラー: ウィンドウのキャプチャに失敗しました。"); break
        if raw_frame is None: time.sleep(MONITORING_INTERVAL); continue
        
        fhd_frame = cv2.resize(raw_frame, (1920, 1080), interpolation=cv2.INTER_AREA)

        if request_debug_capture:
            state = "course_decision" if app_instance.current_course_name is None else "result"
            debug_img = analysis.imaging.draw_debug_overlay(fhd_frame, state, app_instance.current_course_name, app_instance.pre_race_rate)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            os.makedirs(DEBUG_DIR, exist_ok=True)
            save_path = os.path.join(DEBUG_DIR, f"debug_capture_{timestamp}.png")
            if cv2.imwrite(save_path, debug_img):
                app_instance.update_status(f"デバッグ画像を保存しました: {os.path.basename(save_path)}")
            else:
                app_instance.update_status(f"エラー: デバッグ画像の保存に失敗しました。")
            request_debug_capture = False

        def is_rate_detected_in_list(coord_list, frame):
            detected_count = 0
            for coords in coord_list:
                if isinstance(coords, tuple):
                    coords = {'x1': coords[0], 'y1': coords[1], 'x2': coords[2], 'y2': coords[3]}

                x1, y1, x2, y2 = coords['x1'], coords['y1'], coords['x2'], coords['y2']
                if y2 > frame.shape[0] or x2 > frame.shape[1]: continue
                roi = frame[y1:y2, x1:x2]
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
                text = pytesseract.image_to_string(gray, config=r'--psm 7 -c tessedit_char_whitelist=0123456789').strip()
                if text.isdigit() and len(text) >= 3: 
                    detected_count += 1
            return detected_count

        # --- 監視ロジック ---
        if app_instance.current_course_name is None:
            if is_rate_detected_in_list(config.ALL_PLAYER_SLOTS, fhd_frame) >= 1:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_path = os.path.join(OUTPUT_DIR, f"course_screen_{timestamp}.png")
                cv2.imwrite(output_path, fhd_frame)
                time.sleep(0.1) 
                
                app_instance.update_status("コース決定画面を検出。解析中...")
                rate, course, p_count = analysis.get_course_and_pre_race_rate(output_path)
                
                if rate is not None and course != "コース不明":
                    app_instance.pre_race_rate = rate
                    app_instance.current_course_name = course
                    app_instance.participant_count = p_count
                    app_instance.update_status(f"コース:「{course}」({p_count}人) / あなたのレート: {rate} | リザルト画面を待機中...")
                else:
                    app_instance.update_status("コース解析に失敗。再試行します...")
                
                time.sleep(10); continue
        else: # リザルト画面待機中
            rate_count = is_rate_detected_in_list(config.RESULT_COORDS.values(), fhd_frame)
            highlight_found = analysis.imaging.check_for_highlight(fhd_frame)

            if rate_count >= 2 and highlight_found:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_path = os.path.join(OUTPUT_DIR, f"result_screen_{timestamp}.png")
                cv2.imwrite(output_path, fhd_frame)
                time.sleep(0.1)
                
                app_instance.update_status(f"リザルト画面を検出。解析中...")
                try:
                    is_debug = app_instance.debug_mode_var.get()
                    new_result = analysis.process_result_image(output_path, app_instance.current_course_name, app_instance.pre_race_rate, app_instance.participant_count, is_debug)
                    if new_result: app_instance.update_log_display([new_result])
                    
                    app_instance.current_course_name = None; app_instance.pre_race_rate = None; app_instance.participant_count = 0
                    
                    app_instance.update_status("監視中 (コース決定画面を待っています)...")

                except Exception as e:
                    app_instance.update_status(f"解析エラー: {e}")
                    app_instance.current_course_name = None; app_instance.pre_race_rate = None; app_instance.participant_count = 0
                    import traceback; traceback.print_exc(); time.sleep(5)
        
        time.sleep(MONITORING_INTERVAL)
    if cap: cap.release()
    if not app_instance.root.winfo_exists(): return
    app_instance.reset_gui_state()

def load_setting(key):
    config_parser = configparser.ConfigParser(); config_parser.read(CONFIG_FILE)
    if 'Settings' in config_parser and key in config_parser['Settings']: return config_parser['Settings'][key]
    return None

def save_setting(key, value):
    config_parser = configparser.ConfigParser(); config_parser.read(CONFIG_FILE)
    if 'Settings' not in config_parser: config_parser['Settings'] = {}
    config_parser['Settings'][key] = str(value)
    with open(CONFIG_FILE, 'w') as f: config_parser.write(f)

class ControlPanel(tk.Toplevel):
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.master_app = app_instance
        self.title("監視コントロール"); self.geometry("450x200")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)
        control_frame = ttk.LabelFrame(self, text="監視設定", padding="10")
        control_frame.pack(fill="both", expand=True, padx=10, pady=10)
        source_frame = ttk.Frame(control_frame)
        source_frame.pack(fill='x', pady=5)
        ttk.Label(source_frame, text="監視ソース:").pack(side="left", padx=5)
        ttk.Radiobutton(source_frame, text="映像デバイス", variable=self.master_app.source_type_var, value="device", command=self.master_app.update_dropdown).pack(side="left")
        ttk.Radiobutton(source_frame, text="ウィンドウ", variable=self.master_app.source_type_var, value="window", command=self.master_app.update_dropdown).pack(side="left")
        self.master_app.dropdown = ttk.OptionMenu(control_frame, self.master_app.target_var, "")
        self.master_app.dropdown.pack(pady=5, fill='x', expand=True)
        debug_check = ttk.Checkbutton(control_frame, text="デバッグモード (解析画像を保存する)", variable=self.master_app.debug_mode_var)
        debug_check.pack(anchor='w', pady=5)
        self.master_app.start_button_panel = ttk.Button(control_frame, text="監視開始", command=self.master_app.on_start_click)
        self.master_app.start_button_panel.pack(side="bottom", pady=10, fill='x')

class EditRaceWindow(tk.Toplevel):
    def __init__(self, master, app_instance, original_data):
        super().__init__(master)
        self.app = app_instance
        self.original_data = original_data
        self.title("レース記録の編集")
        self.geometry("450x300")

        self.all_courses = ["不明"] + config.COURSE_NAMES
        self.all_courses_with_none = ["（無し）", "不明"] + config.COURSE_NAMES

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="始点コース:").grid(row=0, column=0, sticky="w", pady=2)
        self.course_start_var = tk.StringVar()
        self.course_start_combo = ttk.Combobox(main_frame, textvariable=self.course_start_var, values=self.all_courses, state="readonly")
        self.course_start_combo.grid(row=0, column=1, sticky="ew", pady=2)
        # ★追加: 始点コースが選択されたら、終点の選択肢を更新するイベントを紐付ける
        self.course_start_combo.bind("<<ComboboxSelected>>", self.on_start_course_selected)

        ttk.Label(main_frame, text="終点コース:").grid(row=1, column=0, sticky="w", pady=2)
        self.course_end_var = tk.StringVar()
        self.course_end_combo = ttk.Combobox(main_frame, textvariable=self.course_end_var, state="readonly")
        self.course_end_combo.grid(row=1, column=1, sticky="ew", pady=2)

        # (順位、参加人数、レートの入力欄は変更なし)
        ttk.Label(main_frame, text="順位:").grid(row=2, column=0, sticky="w", pady=2)
        self.rank_var = tk.StringVar()
        self.rank_entry = ttk.Entry(main_frame, textvariable=self.rank_var)
        self.rank_entry.grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(main_frame, text="参加人数:").grid(row=3, column=0, sticky="w", pady=2)
        self.participants_var = tk.StringVar()
        self.participants_entry = ttk.Entry(main_frame, textvariable=self.participants_var)
        self.participants_entry.grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Label(main_frame, text="最終レート:").grid(row=4, column=0, sticky="w", pady=2)
        self.rate_var = tk.StringVar()
        self.rate_entry = ttk.Entry(main_frame, textvariable=self.rate_var)
        self.rate_entry.grid(row=4, column=1, sticky="ew", pady=2)

        main_frame.columnconfigure(1, weight=1)
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(button_frame, text="保存", command=self.save_changes).pack(side="right", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side="right")

        self.populate_data()

    # ★追加: 始点コースが選ばれたときに実行される関数
    def on_start_course_selected(self, event=None):
        selected_start = self.course_start_var.get()
        
        # 「不明」が選ばれたら、終点は自由に選択可能にする
        if selected_start == "不明":
            self.course_end_combo['values'] = self.all_courses_with_none
            self.course_end_var.set("（無し）")
            return
            
        # configから有効な終点のリストを取得
        valid_ends = config.VALID_ROUTES_MAP.get(selected_start, [])
        self.course_end_combo['values'] = valid_ends
        
        # 終点の選択肢があれば、最初の項目をデフォルトで選択
        if valid_ends:
            self.course_end_var.set(valid_ends[0])
        else:
            self.course_end_var.set("") # 選択肢がなければ空にする

    def populate_data(self):
        course_full = self.original_data['Course']
        start_course, end_course = "", ""
        if "→" in course_full:
            start, end = course_full.split("→")
            start_course, end_course = start.strip(), end.strip()
        else:
            start_course, end_course = course_full, "（無し）"
        
        self.course_start_var.set(start_course)
        
        # ★変更: 終点の選択肢を更新してから値を設定
        self.on_start_course_selected()
        
        self.course_end_var.set(end_course)
        self.rank_var.set(self.original_data['Rank'])
        self.participants_var.set(self.original_data['Participants'])
        self.rate_var.set(self.original_data['Rate'])
    
    def save_changes(self):
        try:
            new_data = self.original_data.copy()
            start = self.course_start_var.get()
            end = self.course_end_var.get()
            if end == "（無し）" or not end:
                new_data['Course'] = start
            else:
                new_data['Course'] = f"{start} → {end}"
            new_data['Rank'] = int(self.rank_var.get())
            new_data['Participants'] = int(self.participants_var.get())
            new_data['Rate'] = int(self.rate_var.get())
            self.app.save_edited_race(new_data)
            self.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "順位、参加人数、レートは数字で入力してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"保存中にエラーが発生しました: {e}")

class AddRaceWindow(tk.Toplevel):
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.app = app_instance
        self.title("レース記録の手動追加")
        self.geometry("450x300")

        self.all_courses = ["不明"] + config.COURSE_NAMES
        self.all_courses_with_none = ["（無し）", "不明"] + config.COURSE_NAMES

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="始点コース:").grid(row=0, column=0, sticky="w", pady=2)
        self.course_start_var = tk.StringVar(value="不明")
        self.course_start_combo = ttk.Combobox(main_frame, textvariable=self.course_start_var, values=self.all_courses, state="readonly")
        self.course_start_combo.grid(row=0, column=1, sticky="ew", pady=2)
        # 始点コースが選択されたら、終点の選択肢を更新するイベントを紐付ける
        self.course_start_combo.bind("<<ComboboxSelected>>", self.on_start_course_selected)

        ttk.Label(main_frame, text="終点コース:").grid(row=1, column=0, sticky="w", pady=2)
        self.course_end_var = tk.StringVar(value="（無し）")
        self.course_end_combo = ttk.Combobox(main_frame, textvariable=self.course_end_var, state="readonly")
        self.course_end_combo.grid(row=1, column=1, sticky="ew", pady=2)

        # (順位、参加人数、レートの入力欄は変更なし)
        ttk.Label(main_frame, text="順位:").grid(row=2, column=0, sticky="w", pady=2)
        self.rank_var = tk.StringVar()
        self.rank_entry = ttk.Entry(main_frame, textvariable=self.rank_var)
        self.rank_entry.grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(main_frame, text="参加人数:").grid(row=3, column=0, sticky="w", pady=2)
        self.participants_var = tk.StringVar()
        self.participants_entry = ttk.Entry(main_frame, textvariable=self.participants_var)
        self.participants_entry.grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Label(main_frame, text="最終レート:").grid(row=4, column=0, sticky="w", pady=2)
        self.rate_var = tk.StringVar()
        self.rate_entry = ttk.Entry(main_frame, textvariable=self.rate_var)
        self.rate_entry.grid(row=4, column=1, sticky="ew", pady=2)

        main_frame.columnconfigure(1, weight=1)
        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(button_frame, text="追加", command=self.add_race).pack(side="right", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side="right")
        
        # 初期状態を設定
        self.on_start_course_selected()

    # 始点コースが選ばれたときに実行される関数
    def on_start_course_selected(self, event=None):
        selected_start = self.course_start_var.get()
        
        if selected_start == "不明":
            self.course_end_combo['values'] = self.all_courses_with_none
            self.course_end_var.set("（無し）")
            return
            
        valid_ends = config.VALID_ROUTES_MAP.get(selected_start, [])
        self.course_end_combo['values'] = valid_ends
        
        if valid_ends:
            self.course_end_var.set(valid_ends[0])
        else:
            self.course_end_var.set("")

    def add_race(self):
        try:
            start = self.course_start_var.get()
            end = self.course_end_var.get()
            if end == "（無し）" or not end:
                course = start
            else:
                course = f"{start} → {end}"
            new_data = {
                'Course': course, 'Rank': int(self.rank_var.get()),
                'Participants': int(self.participants_var.get()), 'Rate': int(self.rate_var.get())
            }
            self.app.add_new_race(new_data)
            self.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "順位、参加人数、レートは数字で入力してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"追加中にエラーが発生しました: {e}")

    def __init__(self, master, app_instance):
        super().__init__(master)
        self.app = app_instance
        self.title("レース記録の手動追加")
        self.geometry("450x300")

        self.course_list_start = ["不明"] + config.COURSE_NAMES
        self.course_list_end = ["（無し）", "不明"] + config.COURSE_NAMES

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="始点コース:").grid(row=0, column=0, sticky="w", pady=2)
        self.course_start_var = tk.StringVar(value="不明")
        self.course_start_combo = ttk.Combobox(main_frame, textvariable=self.course_start_var, values=self.course_list_start, state="readonly")
        self.course_start_combo.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(main_frame, text="終点コース:").grid(row=1, column=0, sticky="w", pady=2)
        self.course_end_var = tk.StringVar(value="（無し）")
        self.course_end_combo = ttk.Combobox(main_frame, textvariable=self.course_end_var, values=self.course_list_end, state="readonly")
        self.course_end_combo.grid(row=1, column=1, sticky="ew", pady=2)

        ttk.Label(main_frame, text="順位:").grid(row=2, column=0, sticky="w", pady=2)
        self.rank_var = tk.StringVar()
        self.rank_entry = ttk.Entry(main_frame, textvariable=self.rank_var)
        self.rank_entry.grid(row=2, column=1, sticky="ew", pady=2)
        
        ttk.Label(main_frame, text="参加人数:").grid(row=3, column=0, sticky="w", pady=2)
        self.participants_var = tk.StringVar()
        self.participants_entry = ttk.Entry(main_frame, textvariable=self.participants_var)
        self.participants_entry.grid(row=3, column=1, sticky="ew", pady=2)

        ttk.Label(main_frame, text="最終レート:").grid(row=4, column=0, sticky="w", pady=2)
        self.rate_var = tk.StringVar()
        self.rate_entry = ttk.Entry(main_frame, textvariable=self.rate_var)
        self.rate_entry.grid(row=4, column=1, sticky="ew", pady=2)

        main_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(button_frame, text="追加", command=self.add_race).pack(side="right", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side="right")
    
    def add_race(self):
        try:
            start = self.course_start_var.get()
            end = self.course_end_var.get()
            if end == "（無し）" or not end:
                course = start
            else:
                course = f"{start} → {end}"
            
            new_data = {
                'Course': course,
                'Rank': int(self.rank_var.get()),
                'Participants': int(self.participants_var.get()),
                'Rate': int(self.rate_var.get())
            }

            self.app.add_new_race(new_data)
            self.destroy()

        except ValueError:
            messagebox.showerror("入力エラー", "順位、参加人数、レートは数字で入力してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"追加中にエラーが発生しました: {e}")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("MKW レースリザルト解析")
        self.root.geometry("600x600")
        self.source_type_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.debug_mode_var = tk.BooleanVar(value=False)
        self.targets = {}
        self.LOG_DISPLAY_LIMIT = 30
        
        self.current_course_name = None
        self.pre_race_rate = None
        self.participant_count = 0

        menubar = tk.Menu(root); root.config(menu=menubar)
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=settings_menu)
        settings_menu.add_command(label="監視コントロールを開く", command=self.open_control_panel)
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ツール", menu=tools_menu)
        tools_menu.add_command(label="レース記録を手動追加", command=self.open_add_race_window)
        tools_menu.add_command(label="デバッグキャプチャ", command=self.on_debug_capture)
        tools_menu.add_command(label="監視状態を強制切替", command=self.force_switch_state)
        tools_menu.add_separator()
        tools_menu.add_command(label="一時ファイルを消去", command=self.clear_temp_files)
        tools_menu.add_command(label="デバッグファイルを消去", command=self.clear_debug_files)
        tools_menu.add_separator()
        tools_menu.add_command(label="全ログを消去", command=self.clear_logs)
        main_frame = ttk.Frame(root, padding="10"); main_frame.pack(fill="both", expand=True)
        stats_frame = ttk.LabelFrame(main_frame, text="統計情報", padding="10")
        stats_frame.pack(fill="x", pady=5)
        self.total_races_var = tk.StringVar(value="合計レース数: -")
        self.avg_rate_var = tk.StringVar(value="平均レート(100戦): -")
        self.max_rate_var = tk.StringVar(value="最高レート: -")
        self.min_rate_var = tk.StringVar(value="最低レート: -")
        ttk.Label(stats_frame, textvariable=self.total_races_var).grid(row=0, column=0, sticky='w', padx=5)
        ttk.Label(stats_frame, textvariable=self.avg_rate_var).grid(row=0, column=1, sticky='w', padx=5)
        ttk.Label(stats_frame, textvariable=self.max_rate_var).grid(row=1, column=0, sticky='w', padx=5)
        ttk.Label(stats_frame, textvariable=self.min_rate_var).grid(row=1, column=1, sticky='w', padx=5)
        self.status_label = ttk.Label(main_frame, text="待機中...", font=("", 10), wraplength=580)
        self.status_label.pack(pady=10, fill='x')
        self.stop_button_main = ttk.Button(main_frame, text="監視停止", command=self.on_stop_click, state="disabled")
        self.stop_button_main.pack(fill='x', pady=5)
        log_frame = ttk.LabelFrame(main_frame, text=f"直近{self.LOG_DISPLAY_LIMIT}レースの結果", padding="10")
        log_frame.pack(fill="both", expand=True, pady=10)
        
        columns = ("timestamp", "course", "rank", "rate", "rate_change")
        self.log_tree = ttk.Treeview(log_frame, columns=columns, show="headings")
        self.log_tree.heading("timestamp", text="記録時間"); self.log_tree.heading("course", text="コース"); self.log_tree.heading("rank", text="順位/人数"); self.log_tree.heading("rate", text="最終レート"); self.log_tree.heading("rate_change", text="レート変動")
        self.log_tree.column("timestamp", width=140, anchor='center'); self.log_tree.column("course", width=150); self.log_tree.column("rank", width=60, anchor='center'); self.log_tree.column("rate", width=70, anchor='center'); self.log_tree.column("rate_change", width=70, anchor='center')
        
        self.log_tree.bind("<Double-1>", self.on_double_click)

        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_tree.yview); self.log_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y'); self.log_tree.pack(fill="both", expand=True)
        self.control_panel = ControlPanel(self.root, self); self.control_panel.withdraw()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        os.makedirs(OUTPUT_DIR, exist_ok=True); os.makedirs(DEBUG_DIR, exist_ok=True)
        self.load_initial_logs_and_stats()
        self.update_dropdown()
        self.root.after(100, self.initialize_source)

    def on_double_click(self, event):
        item_id = self.log_tree.focus()
        if not item_id: return
        
        values = self.log_tree.item(item_id, 'values')
        target_timestamp = values[0]
        full_data_row = self.find_row_in_csv(target_timestamp)

        if full_data_row:
            EditRaceWindow(self.root, self, full_data_row)
        else:
            messagebox.showerror("エラー", "元のデータが見つかりませんでした。")
            
    def find_row_in_csv(self, timestamp):
        try:
            with open(analysis.OUTPUT_CSV_PATH, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                header = next(reader)
                ts_idx = header.index('Timestamp')
                for row in reader:
                    if row[ts_idx] == timestamp:
                        return dict(zip(header, row))
        except (FileNotFoundError, ValueError, IndexError) as e:
            print(f"CSV検索エラー: {e}")
            return None
        return None

    def save_edited_race(self, new_data):
        try:
            with open(analysis.OUTPUT_CSV_PATH, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                header = next(reader)
                all_rows = list(reader)

            filename_idx = header.index('Filename')
            updated = False
            for i, row in enumerate(all_rows):
                if row[filename_idx] == new_data['Filename']:
                    updated_row = [new_data.get(h, '') for h in header]
                    
                    prev_rate = int(all_rows[i-1][header.index('Rate')]) if i > 0 else 0
                    if prev_rate > 0:
                        new_rate = int(new_data['Rate'])
                        updated_row[header.index('Rate Change')] = new_rate - prev_rate

                    all_rows[i] = updated_row
                    updated = True
                    break
            
            if not updated: raise ValueError("更新対象の行が見つかりません。")

            with open(analysis.OUTPUT_CSV_PATH, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(all_rows)

            self.load_initial_logs_and_stats()
            self.update_status("レース記録を更新しました。")

        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました: {e}")

    def open_add_race_window(self):
        AddRaceWindow(self.root, self)

    def add_new_race(self, new_data):
        try:
            header = ['Filename', 'Timestamp', 'Course', 'Rank', 'Participants', 'Rate', 'Rate Change']
            last_rate = analysis.get_last_race_rate(analysis.OUTPUT_CSV_PATH)
            
            final_rate = new_data['Rate']
            rate_change = 0
            if last_rate is not None:
                rate_change = final_rate - last_rate
            
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            new_row = [
                f"manual_{timestamp.replace(' ', '_').replace(':', '')}",
                timestamp,
                new_data['Course'],
                new_data['Rank'],
                new_data['Participants'],
                final_rate,
                rate_change
            ]
            
            is_new_file = not os.path.exists(analysis.OUTPUT_CSV_PATH) or os.path.getsize(analysis.OUTPUT_CSV_PATH) == 0
            with open(analysis.OUTPUT_CSV_PATH, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                if is_new_file:
                    writer.writerow(header)
                writer.writerow(new_row)

            self.load_initial_logs_and_stats()
            self.update_status("レース記録を手動で追加しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"データの追加に失敗しました: {e}")

    def force_switch_state(self):
        if not monitoring_active:
            messagebox.showwarning("情報", "監視中にのみ実行できます。")
            return
        
        if self.current_course_name is None:
            self.current_course_name = "手動切替"
            self.pre_race_rate = 0
            self.participant_count = 0
            self.update_status("状態を強制的に「リザルト待機」に変更しました。")
        else:
            self.current_course_name = None
            self.pre_race_rate = None
            self.participant_count = 0
            self.update_status("状態を強制的に「コース決定待機」に変更しました。")

    def on_debug_capture(self):
        global request_debug_capture
        if monitoring_active:
            request_debug_capture = True
            self.update_status("デバッグキャプチャをリクエストしました。")
        else:
            messagebox.showwarning("情報", "デバッグキャプチャは監視中にのみ実行できます。")

    def initialize_source(self):
        last_name = load_setting('last_source_name'); last_type = load_setting('last_source_type')
        if not last_name or not last_type:
            self.update_status("初回起動です。メニューから[設定]>[監視コントロールを開く]で監視ソースを選択してください。")
            self.open_control_panel(); return
        
        self.source_type_var.set(last_type)
        self.update_dropdown()
        
        if last_name in self.targets:
            self.target_var.set(last_name)
            self.update_status(f"前回のソース '{last_name}' を検出し、監視を自動開始します...")
            self.on_start_click()
        else:
            self.update_status(f"エラー: 前回のソース '{last_name}' が見つかりません。再選択してください。")
            self.open_control_panel()

    def open_control_panel(self):
        self.control_panel.deiconify(); self.control_panel.lift(); self.control_panel.focus_set()

    def update_dropdown(self):
        menu = self.dropdown["menu"]; menu.delete(0, "end")
        source_type = self.source_type_var.get()
        if source_type == "device": self.targets = {name: i for i, name in enumerate(FilterGraph().get_input_devices())}
        else: self.targets = {win.title: win.title for win in gw.getWindowsWithTitle('') if win.title and win.visible and win.title != self.root.title() and win.title != self.control_panel.title()}
        if not self.targets: self.targets = {"利用可能なターゲットがありません": None}
        for name in self.targets.keys(): menu.add_command(label=name, command=lambda value=name: self.target_var.set(value))
        current_selection = self.target_var.get()
        if not current_selection or current_selection not in self.targets: self.target_var.set(list(self.targets.keys())[0])

    def on_start_click(self):
        target_key = self.target_var.get(); source_type = self.source_type_var.get(); target_value = self.targets.get(target_key)
        if target_value is None: self.update_status("エラー: 有効な監視ターゲットが選択されていません。"); return
        save_setting('last_source_name', target_key); save_setting('last_source_type', source_type)
        self.start_button_panel.config(state="disabled"); self.stop_button_main.config(state="normal")
        self.control_panel.withdraw()
        self.root.focus_set()
        self.root.lift()
        threading.Thread(target=monitor_loop, args=(target_value, source_type, self), daemon=True).start()

    def on_stop_click(self):
        global monitoring_active; monitoring_active = False; self.reset_gui_state()
    def on_closing(self):
        global monitoring_active; monitoring_active = False; self.root.destroy()

    def reset_gui_state(self):
        if not self.root.winfo_exists(): return
        current_text = self.status_label.cget("text")
        if "エラー" not in current_text and "クールダウン中" not in current_text and "解析完了" not in current_text:
                self.update_status("停止しました。")
        self.start_button_panel.config(state="normal"); self.stop_button_main.config(state="disabled")
        self.open_control_panel()

    def update_status(self, text):
        if self.root.winfo_exists(): self.root.after(0, self.status_label.config, {'text': text})
    
    def load_initial_logs_and_stats(self):
        self.log_tree.delete(*self.log_tree.get_children())
        self.update_stats()
        if not os.path.exists(analysis.OUTPUT_CSV_PATH): return
        with open(analysis.OUTPUT_CSV_PATH, 'r', newline='', encoding='utf-8-sig') as f:
            reader=list(csv.reader(f))
            if len(reader) < 2: return
            header=reader[0]; all_logs=reader[1:]; all_logs.reverse()
            recent_logs=all_logs[:self.LOG_DISPLAY_LIMIT]
            try:
                ts_idx, course_idx, rank_idx, p_idx, rate_idx, change_idx = header.index('Timestamp'), header.index('Course'), header.index('Rank'), header.index('Participants'), header.index('Rate'), header.index('Rate Change')
                for row in recent_logs:
                    rank_str = f"{row[rank_idx]}/{row[p_idx]}"
                    rate_change = int(row[change_idx])
                    formatted_change = f"+{rate_change}" if rate_change >= 0 else str(rate_change)
                    self.log_tree.insert('', 'end', values=(row[ts_idx], row[course_idx], rank_str, row[rate_idx], formatted_change))
            except (ValueError, IndexError): print("CSVヘッダーの形式が正しくないか、データが不足しています。")

    def update_log_display(self, new_results):
        if self.root.winfo_exists(): self.root.after(0, self._update_log_display, new_results)

    def _update_log_display(self, new_results):
        for result in reversed(new_results):
            # result = [filename, timestamp, course, rank, p_count, rate, rate_change]
            timestamp, course, rank, p_count, rate, rate_change = result[1], result[2], result[3], result[4], result[5], result[6]
            rank_str = f"{rank}/{p_count}"
            formatted_change = f"+{rate_change}" if rate_change >= 0 else str(rate_change)
            self.log_tree.insert('', 0, values=(timestamp, course, rank_str, rate, formatted_change))
        while len(self.log_tree.get_children()) > self.LOG_DISPLAY_LIMIT:
            self.log_tree.delete(self.log_tree.get_children()[-1])
        self.update_stats()

    def clear_logs(self):
        if not os.path.exists(analysis.OUTPUT_CSV_PATH): messagebox.showinfo("情報", "消去するログがありません。"); return
        if messagebox.askyesno("確認", "本当にすべてのレースログを消去しますか？\nこの操作は元に戻せません。"):
            try:
                os.remove(analysis.OUTPUT_CSV_PATH); self.load_initial_logs_and_stats()
                messagebox.showinfo("成功", "すべてのログを消去しました。"); self.update_status("全ログを消去しました。")
            except Exception as e: messagebox.showerror("エラー", f"ログの消去に失敗しました: {e}")

    def clear_temp_files(self):
        if messagebox.askyesno("確認", "一時ファイル（スクリーンショット、切り抜き画像）をすべて消去しますか？"):
            try:
                count = 0
                temp_cropped_dir = os.path.join(OUTPUT_DIR, 'cropped')
                for folder in [OUTPUT_DIR, temp_cropped_dir]:
                    if not os.path.exists(folder): continue
                    for filename in os.listdir(folder):
                        file_path = os.path.join(folder, filename)
                        if os.path.isfile(file_path): os.remove(file_path); count += 1
                messagebox.showinfo("成功", f"{count} 個の一時ファイルを消去しました。")
            except Exception as e: messagebox.showerror("エラー", f"一時ファイルの消去に失敗しました: {e}")

    def clear_debug_files(self):
        if not os.path.exists(DEBUG_DIR) or not os.listdir(DEBUG_DIR): messagebox.showinfo("情報", "消去するデバッグファイルがありません。"); return
        if messagebox.askyesno("確認", "すべてのデバッグファイルを消去しますか？"):
            try:
                shutil.rmtree(DEBUG_DIR); os.makedirs(DEBUG_DIR)
                messagebox.showinfo("成功", "すべてのデバッグファイルを消去しました。")
            except Exception as e: messagebox.showerror("エラー", f"デバッグファイルの消去に失敗しました: {e}")

    def update_stats(self):
        if not os.path.exists(analysis.OUTPUT_CSV_PATH):
            self.total_races_var.set("合計レース数: 0"); self.avg_rate_var.set("平均レート(100戦): -")
            self.max_rate_var.set("最高レート: -"); self.min_rate_var.set("最低レート: -"); return
        with open(analysis.OUTPUT_CSV_PATH, 'r', newline='', encoding='utf-8-sig') as f:
            reader=list(csv.reader(f)); header=reader[0]; all_logs=reader[1:]
            if not all_logs: self.total_races_var.set("合計レース数: 0"); return
            try:
                rate_idx = header.index('Rate')
                all_rates = [int(row[rate_idx]) for row in all_logs]
                recent_100_rates = all_rates[-100:]
                self.total_races_var.set(f"合計レース数: {len(all_rates)}")
                self.avg_rate_var.set(f"平均レート(100戦): {sum(recent_100_rates)/len(recent_100_rates):.0f}")
                self.max_rate_var.set(f"最高レート: {max(all_rates)}")
                self.min_rate_var.set(f"最低レート: {min(all_rates)}")
            except (ValueError, IndexError):
                print("CSVヘッダーまたはデータ形式が不正です。統計を更新できません。")

    def get_previous_course_name(self):
        # この関数は現在直接使用されませんが、デバッグ等のために残しておきます。
        if not os.path.exists(analysis.OUTPUT_CSV_PATH): return None
        with open(analysis.OUTPUT_CSV_PATH, 'r', newline='', encoding='utf-8-sig') as f:
            reader=list(csv.reader(f))
            if len(reader) < 2: return None
            header=reader[0]; last_log=reader[-1]
            try:
                course_idx = header.index('Course')
                return last_log[course_idx]
            except (ValueError, IndexError): return None

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()