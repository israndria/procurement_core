import cv2
import numpy as np
import mss
import tkinter as tk
import win32api 
import win32con
import win32gui 
import time
import csv
import os
import ctypes
from datetime import datetime
from threading import Thread

# --- KONFIGURASI ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()

# 1. UKURAN KOTAK SNAP (DIPERKECIL AGAR LEBIH PAS)
CROP_WIDTH = 250   
CROP_HEIGHT = 35   

# AREA SCAN (Sisi Kanan Saja)
SIDEBAR_WIDTH = 550 

FILE_TXT_PREFIX = "History_Game"

# SENSITIVITAS
THRESHOLD = 0.50   

# POSISI TEKS
OFFSET_TEKS = 180  

# STABILIZER
DEADZONE = 4

# WARNA
BTN_COLOR = "#111111" 
BTN_MONSTER = "#003300"
BTN_UNDO = "#331100"
COLOR_TEXT = "yellow"
COLOR_ORDER = "#00FF00"

class PlayerData:
    def __init__(self, templates_list):
        self.templates = templates_list # List gambar (Normal, Besar, Extra Besar)
        self.last_x = 0
        self.last_y = 0
        self.is_visible = False
        self.missing_frames = 0 
        self.is_dead = False
        self.urutan_ketemu = 0 

class TrackerV13_Fixed:
    def __init__(self):
        self.root = tk.Tk()
        self.screen_w = self.root.winfo_screenwidth()
        self.screen_h = self.root.winfo_screenheight()

        print(f"--- TRACKER V13 (FIXED CLICK LOGIC) ---")
        print(f"Klik Kiri: + Pertemuan")
        print(f"Klik Kanan: - Pertemuan (Undo)")
        print(f"Double Klik: Mati/Hidup")
        
        self.root.overrideredirect(True)
        self.root.wm_attributes("-topmost", True)
        self.root.wm_attributes("-transparentcolor", "black")
        self.root.geometry(f"{self.screen_w}x{self.screen_h}+0+0")
        
        self.canvas = tk.Canvas(self.root, width=self.screen_w, height=self.screen_h, bg="black", highlightthickness=0)
        self.canvas.pack()

        self.players = {} 
        self.action_history = [] 
        self.ui_hidden = False   
        self.is_tracking = False
        
        self.game_ke = 1
        self.ronde_ke = 1
        self.counter_lawan = 0 
        self.sct_main = mss.mss()
        
        self.check_game_id()
        self.setup_ui_buttons()

        # Kotak Merah (Aim Box)
        self.aim_box = self.canvas.create_rectangle(0, 0, CROP_WIDTH, CROP_HEIGHT, outline="red", width=2, tags="hideable_box")

        self.thread_track = Thread(target=self.process_tracking, daemon=True)
        self.thread_track.start()
        self.thread_input = Thread(target=self.input_loop, daemon=True)
        self.thread_input.start()
        
        # --- MOUSE BINDING (UPDATED) ---
        self.canvas.bind("<Button-1>", lambda e: self.on_canvas_click(e, "left"))
        self.canvas.bind("<Button-3>", lambda e: self.on_canvas_click(e, "right"))
        self.canvas.bind("<Double-Button-1>", lambda e: self.on_double_click(e)) # Tambahan Double Click

        self.root.mainloop()

    def check_game_id(self):
        i = 1
        while os.path.exists(f"{FILE_TXT_PREFIX}_{i}.txt"): i += 1
        self.game_ke = i

    def setup_ui_buttons(self):
        self.buttons = {} 
        btn_w, btn_h, gap = 45, 35, 5
        start_y = 100 
        
        # Hitung lebar total (7 Player + 4 Kontrol = 11 Tombol)
        total_width = (11 * (btn_w + gap)) 
        current_x = (self.screen_w // 2) - (total_width // 2)

        # 2. PLAYER 1-7 (HANYA SAMPAI 7)
        for i in range(1, 8):
            self.create_btn(i, str(i), current_x, start_y, BTN_COLOR, "cyan")
            current_x += btn_w + gap

        current_x += 10
        # Controls
        self.create_btn(99, "M", current_x, start_y, BTN_MONSTER, "#00FF00") # Monster
        current_x += btn_w + gap
        self.create_btn(88, "U", current_x, start_y, BTN_UNDO, "orange") # Undo
        current_x += btn_w + gap
        self.create_btn(0, "R", current_x, start_y, "#330000", "red") # Reset
        current_x += btn_w + gap
        self.create_btn(77, "-", current_x, start_y, "#222222", "white", hideable=False) # Hide

        # Info Labels (HANYA PESAN SINGKAT, RONDE DIHAPUS)
        self.label_msg = self.canvas.create_text(self.screen_w // 2, start_y + 50, text="", fill="yellow", font=("Arial", 10), tags="hideable")

    def create_btn(self, bid, txt, x, y, bg, fg, hideable=True):
        tag = (f"btn_{bid}", "hideable") if hideable else (f"btn_{bid}")
        tag_txt = (f"txt_{bid}", "hideable") if hideable else (f"txt_{bid}")
        self.buttons[bid] = (x, y, x+45, y+35)
        self.canvas.create_rectangle(x, y, x+45, y+35, fill=bg, outline=fg, width=2, tags=tag)
        self.canvas.create_text(x+22, y+17, text=txt, fill=fg, font=("Arial", 12, "bold"), tags=tag_txt)

    def input_loop(self):
        import mouse
        # Mapping 1-7 Keyboad (Hapus angka 8)
        KEYS_LOCK = {0x31:1, 0x32:2, 0x33:3, 0x34:4, 0x35:5, 0x36:6, 0x37:7} 
        last_press = 0 
        
        while True:
            # Update Aim Box
            if not self.is_tracking and not self.ui_hidden:
                mx, my = mouse.get_position()
                self.run_on_main(lambda c=[mx-CROP_WIDTH//2, my-CROP_HEIGHT//2, mx+CROP_WIDTH//2, my+CROP_HEIGHT//2]: self.canvas.coords(self.aim_box, *c))

            current_time = time.time()
            if current_time - last_press > 0.2:
                for code, num in KEYS_LOCK.items():
                    if win32api.GetAsyncKeyState(code) & 0x8000:
                        self.run_on_main(lambda n=num: self.lock_target(n))
                        last_press = current_time
                if win32api.GetAsyncKeyState(win32con.VK_F2) & 0x8000:
                    self.run_on_main(self.toggle_track)
                    last_press = current_time
                if win32api.GetAsyncKeyState(0x4D) & 0x8000: # M
                    self.run_on_main(self.log_monster)
                    last_press = current_time
                if win32api.GetAsyncKeyState(win32con.VK_ESCAPE) & 0x8000:
                    self.root.quit()
            time.sleep(0.02)

    def lock_target(self, number):
        if self.is_tracking: return
        import mouse
        x, y = mouse.get_position()
        
        # 1. Capture Area
        left = int(x - CROP_WIDTH // 2)
        top = int(y - CROP_HEIGHT // 2)
        bbox = {'top': top, 'left': left, 'width': CROP_WIDTH, 'height': CROP_HEIGHT}
        sct_img = self.sct_main.grab(bbox)
        img_np = np.array(sct_img)
        gray = cv2.cvtColor(img_np, cv2.COLOR_BGRA2GRAY)
        
        # 2. X-RAY (Hitam Putih)
        _, base_binary = cv2.threshold(gray, 160, 255, cv2.THRESH_BINARY)
        
        # 3. AUTO-ZOOM GENERATION (Membuat 3 ukuran font sekaligus)
        templates = []
        h, w = base_binary.shape
        
        # Varian 1: Normal (100%)
        templates.append(base_binary)
        
        # Varian 2: Zoom 10% (110%) - Untuk antisipasi font membesar
        zoom1 = cv2.resize(base_binary, (int(w * 1.1), int(h * 1.1)))
        templates.append(zoom1)
        
        # Varian 3: Zoom 15% (115%)
        zoom2 = cv2.resize(base_binary, (int(w * 1.15), int(h * 1.15)))
        templates.append(zoom2)

        # Simpan
        new_player = PlayerData(templates)
        new_player.last_x = x
        new_player.last_y = y
        self.players[number] = new_player
        
        self.canvas.delete(f"static_{number}")
        self.canvas.create_text(x - OFFSET_TEKS, y, text=str(number), fill=COLOR_TEXT, font=("Verdana", 12, "bold"), tag=f"static_{number}")
        self.show_msg(f"Target {number} LOCKED")

    def process_tracking(self):
        with mss.mss() as sct_thread:
            while True:
                if self.is_tracking and self.players:
                    # Scan Sisi Kanan Saja
                    start_x = self.screen_w - SIDEBAR_WIDTH
                    screen_bbox = {'top': 0, 'left': start_x, 'width': SIDEBAR_WIDTH, 'height': self.screen_h}
                    sct_img = sct_thread.grab(screen_bbox)
                    img_np = np.array(sct_img)
                    gray = cv2.cvtColor(img_np, cv2.COLOR_BGRA2GRAY)
                    
                    _, screen_binary = cv2.threshold(gray, 160, 255, cv2.THRESH_BINARY)

                    draw_data = [] 
                    
                    for pid, player in self.players.items():
                        if player.is_dead: continue
                        
                        best_val = 0
                        best_loc = (0, 0)
                        
                        # Cek semua ukuran template (Normal, Besar, Extra Besar)
                        for tmpl in player.templates:
                            if tmpl.shape[0] > screen_binary.shape[0] or tmpl.shape[1] > screen_binary.shape[1]: continue
                            res = cv2.matchTemplate(screen_binary, tmpl, cv2.TM_CCOEFF_NORMED)
                            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
                            
                            if max_val > best_val:
                                best_val = max_val
                                best_loc = max_loc
                        
                        # Logic Stabilizer
                        if best_val > THRESHOLD:
                            real_x = start_x + best_loc[0] + (CROP_WIDTH // 2)
                            real_y = best_loc[1] + (CROP_HEIGHT // 2)
                            
                            # Deadzone 4px (Anti-Getar)
                            if abs(real_x - player.last_x) > DEADZONE or abs(real_y - player.last_y) > DEADZONE:
                                # Smoothing Move
                                player.last_x = int(player.last_x * 0.6 + real_x * 0.4)
                                player.last_y = int(player.last_y * 0.6 + real_y * 0.4)
                                
                                # First time teleport
                                if player.last_x < 100:
                                    player.last_x = real_x
                                    player.last_y = real_y

                            player.missing_frames = 0 
                            player.is_visible = True
                        else:
                            player.missing_frames += 1
                            # Toleransi Cepat (15 frame = 0.7 detik)
                            # Biar kalau buka shop langsung hilang
                            if player.missing_frames < 15: player.is_visible = True
                            else: player.is_visible = False 

                        if player.is_visible:
                            draw_data.append((pid, player.last_x, player.last_y, player.urutan_ketemu))

                    self.run_on_main(lambda: self.draw_markers(draw_data))
                time.sleep(0.05) 

    def draw_markers(self, data):
        self.canvas.delete("markers") 
        for pid, x, y, urutan in data:
            self.canvas.create_text(x - OFFSET_TEKS, y, text=str(pid), fill=COLOR_TEXT, font=("Verdana", 14, "bold"), tags="markers")
            if urutan > 0:
                self.canvas.create_text(x - OFFSET_TEKS + 35, y, text=f"#{urutan}", fill=COLOR_ORDER, font=("Verdana", 14, "bold"), tags="markers")

    # --- UI EVENT HANDLERS ---
    def on_canvas_click(self, event, click_type):
        x, y = event.x, event.y
        for bid, coords in self.buttons.items():
            if coords[0] <= x <= coords[2] and coords[1] <= y <= coords[3]:
                if self.ui_hidden and bid != 77: continue
                if click_type == "left": self.handle_left(bid)
                elif click_type == "right": self.handle_right(bid)
                self.return_focus()
                return

    # HANDLE DOUBLE CLICK (Untuk MATIKAN PLAYER)
    def on_double_click(self, event):
        x, y = event.x, event.y
        for bid, coords in self.buttons.items():
            if coords[0] <= x <= coords[2] and coords[1] <= y <= coords[3]:
                if self.ui_hidden and bid != 77: continue
                if bid in self.players:
                    self.handle_kill(bid)
                return

    def return_focus(self):
        try:
            point = win32api.GetCursorPos()
            hwnd = win32gui.WindowFromPoint(point)
            if hwnd != win32gui.GetForegroundWindow(): pass
        except: pass

    def handle_left(self, bid):
        if bid == 0: self.reset_game()
        elif bid == 99: self.log_monster()
        elif bid == 88: self.undo_action()
        elif bid == 77: self.toggle_hide()
        elif bid in self.players:
            # KLIK KIRI 1x: MENAIKKAN PERTEMUAN
            p = self.players[bid]
            if p.is_dead: return
            
            self.counter_lawan += 1
            p.urutan_ketemu = self.counter_lawan
            self.canvas.itemconfig(f"btn_{bid}", outline="red", fill="#330000")
            self.canvas.itemconfig(f"txt_{bid}", fill="red")
            self.show_msg(f"VS P{bid}")
            self.ronde_ke += 1
            self.write_log(f"VS Player {bid} (#{self.counter_lawan})")

    def handle_right(self, bid):
        # KLIK KANAN 1x: MENURUNKAN PERTEMUAN (UNDO SPESIFIK)
        if bid in self.players:
            p = self.players[bid]
            # Hanya undo jika sudah pernah ketemu
            if p.urutan_ketemu > 0:
                p.urutan_ketemu = 0
                # Reset visual ke Cyan
                self.canvas.itemconfig(f"btn_{bid}", outline="cyan", fill=BTN_COLOR)
                self.canvas.itemconfig(f"txt_{bid}", fill="white")
                
                # Turunkan counter global
                if self.counter_lawan > 0:
                    self.counter_lawan -= 1
                
                # Turunkan ronde
                if self.ronde_ke > 1:
                    self.ronde_ke -= 1
                    
                self.show_msg(f"UNDO P{bid}")
                self.write_log(f"UNDO P{bid}")

    def handle_kill(self, bid):
        # DOUBLE CLICK: MATIKAN / HIDUPKAN
        if bid in self.players:
            p = self.players[bid]
            if p.is_dead:
                p.is_dead = False
                self.canvas.itemconfig(f"btn_{bid}", fill=BTN_COLOR)
                self.canvas.itemconfig(f"txt_{bid}", fill="white")
                self.show_msg(f"P{bid} HIDUP")
            else:
                p.is_dead = True
                # Jika dia sebelumnya ditandai ketemu, batalkan pertemuannya agar counter tidak kacau
                if p.urutan_ketemu > 0:
                    p.urutan_ketemu = 0
                    if self.counter_lawan > 0: self.counter_lawan -= 1
                    if self.ronde_ke > 1: self.ronde_ke -= 1
                
                self.canvas.itemconfig(f"btn_{bid}", fill="#444444")
                self.canvas.itemconfig(f"txt_{bid}", fill="gray")
                self.canvas.delete(f"static_{bid}")
                self.show_msg(f"P{bid} MATI")

    def toggle_track(self):
        if not self.players:
            self.show_msg("ERROR: Belum ada target!")
            return
        self.is_tracking = not self.is_tracking
        if self.is_tracking:
            self.canvas.itemconfig("hideable_box", state='hidden')
            self.canvas.itemconfig("hideable", state='hidden') # Hide UI saat tracking
            self.canvas.itemconfig(f"btn_77", state='normal') # Kecuali tombol hide
            self.canvas.itemconfig(f"txt_77", state='normal')
            for i in range(1, 8): self.canvas.delete(f"static_{i}") 
            
            with open(f"{FILE_TXT_PREFIX}_{self.game_ke}.txt", "w") as f:
                f.write(f"=== GAME {self.game_ke} START ===\n")
        else:
            self.show_msg("PAUSED")
            self.canvas.delete("markers")
            if not self.ui_hidden: self.canvas.itemconfig("hideable", state='normal')

    def toggle_hide(self):
        self.ui_hidden = not self.ui_hidden
        state = 'hidden' if self.ui_hidden else 'normal'
        self.canvas.itemconfig("hideable", state=state)
        self.canvas.itemconfig("txt_77", text="+" if self.ui_hidden else "-")

    def log_monster(self):
        self.action_history.append("monster")
        self.ronde_ke += 1
        self.show_msg("Monster Round")
        self.write_log("Monster Round")

    def undo_action(self):
        self.ronde_ke -= 1
        self.show_msg("UNDO")
        self.write_log("UNDO Action")

    def reset_game(self):
        self.game_ke += 1
        self.ronde_ke = 1
        self.counter_lawan = 0
        self.players = {}
        self.is_tracking = False
        self.canvas.delete("all")
        self.setup_ui_buttons()
        self.aim_box = self.canvas.create_rectangle(0, 0, CROP_WIDTH, CROP_HEIGHT, outline="red", width=2, tags="hideable_box")

    def run_on_main(self, func):
        self.root.after(0, func)

    def show_msg(self, text):
        self.canvas.itemconfig(self.label_msg, text=text)
        self.root.after(2000, lambda: self.canvas.itemconfig(self.label_msg, text=""))

    def write_log(self, text):
        try:
            with open(f"{FILE_TXT_PREFIX}_{self.game_ke}.txt", "a") as f:
                f.write(f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
        except: pass

if __name__ == "__main__":
    TrackerV13_Fixed()