import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyautogui
import pyperclip
import time
import threading
import os

ABSENT_LIST_FILE = "danh_sach_vang_zalo.txt"

class ZaloForwardAutomationTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Zalo Auto Forward Mouse-Click by MRX :>")
        self.root.geometry("450x650")
        self.root.configure(bg='#f0f0f0')
        
        # Always on top so it doesn't get hidden
        self.root.attributes('-topmost', True)
        
        # Center the window
        self.center_window(450, 650)
        
        self.running = False
        self.search_pos = None
        self.tick_pos = None
        
        self.create_widgets()
        self.auto_load_absent_list()

    def auto_load_absent_list(self):
        if os.path.exists(ABSENT_LIST_FILE):
            try:
                with open(ABSENT_LIST_FILE, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if content:
                        # Count the number of lines/codes
                        lines = [line for line in content.split('\n') if line.strip()]
                        count = len(lines)
                        
                        self.group_text.insert(tk.END, content)
                        self.update_status(f"Đã tự động tải {count} mã số từ BGO!", "green")
            except:
                pass

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        # Instructions
        instruction_frame = tk.Frame(self.root, bg='#f0f0f0')
        instruction_frame.pack(pady=10, padx=15, fill=tk.X)
        
        tk.Label(instruction_frame, text="HƯỚNG DẪN SỬ DỤNG (BẢN CHUỘT):", font=('Arial', 10, 'bold'), bg='#f0f0f0').pack(anchor='w')
        instructions = (
            "1. Lấy vị trí: Bấm lấy tọa độ Ô TÌM KIẾM rồi ĐỂ YÊN CHUỘT ở đó 3s.\n"
            "2. Lấy vị trí: Bấm lấy tọa độ Ô TICK kết quả rồi ĐỂ YÊN CHUỘT 3s.\n"
            "3. Dán danh sách Tên nhóm/người nhận vào ô bên dưới.\n"
            "4. Mở cửa sổ Zalo Chia sẻ ra, rồi bấm Bắt Đầu Chạy."
        )
        tk.Label(instruction_frame, text=instructions, font=('Arial', 9), justify=tk.LEFT, bg='#f0f0f0').pack(anchor='w', pady=(5,0))

        # Coordinate config
        coord_frame = tk.LabelFrame(self.root, text="Thiết lập vị trí chuột", bg='#f0f0f0', font=('Arial', 9, 'bold'))
        coord_frame.pack(pady=5, padx=15, fill=tk.X)
        
        self.btn_search_pos = tk.Button(coord_frame, text="1. Bấm để lấy vị trí Ô TÌM KIẾM (Chờ 3s)", command=lambda: self.get_pos('search'))
        self.btn_search_pos.pack(pady=5, padx=5, fill=tk.X)
        lbl_s_text = "Tọa độ Ô TÌM KIẾM: Chưa có" if not self.search_pos else f"Tọa độ Ô TÌM KIẾM: Đã lưu (X={self.search_pos[0]}, Y={self.search_pos[1]})"
        self.lbl_search_pos = tk.Label(coord_frame, text=lbl_s_text, fg='red' if not self.search_pos else 'green', bg='#f0f0f0')
        self.lbl_search_pos.pack(pady=2)
        
        self.btn_tick_pos = tk.Button(coord_frame, text="2. Bấm để lấy vị trí Ô TICK (Chờ 3s)", command=lambda: self.get_pos('tick'))
        self.btn_tick_pos.pack(pady=5, padx=5, fill=tk.X)
        lbl_t_text = "Tọa độ Ô TICK: Chưa có" if not self.tick_pos else f"Tọa độ Ô TICK: Đã lưu (X={self.tick_pos[0]}, Y={self.tick_pos[1]})"
        self.lbl_tick_pos = tk.Label(coord_frame, text=lbl_t_text, fg='red' if not self.tick_pos else 'green', bg='#f0f0f0')
        self.lbl_tick_pos.pack(pady=2)

        # Text input for groups
        tk.Label(self.root, text="Danh sách Tên (Mỗi dòng 1 tên):", font=('Arial', 9, 'bold'), bg='#f0f0f0').pack(anchor='w', padx=15, pady=(5, 0))
        self.group_text = scrolledtext.ScrolledText(self.root, width=50, height=10, font=('Arial', 9))
        self.group_text.pack(pady=2, padx=15)
        
        # Delay setting
        delay_frame = tk.Frame(self.root, bg='#f0f0f0')
        delay_frame.pack(pady=5, padx=15, fill=tk.X)
        tk.Label(delay_frame, text="Thời gian chờ load kết quả (giây):", font=('Arial', 9), bg='#f0f0f0').pack(side=tk.LEFT)
        self.delay_var = tk.StringVar(value="1.0")
        tk.Entry(delay_frame, textvariable=self.delay_var, width=5).pack(side=tk.LEFT, padx=5)

        # Status Label
        self.status_label = tk.Label(self.root, text="Trạng thái: Sẵn sàng", font=('Arial', 10, 'bold'), fg='blue', bg='#f0f0f0')
        self.status_label.pack(pady=5)

        # Start Button
        self.start_btn = tk.Button(self.root, text="🚀 BẮT ĐẦU CHẠY (Chờ 3s)", font=('Arial', 11, 'bold'), bg='#4CAF50', fg='white', command=self.start_automation_thread, height=2)
        self.start_btn.pack(pady=5, padx=15, fill=tk.X)
        
        # Stop Button
        self.stop_btn = tk.Button(self.root, text="🛑 DỪNG KHẨN CẤP", font=('Arial', 10, 'bold'), bg='#f44336', fg='white', command=self.stop_automation, state=tk.DISABLED)
        self.stop_btn.pack(pady=5, padx=15, fill=tk.X)

    def get_pos(self, target):
        def grab():
            for i in range(3, 0, -1):
                self.update_status(f"Hãy để mũi tên lên {target} trong {i}s...", "orange")
                time.sleep(1)
            x, y = pyautogui.position()
            if target == 'search':
                self.search_pos = (x, y)
                self.lbl_search_pos.config(text=f"Tọa độ Ô TÌM KIẾM: X={x}, Y={y}", fg='green')
            else:
                self.tick_pos = (x, y)
                self.lbl_tick_pos.config(text=f"Tọa độ Ô TICK: X={x}, Y={y}", fg='green')
            self.update_status("Sẵn sàng", "blue")
        
        threading.Thread(target=grab, daemon=True).start()

    def start_automation_thread(self):
        if self.running:
            return
            
        if not self.search_pos or not self.tick_pos:
            messagebox.showwarning("Cảnh báo", "Bạn chưa lấy tọa độ của Ô tìm kiếm hoặc Ô tick!")
            return
            
        groups = self.group_text.get("1.0", tk.END).strip().split('\n')
        groups = [g.strip() for g in groups if g.strip()]
        
        if not groups:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập danh sách nhóm!")
            return
            
        try:
            search_delay = float(self.delay_var.get())
        except ValueError:
            messagebox.showwarning("Cảnh báo", "Thời gian chờ không hợp lệ!")
            return

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.btn_search_pos.config(state=tk.DISABLED)
        self.btn_tick_pos.config(state=tk.DISABLED)
        
        # Run in a separate thread so UI doesn't freeze
        thread = threading.Thread(target=self.run_automation_loop, args=(groups, search_delay))
        thread.daemon = True
        thread.start()

    def stop_automation(self):
        self.running = False
        self.status_label.config(text="Trạng thái: Đã dừng", fg='red')
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.btn_search_pos.config(state=tk.NORMAL)
        self.btn_tick_pos.config(state=tk.NORMAL)

    def update_status(self, text, color='blue'):
        self.status_label.config(text=text, fg=color)
        self.root.update()

    def run_automation_loop(self, groups, search_delay):
        try:
            # 3 Seconds countdown
            for i in range(3, 0, -1):
                if not self.running: return
                self.update_status(f"Chuẩn bị chạy sau {i} giây...", "orange")
                time.sleep(1)
                
            self.update_status("Đang chạy automation...", "green")
            
            for index, group in enumerate(groups, 1):
                if not self.running:
                    break
                    
                self.update_status(f"Đang xử lý ({index}/{len(groups)}): {group}", "green")
                
                # 1. Click Search Box
                pyautogui.click(x=self.search_pos[0], y=self.search_pos[1])
                time.sleep(0.1)
                
                # Clear existing text just in case
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.press('backspace')
                time.sleep(0.1)
                
                # 2. Copy and Paste name into Zalo search box
                pyperclip.copy(group)
                pyautogui.hotkey('ctrl', 'v')
                
                # 3. Wait for Zalo to render search results securely
                time.sleep(search_delay)
                
                # 4. Click the exact position of the Tick box
                pyautogui.click(x=self.tick_pos[0], y=self.tick_pos[1])
                time.sleep(0.2)
                
            if self.running:
                self.update_status("HOÀN THÀNH! Vui lòng tự bấm nút [Chia sẻ].", "blue")
                
        except Exception as e:
            self.update_status(f"Lỗi: {str(e)}", "red")
            
        finally:
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            self.root.after(0, lambda: self.btn_search_pos.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_tick_pos.config(state=tk.NORMAL))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ZaloForwardAutomationTool()
    app.run()
