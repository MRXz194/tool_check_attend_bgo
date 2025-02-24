from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import tkinter as tk
from tkinter import ttk, messagebox
import time

class StudentAttendanceChecker:
    def __init__(self):
        self.driver = None
        self.online_students = set()  # store onl id
        self.failed_students = []  # Store fail student id and their erors
        self.initialize_browser()  # Start browser 
        self.setup_gui()
        
    def initialize_browser(self):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--start-maximized')
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            
            # Open BGO web
            self.driver.get('http://quanly.bgo.edu.vn/')
            self.driver.execute_script("document.body.style.zoom='67%'")
            
            # message box
            messagebox.showinfo(
                "Hướng dẫn", 
                "Vui lòng đăng nhập và chọn buổi học trên trình duyệt.\n"
                "Sau đó quay lại tool và nhập danh sách học sinh để điểm danh."
            )
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể khởi động trình duyệt: {str(e)}")
            if self.driver:
                self.driver.quit()
            self.driver = None
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Student Attendance Checker")
        self.root.geometry("600x400")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Main tab
        main_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(main_frame, text="Điểm danh")
        
        # Report tab
        self.report_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.report_frame, text="Báo cáo lỗi")
        
        # Setup report tab
        self.setup_report_tab()
        
        # Main tab content
        # chon kieu tiet hoc
        lesson_frame = ttk.Frame(main_frame)
        lesson_frame.pack(fill=tk.X, pady=5)
        ttk.Label(lesson_frame, text="Loại tiết học:").pack(side=tk.LEFT, padx=5)
        self.lesson_type = tk.StringVar(value="theory")
        ttk.Radiobutton(lesson_frame, text="Lý thuyết", variable=self.lesson_type, value="theory").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(lesson_frame, text="Luyện bài", variable=self.lesson_type, value="practice").pack(side=tk.LEFT, padx=5)
        
        # hs onl entry
        online_frame = ttk.LabelFrame(main_frame, text="Học sinh học online", padding="5")
        online_frame.pack(fill=tk.X, pady=5)
        ttk.Label(online_frame, text="Nhập mã học sinh online (phân cách bằng dấu phẩy):").pack(pady=2)
        self.online_ids_entry = ttk.Entry(online_frame, width=50)
        self.online_ids_entry.pack(pady=2)
        
        # All hs entry
        all_students_frame = ttk.LabelFrame(main_frame, text="Tất cả học sinh", padding="5")
        all_students_frame.pack(fill=tk.X, pady=5)
        ttk.Label(all_students_frame, text="Nhập mã học sinh (phân cách bằng dấu phẩy, nên nhập 4 số)(EX:0053,0012,...):").pack(pady=2)
        self.student_ids_entry = ttk.Entry(all_students_frame, width=50)
        self.student_ids_entry.pack(pady=2)
        
        # button
        ttk.Button(main_frame, text="Điểm danh", command=self.process_students).pack(pady=10)
        
        # Status lab
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=5)
        
    def setup_report_tab(self):
        
        report_container = ttk.Frame(self.report_frame)
        report_container.pack(fill=tk.BOTH, expand=True)
        
        self.report_text = tk.Text(report_container, height=15, width=60)
        scrollbar = ttk.Scrollbar(report_container, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=scrollbar.set)
        
        self.report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        
        self.report_text.config(state=tk.DISABLED)
        
    def update_report(self):
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        
        if not self.failed_students:
            self.report_text.insert(tk.END, "Không có học sinh nào bị lỗi trong lần chạy gần nhất.")
        else:
            self.report_text.insert(tk.END, "Danh sách học sinh bị lỗi:\n\n")
            for student_id, error in self.failed_students:
                self.report_text.insert(tk.END, f"Mã học sinh: {student_id}\n")
                self.report_text.insert(tk.END, f"Lỗi: {error}\n")
                self.report_text.insert(tk.END, "-" * 50 + "\n")
        
        self.report_text.config(state=tk.DISABLED)
        
    def process_students(self):
        # Clear prev fail student
        self.failed_students = []
        
        # Get student ids
        student_ids = self.student_ids_entry.get().strip()
        if not student_ids:
            messagebox.showerror("Lỗi", "Vui lòng nhập ít nhất một mã học sinh")
            return
            
        # Get online hs id
        online_ids = self.online_ids_entry.get().strip()
        self.online_students = set(id.strip() for id in online_ids.split(',') if id.strip())
            
        student_list = [id.strip() for id in student_ids.split(',')]
        self.status_label.config(text="Đang xử lý... ")
        self.root.update()
        
        try:
            # Check if browser is still running
            if not self.driver:
                self.initialize_browser()
                if not self.driver:  # If initialization failed
                    return
                
            self.process_with_selenium(student_list)
            
            if self.failed_students:
                self.status_label.config(text=f"Hoàn thành! Có {len(self.failed_students)} học sinh bị lỗi.")
                self.notebook.select(1)  # Switch to report tab
            else:
                self.status_label.config(text="Hoàn thành điểm danh!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")
            self.status_label.config(text="lỗi rùi huhu")
            if self.driver:
                self.driver.save_screenshot("error_screenshot.png")
                self.driver.quit()
                self.driver = None
                
        # Update report tab
        self.update_report()
            
    def process_with_selenium(self, student_ids):
        for student_id in student_ids:
            try:
                # Update status for curent student
                self.status_label.config(text=f"Đang điểm danh học sinh: {student_id}")
                self.root.update()
                
                # tim va enter student id
                find = self.wait.until(EC.presence_of_element_located((
                    By.XPATH, '/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input'
                )))
                find.clear()
                find.send_keys(student_id)
                
                # Check if student exists (look row) - Optimized fix 
                try:
                    #  shorter timeout and  specific XPath optimize tuyet voi luon
                    WebDriverWait(self.driver, 0.5).until(
                        EC.presence_of_element_located((
                            By.XPATH, "//div[@role='row' and contains(@class, 'ag-row') and not(contains(@class, 'ag-row-header'))]"
                        ))
                    )
                except TimeoutException:
                    raise Exception("Không tìm thấy học sinh trong danh sách")
                
                # Set den dung h
                div_1 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='arrivalStatus']")
                ))[1]
                option_1 = div_1.find_element(By.XPATH, '//option[@value="ON_TIME"]')
                option_1.click()
                time.sleep(1)
                
                # set attend base onl or off
                div_2 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='4']")
                ))[1]
                select_2 = div_2.find_element(By.XPATH, '//select')
                # value="2" onl, value="1"  offline
                attendance_value = "2" if student_id in self.online_students else "1"
                option_2 = select_2.find_elements(By.XPATH, f'//option[@value="{attendance_value}"]')[1]
                option_2.click()
                time.sleep(1)
                
                # set vo ghi dua tren lesson type
                div_3 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='10']")
                ))[1]
                select_3 = div_3.find_element(By.XPATH, '//select')
                
                # value=1 xho ly thuyet, value=2 cho luyen de
                notebook_value = "1" if self.lesson_type.get() == "theory" else "2"
                option_3 = select_3.find_elements(By.XPATH, f'//option[@value="{notebook_value}"]')[-1]
                option_3.click()
                time.sleep(2)
                
            except Exception as e:
                # Add failed student to list and continue 
                self.failed_students.append((student_id, str(e)))
                continue
    
    def on_closing(self):
        if self.driver:
            self.driver.quit()
        self.root.destroy()
            
    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

if __name__ == "__main__":
    app = StudentAttendanceChecker()
    app.run()