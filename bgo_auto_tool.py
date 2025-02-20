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
        self.setup_gui()
        self.driver = None
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Student Attendance Checker")
        self.root.geometry("500x250")
        
        
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # chon kieu tiet hoc
        lesson_frame = ttk.Frame(frame)
        lesson_frame.pack(fill=tk.X, pady=5)
        ttk.Label(lesson_frame, text="Loại tiết học:").pack(side=tk.LEFT, padx=5)
        self.lesson_type = tk.StringVar(value="theory")
        ttk.Radiobutton(lesson_frame, text="Lý thuyết", variable=self.lesson_type, value="theory").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(lesson_frame, text="Luyện bài", variable=self.lesson_type, value="practice").pack(side=tk.LEFT, padx=5)
        
        # Student IDs entryyyyy
        ttk.Label(frame, text="Nhập mã học sinh (phân cách bằng dấu phẩy)(EX:053,001,123,...):").pack(pady=5)
        self.student_ids_entry = ttk.Entry(frame, width=50)
        self.student_ids_entry.pack(pady=5)
        
        # Process button
        ttk.Button(frame, text="Điểm danh", command=self.process_students).pack(pady=10)
        
        # Status label
        self.status_label = ttk.Label(frame, text="")
        self.status_label.pack(pady=5)
        
    def process_students(self):
        student_ids = self.student_ids_entry.get().strip()
        if not student_ids:
            messagebox.showerror("Lỗi", "Vui lòng nhập ít nhất một mã học sinh")
            return
            
        student_list = [id.strip() for id in student_ids.split(',')]
        self.status_label.config(text="Đang xử lý... ")
        self.root.update()
        
        try:
            if not self.driver:
                chrome_options = Options()
                chrome_options.add_argument('--start-maximized')
                self.driver = webdriver.Chrome(options=chrome_options)
                self.wait = WebDriverWait(self.driver, 20)
                
                # Navigate to the website
                self.driver.get('http://quanly.bgo.edu.vn/')
                time.sleep(60)  # Wait for initial load
                self.driver.execute_script("document.body.style.zoom='67%'")
            
            self.process_with_selenium(student_list)
            self.status_label.config(text="Hoàn thành điểm danh!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")
            self.status_label.config(text="lỗi rùi huhu")
            if self.driver:
                self.driver.save_screenshot("error_screenshot.png")
                self.driver.quit()
                self.driver = None
            
    def process_with_selenium(self, student_ids):
        for student_id in student_ids:
            try:
                # Update status for current student
                self.status_label.config(text=f"Đang điểm danh học sinh: {student_id}")
                self.root.update()
                
                # Find and enter student ID
                find = self.wait.until(EC.presence_of_element_located((
                    By.XPATH, '/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input'
                )))
                find.clear()
                find.send_keys(student_id)
                time.sleep(2)
                
                # Set arrival status (ON_TIME)
                div_1 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='arrivalStatus']")
                ))[1]
                option_1 = div_1.find_element(By.XPATH, '//option[@value="ON_TIME"]')
                option_1.click()
                time.sleep(1)
                
                # Set offline attendance
                div_2 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='4']")
                ))[1]
                select_2 = div_2.find_element(By.XPATH, '//select')
                option_2 = select_2.find_elements(By.XPATH, '//option[@value="1"]')[1]
                option_2.click()
                time.sleep(1)
                
                # Set notebook value based on lesson type
                div_3 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='10']")
                ))[1]
                select_3 = div_3.find_element(By.XPATH, '//select')
                
                # Choose value=1 for theory lessons, value=2 for practice lessons
                notebook_value = "1" if self.lesson_type.get() == "theory" else "2"
                option_3 = select_3.find_elements(By.XPATH, f'//option[@value="{notebook_value}"]')[-1]
                option_3.click()
                time.sleep(2)
                
            except Exception as e:
                raise Exception(f"Lỗi khi xử lý học sinh {student_id}: {str(e)}")
    
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