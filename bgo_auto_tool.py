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
        self.root.geometry("400x200")
        
        
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Enter Student IDs (comma-separated)(EX:005,123,001,...):").pack(pady=5)
        self.student_ids_entry = ttk.Entry(frame, width=50)
        self.student_ids_entry.pack(pady=5)
        
        ttk.Button(frame, text="Process Students", command=self.process_students).pack(pady=10)
        
        self.status_label = ttk.Label(frame, text="")
        self.status_label.pack(pady=5)
        
    def process_students(self):
        student_ids = self.student_ids_entry.get().strip()
        if not student_ids:
            messagebox.showerror("Error", "Please enter at least one student ID")
            return
            
        student_list = [id.strip() for id in student_ids.split(',')]
        self.status_label.config(text="Processing... ")
        self.root.update()
        
        try:
            if not self.driver:
                chrome_options = Options()
                chrome_options.add_argument('--start-maximized')
                self.driver = webdriver.Chrome(options=chrome_options)
                self.wait = WebDriverWait(self.driver, 20)
                
                # tuong tac web
                self.driver.get('http://quanly.bgo.edu.vn/')
                time.sleep(60)  # time doi load dau
                self.driver.execute_script("document.body.style.zoom='67%'")
            
            self.process_with_selenium(student_list)
            self.status_label.config(text="Done")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.config(text="Error occurred during processing")
            if self.driver:
                self.driver.save_screenshot("error_screenshot.png")
                self.driver.quit()
                self.driver = None
            
    def process_with_selenium(self, student_ids):
        for student_id in student_ids:
            try:
                # cap nhat status cho hoc sinh hien tai 
                self.status_label.config(text=f"Processing student ID: {student_id}")
                self.root.update()
                
                # tim va nhap vao student id
                find = self.wait.until(EC.presence_of_element_located((
                    By.XPATH, '/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input'
                )))
                find.clear()
                find.send_keys(student_id)
                time.sleep(2)
                
                # Set dung gio
                div_1 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='arrivalStatus']")
                ))[1]
                option_1 = div_1.find_element(By.XPATH, '//option[@value="ON_TIME"]')
                option_1.click()
                time.sleep(1)
                
                # Set offline 
                div_2 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='4']")
                ))[1]
                select_2 = div_2.find_element(By.XPATH, '//select')
                option_2 = select_2.find_elements(By.XPATH, '//option[@value="1"]')[1]
                option_2.click()
                time.sleep(1)
                
                # Set vo hoan hao
                div_3 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='10']")
                ))[1]
                select_3 = div_3.find_element(By.XPATH, '//select')
                option_3 = select_3.find_elements(By.XPATH, '//option[@value="1"]')[-1]
                option_3.click()
                time.sleep(2)
                
            except Exception as e:
                raise Exception(f"Error processing student {student_id}: {str(e)}")
    
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