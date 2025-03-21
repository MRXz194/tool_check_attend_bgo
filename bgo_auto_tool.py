from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading
import logging
import re
from concurrent.futures import ThreadPoolExecutor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('bgo_auto.log'),
        logging.StreamHandler()
    ]
)

class ClassAttendanceWorker:
    def __init__(self, student_ids, online_students, lesson_type):
        self.driver = None
        self.wait = None
        self.student_ids = student_ids
        self.online_students = online_students
        self.lesson_type = lesson_type
        self.failed_students = []
        
    def initialize_browser(self):
        # Only initialize if we don't already have a browser
        if self.driver is None:
            try:
                chrome_options = Options()
                chrome_options.add_argument('--start-maximized')
                chrome_options.add_argument('--log-level=3')
                self.driver = webdriver.Chrome(options=chrome_options)
                self.wait = WebDriverWait(self.driver, 20)
                self.driver.get('http://quanly.bgo.edu.vn/')
                self.driver.execute_script("document.body.style.zoom='67%'")
                logging.info("Browser initialized successfully")
                return True
            except Exception as e:
                logging.error(f"Browser initialization failed: {str(e)}")
                if self.driver:
                    self.driver.quit()
                self.driver = None
                return False
        else:
            # We already have a browser, just make sure we have a wait object
            if self.wait is None:
                self.wait = WebDriverWait(self.driver, 20)
            return True

    def process_student(self, student_id):
        try:
            logging.info(f"Processing student: {student_id}")
            
            find = self.wait.until(EC.presence_of_element_located((
                By.XPATH, '/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input'
            )))
            find.clear()
            find.send_keys(student_id)
            
            try:
                WebDriverWait(self.driver, 0.5).until(
                    EC.presence_of_element_located((
                        By.XPATH, "//div[@role='row' and contains(@class, 'ag-row') and not(contains(@class, 'ag-row-header'))]"
                    ))
                )
            except TimeoutException:
                raise Exception("Không tìm thấy học sinh trong danh sách")
            
            div_1 = self.wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//div[@col-id='arrivalStatus']")
            ))[1]
            option_1 = div_1.find_element(By.XPATH, '//option[@value="ON_TIME"]')
            option_1.click()
            time.sleep(1)
            
            div_2 = self.wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//div[@col-id='4']")
            ))[1]
            select_2 = div_2.find_element(By.XPATH, '//select')
            if self.lesson_type == "review":
                attendance_value = "2"
            else:
                attendance_value = "2" if student_id in self.online_students else "1"
            option_2 = select_2.find_elements(By.XPATH, f'//option[@value="{attendance_value}"]')[1]
            option_2.click()
            time.sleep(1)
            
            if self.lesson_type != "review":
                div_3 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='10']")
                ))[1]
                select_3 = div_3.find_element(By.XPATH, '//select')
                
                notebook_value = "1" if self.lesson_type == "theory" else "2"
                option_3 = select_3.find_elements(By.XPATH, f'//option[@value="{notebook_value}"]')[-1]
                option_3.click()
                time.sleep(2)
            
            return None
        except Exception as e:
            return (student_id, str(e))

    def process_class(self):
        should_quit_browser = False
        if not self.driver:
            if not self.initialize_browser():
                return [(None, "Không thể khởi động trình duyệt")]
            should_quit_browser = True  # Only quit if we created the browser here
            
        try:
            for student_id in self.student_ids:
                result = self.process_student(student_id)
                if result:
                    self.failed_students.append(result)
                    
        finally:
            if should_quit_browser and self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                    
        return self.failed_students

class StudentAttendanceChecker:
    def __init__(self):
        self.class_tabs = []
        self.tab_browsers = {}  # Store browser instances for each tab
        self.setup_gui()
        self.root.update()

    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Student Attendance Checker")
        self.root.geometry("800x600")
        
        self.root.attributes('-topmost', True)
        self.root.update()
        
        # Main container
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Top control panel
        control_panel = ttk.Frame(main_container)
        control_panel.pack(fill=tk.X, pady=(0, 10))
        
        # Add class button
        add_class_btn = ttk.Button(control_panel, text="Thêm lớp mới", command=self.add_class_tab)
        add_class_btn.pack(side=tk.LEFT, padx=5)
        
        # Process all button
        process_all_btn = ttk.Button(control_panel, text="Điểm danh tất cả", command=self.process_all_classes)
        process_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Report tab (always last)
        self.report_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.report_frame, text="Báo cáo lỗi")
        self.setup_report_tab()
        
        # Add first class tab
        self.add_class_tab()
        
        self.root.after(2000, lambda: self.root.attributes('-topmost', False))

    def create_class_tab(self, tab_id):
        class_frame = ttk.Frame(self.notebook, padding="10")
        
        # Top control bar for the tab
        tab_control = ttk.Frame(class_frame)
        tab_control.pack(fill=tk.X, pady=(0, 10))
        
        # Delete button for this tab
        delete_btn = ttk.Button(
            tab_control, 
            text="Xóa lớp này", 
            command=lambda: self.delete_class_tab(tab_id)
        )
        delete_btn.pack(side=tk.RIGHT)

        # Open Browser button
        open_browser_btn = ttk.Button(
            tab_control,
            text="Mở trình duyệt",
            command=lambda: self.open_browser_for_tab(tab_id)
        )
        open_browser_btn.pack(side=tk.LEFT, padx=5)

        # Take attendance button for this tab
        take_attendance_btn = ttk.Button(
            tab_control,
            text="Điểm danh lớp này",
            command=lambda: self.process_single_class(tab_id)
        )
        take_attendance_btn.pack(side=tk.LEFT, padx=5)
        
        # Lesson type selection
        lesson_frame = ttk.Frame(class_frame)
        lesson_frame.pack(fill=tk.X, pady=5)
        ttk.Label(lesson_frame, text="Loại tiết học:").pack(side=tk.LEFT, padx=5)
        lesson_type = tk.StringVar(value="theory")
        ttk.Radiobutton(lesson_frame, text="Lý thuyết", variable=lesson_type, value="theory").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(lesson_frame, text="Luyện đề", variable=lesson_type, value="practice").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(lesson_frame, text="Chữa đề", variable=lesson_type, value="review").pack(side=tk.LEFT, padx=5)
        
        # Online students entry
        online_frame = ttk.LabelFrame(class_frame, text="Học sinh học online", padding="5")
        online_frame.pack(fill=tk.X, pady=5)
        ttk.Label(online_frame, text="Nhập mã học sinh online (phân cách bằng dấu phẩy):").pack(pady=2)
        online_ids_entry = ttk.Entry(online_frame, width=50)
        online_ids_entry.pack(pady=2)
        
        # All students entry
        all_students_frame = ttk.LabelFrame(class_frame, text="Tất cả học sinh", padding="5")
        all_students_frame.pack(fill=tk.X, pady=5)
        ttk.Label(all_students_frame, text="Nhập tất cả mã học sinh (nên nhập 4 số*neu id nhiều số 0*, 3 số)(EX:111,222,...").pack(pady=2)
        student_ids_entry = ttk.Entry(all_students_frame, width=50)
        student_ids_entry.pack(pady=2)
        
        # Status label
        status_label = ttk.Label(class_frame, text="")
        status_label.pack(pady=5)
        
        return {
            'id': tab_id,
            'frame': class_frame,
            'lesson_type': lesson_type,
            'online_ids_entry': online_ids_entry,
            'student_ids_entry': student_ids_entry,
            'status_label': status_label
        }

    def add_class_tab(self):
        tab_id = len(self.class_tabs) + 1
        tab_info = self.create_class_tab(tab_id)
        
        # Insert before report tab
        self.notebook.insert(self.notebook.index(self.report_frame), tab_info['frame'], text=f"Lớp {tab_id}")
        self.class_tabs.append(tab_info)
        
        # Select the new tab
        self.notebook.select(tab_info['frame'])

    def delete_class_tab(self, tab_id):
        # Close browser if it exists
        if tab_id in self.tab_browsers:
            try:
                self.tab_browsers[tab_id].quit()
            except:
                pass
            self.tab_browsers.pop(tab_id)

        for i, tab in enumerate(self.class_tabs):
            if tab['id'] == tab_id:
                self.notebook.forget(tab['frame'])
                self.class_tabs.pop(i)
                break
        
        # Renumber remaining tabs
        for i, tab in enumerate(self.class_tabs, 1):
            tab['id'] = i
            tab_index = self.notebook.index(tab['frame'])
            self.notebook.tab(tab_index, text=f"Lớp {i}")

    def setup_report_tab(self):
        report_container = ttk.Frame(self.report_frame)
        report_container.pack(fill=tk.BOTH, expand=True)
        
        self.report_text = tk.Text(report_container, height=15, width=60)
        scrollbar = ttk.Scrollbar(report_container, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=scrollbar.set)
        
        self.report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.report_text.config(state=tk.DISABLED)

    def validate_student_id(self, student_id):
        student_id = student_id.strip()
        if not student_id.isdigit():
            return False, "Mã học sinh chỉ được chứa số"
        if len(student_id) < 2 or len(student_id) > 6:
            return False, "Độ dài mã học sinh không hợp lệ"
        return True, ""

    def validate_class_input(self, student_ids):
        if not student_ids:
            return False, "Vui lòng nhập ít nhất một mã học sinh"
            
        student_list = []
        for student_id in student_ids.split(','):
            student_id = student_id.strip()
            is_valid, error_msg = self.validate_student_id(student_id)
            if not is_valid:
                return False, f"Mã học sinh không hợp lệ '{student_id}': {error_msg}"
            student_list.append(student_id)
            
        return True, student_list

    def process_all_classes(self):
        if not self.class_tabs:
            messagebox.showwarning("Cảnh báo", "Không có lớp nào để điểm danh!")
            return
            
        all_failed_students = []
        
        # Validate all inputs first
        class_configs = []
        for tab in self.class_tabs:
            student_ids = tab['student_ids_entry'].get().strip()
            is_valid, result = self.validate_class_input(student_ids)
            if not is_valid:
                messagebox.showerror("Lỗi", f"Lỗi ở lớp {tab['id']}: {result}")
                return
                
            online_ids = tab['online_ids_entry'].get().strip()
            online_students = set(id.strip() for id in online_ids.split(',') if id.strip())
            
            # Get existing browser if available
            browser = None
            if tab['id'] in self.tab_browsers:
                try:
                    browser = self.tab_browsers[tab['id']]
                    # Check if browser is still responsive
                    browser.current_url
                except:
                    # Browser was closed by user
                    self.tab_browsers.pop(tab['id'])
                    browser = None
            
            class_configs.append({
                'student_ids': result,
                'online_students': online_students,
                'lesson_type': tab['lesson_type'].get(),
                'status_label': tab['status_label'],
                'browser': browser,
                'tab_id': tab['id']
            })
        
        # Process all classes in parallel
        with ThreadPoolExecutor(max_workers=len(class_configs)) as executor:
            futures = []
            
            for config in class_configs:
                worker = ClassAttendanceWorker(
                    config['student_ids'],
                    config['online_students'],
                    config['lesson_type']
                )
                
                # Set existing browser if available
                if config['browser']:
                    worker.driver = config['browser']
                    worker.wait = WebDriverWait(config['browser'], 20)
                
                future = executor.submit(worker.process_class)
                futures.append((future, config['status_label'], config['tab_id']))
                
            # Wait for all tasks to complete
            for future, status_label, tab_id in futures:
                try:
                    failed_students = future.result()
                    if failed_students:
                        # Add class information to each failed student
                        class_failed_students = [(student_id, error, tab_id) for student_id, error in failed_students]
                        all_failed_students.extend(class_failed_students)
                        status_label.config(text=f"Hoàn thành! Có {len(failed_students)} học sinh bị lỗi.")
                    else:
                        status_label.config(text="Hoàn thành điểm danh!")
                except Exception as e:
                    logging.error(f"Error in worker thread: {str(e)}")
                    status_label.config(text="Lỗi rùi huhu")
                    all_failed_students.append((None, str(e), tab_id))
                    
        # Update report
        self.update_report(all_failed_students)
        if all_failed_students:
            self.notebook.select(self.report_frame)

    def process_single_class(self, tab_id):
        tab = next((tab for tab in self.class_tabs if tab['id'] == tab_id), None)
        if not tab:
            return

        student_ids = tab['student_ids_entry'].get().strip()
        is_valid, result = self.validate_class_input(student_ids)
        if not is_valid:
            messagebox.showerror("Lỗi", result)
            return

        online_ids = tab['online_ids_entry'].get().strip()
        online_students = set(id.strip() for id in online_ids.split(',') if id.strip())

        # Create worker and set its browser if we have one
        worker = ClassAttendanceWorker(
            result,
            online_students,
            tab['lesson_type'].get()
        )

        # Use existing browser if available
        if tab_id in self.tab_browsers:
            try:
                driver = self.tab_browsers[tab_id]
                # Check if browser is still responsive
                driver.current_url
                # Set the existing browser
                worker.driver = driver
                worker.wait = WebDriverWait(driver, 20)
            except:
                # Browser was closed by user
                self.tab_browsers.pop(tab_id)

        try:
            failed_students = worker.process_class()
            if failed_students:
                # Add class information to each failed student
                class_failed_students = [(student_id, error, tab_id) for student_id, error in failed_students]
                tab['status_label'].config(text=f"Hoàn thành! Có {len(failed_students)} học sinh bị lỗi.")
                self.update_report(class_failed_students)
                self.notebook.select(self.report_frame)
            else:
                tab['status_label'].config(text="Hoàn thành điểm danh!")
        except Exception as e:
            logging.error(f"Error processing class: {str(e)}")
            tab['status_label'].config(text="Lỗi rùi huhu")
            self.update_report([(None, str(e), tab_id)])
            self.notebook.select(self.report_frame)

    def open_browser_for_tab(self, tab_id):
        if tab_id in self.tab_browsers:
            try:
                # Check if browser is still responsive
                self.tab_browsers[tab_id].current_url
                messagebox.showinfo("Thông báo", "Trình duyệt đã được mở cho lớp này rồi!")
                return
            except:
                # Browser was closed, remove it from our tracking
                self.tab_browsers.pop(tab_id)

        def open_browser_thread():
            try:
                chrome_options = Options()
                chrome_options.add_argument('--start-maximized')
                chrome_options.add_argument('--log-level=3')
                driver = webdriver.Chrome(options=chrome_options)
                self.tab_browsers[tab_id] = driver
                driver.get('http://quanly.bgo.edu.vn/')
                driver.execute_script("document.body.style.zoom='67%'")
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã mở trình duyệt. Vui lòng điều hướng đến trang điểm danh."))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Không thể mở trình duyệt: {str(e)}"))

        threading.Thread(target=open_browser_thread, daemon=True).start()

    def update_report(self, failed_students=None):
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        
        if not failed_students:
            self.report_text.insert(tk.END, "Không có học sinh nào bị lỗi trong lần chạy gần nhất.")
        else:
            self.report_text.insert(tk.END, "Danh sách học sinh bị lỗi:\n\n")
            # Group errors by class
            errors_by_class = {}
            for student_id, error, class_id in failed_students:
                if class_id not in errors_by_class:
                    errors_by_class[class_id] = []
                errors_by_class[class_id].append((student_id, error))
            
            # Display errors grouped by class
            for class_id in sorted(errors_by_class.keys()):
                self.report_text.insert(tk.END, f"=== Lớp {class_id} ===\n\n")
                for student_id, error in errors_by_class[class_id]:
                    if student_id:  # Skip browser initialization errors
                        self.report_text.insert(tk.END, f"Mã học sinh: {student_id}\n")
                        self.report_text.insert(tk.END, f"Lỗi: {error}\n")
                        self.report_text.insert(tk.END, "-" * 50 + "\n")
                    else:
                        self.report_text.insert(tk.END, f"Lỗi khởi tạo: {error}\n")
                        self.report_text.insert(tk.END, "-" * 50 + "\n")
                self.report_text.insert(tk.END, "\n")
        
        self.report_text.config(state=tk.DISABLED)

    def on_closing(self):
        # Close all browsers
        for driver in self.tab_browsers.values():
            try:
                driver.quit()
            except:
                pass
        logging.info("Shutting down application")
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

if __name__ == "__main__":
    app = StudentAttendanceChecker()
    app.run()
