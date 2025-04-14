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
from PIL import Image, ImageTk
import requests
from io import BytesIO

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
                
                # Add page refresh detection
                self.setup_refresh_detection()
                
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
                
            # Check if we need to re-navigate to the attendance page
            try:
                current_url = self.driver.current_url
                if "class-attendances" not in current_url:
                    logging.info("Browser not on attendance page, redirecting...")
                    self.driver.get('http://quanly.bgo.edu.vn/')
                    # Add refresh detection again after navigation
                    self.setup_refresh_detection()
            except:
                logging.warning("Could not check browser URL, continuing anyway")
                
            return True
            
    def setup_refresh_detection(self):
        """Add JavaScript to detect page refreshes and maintain state"""
        try:
            # Add a flag to check if the page was refreshed
            js_code = """
            window.wasRefreshed = false;
            window.addEventListener('beforeunload', function() {
                window.wasRefreshed = true;
            });
            """
            self.driver.execute_script(js_code)
            logging.info("Refresh detection script added")
        except Exception as e:
            logging.error(f"Failed to add refresh detection: {str(e)}")

    def check_for_refresh(self):
        """Check if the page was refreshed and handle accordingly"""
        try:
            was_refreshed = self.driver.execute_script("return window.wasRefreshed === true")
            if was_refreshed:
                logging.info("Page refresh detected, re-initializing elements")
                # Reset the flag
                self.driver.execute_script("window.wasRefreshed = false")
                # Wait for page to load after refresh
                time.sleep(2)
                # Re-setup refresh detection
                self.setup_refresh_detection()
                return True
            return False
        except Exception as e:
            logging.error(f"Error checking for refresh: {str(e)}")
            return False

    def process_student(self, student_id):
        try:
            logging.info(f"Processing student: {student_id}")
            
            # Check if page was refreshed and handle it
            self.check_for_refresh()
            
            # Try to find the search input with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    find = self.wait.until(EC.presence_of_element_located((
                        By.XPATH, '/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input'
                    )))
                    break
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    logging.warning(f"Attempt {attempt+1}/{max_retries} to find search box failed, retrying...")
                    time.sleep(1)
                    # Check if we need to handle a refresh
                    if self.check_for_refresh():
                        time.sleep(1)  # Wait for page to stabilize after refresh
            
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
            time.sleep(0.5)  # Reduced sleep time to improve responsiveness
            
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
            time.sleep(0.5)  # Reduced sleep time
            
            if self.lesson_type != "review":
                div_3 = self.wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@col-id='10']")
                ))[1]
                select_3 = div_3.find_element(By.XPATH, '//select')
                
                notebook_value = "1" if self.lesson_type == "theory" else "2"
                option_3 = select_3.find_elements(By.XPATH, f'//option[@value="{notebook_value}"]')[-1]
                option_3.click()
                time.sleep(1)  # Reduced sleep time
            
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
            for i, student_id in enumerate(self.student_ids):
                result = self.process_student(student_id)
                if result:
                    self.failed_students.append(result)
                
                # Yield progress after each student for UI updates
                yield i + 1, len(self.student_ids)
                    
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
        self.root.geometry("1000x700")
        
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
        
        # About button in top right
        about_btn = ttk.Button(control_panel, text="About", command=self.show_about)
        about_btn.pack(side=tk.RIGHT, padx=5)
        
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

    def show_about(self):
        about_text = """BGO Auto Tool
Version 1.0

Author: M1xz0
GitHub: https://github.com/MRXz194

© 2025 All Rights Reserved"""
        messagebox.showinfo("About", about_text)

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

        # Rename button for this tab
        rename_btn = ttk.Button(
            tab_control,
            text="Đổi tên lớp",
            command=lambda: self.rename_class_tab(tab_id)
        )
        rename_btn.pack(side=tk.RIGHT, padx=5)

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
        
        # All students entry - improved interface
        all_students_frame = ttk.LabelFrame(class_frame, text="Tất cả học sinh", padding="10")
        all_students_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        entry_frame = ttk.Frame(all_students_frame)
        entry_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(entry_frame, text="Nhập mã học sinh:", width=15).pack(side=tk.LEFT, padx=5)
        student_id_var = tk.StringVar()
        student_id_entry = ttk.Entry(entry_frame, textvariable=student_id_var, width=20, font=('Arial', 12))
        student_id_entry.pack(side=tk.LEFT, padx=5)
        
        # Enter key binding to add ID
        student_id_entry.bind("<Return>", lambda e: add_id())
        
        # ID status label for showing validation errors
        id_status_var = tk.StringVar()
        id_status_label = ttk.Label(entry_frame, textvariable=id_status_var, foreground="red", font=('Arial', 10))
        id_status_label.pack(side=tk.LEFT, padx=5)
        
        # List to display all entered IDs
        list_frame = ttk.Frame(all_students_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Student ID list display with scrollbar
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        ttk.Label(list_container, text="Danh sách mã học sinh đã nhập:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        
        id_listbox = tk.Listbox(list_container, height=10, width=25, font=('Arial', 12))
        id_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=id_listbox.yview)
        id_listbox.configure(yscrollcommand=id_scrollbar.set)
        id_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        id_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Stats display
        stats_frame = ttk.Frame(list_frame)
        stats_frame.pack(fill=tk.Y, side=tk.RIGHT, padx=15)
        
        ttk.Label(stats_frame, text="Thống kê:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Create variables for statistics
        total_count_var = tk.StringVar(value="Tổng số ID: 0")
        digits_3_count_var = tk.StringVar(value="ID 3 chữ số: 0")
        digits_4_count_var = tk.StringVar(value="ID 4 chữ số: 0")
        digits_other_count_var = tk.StringVar(value="ID khác: 0")
        
        # Create labels for statistics
        ttk.Label(stats_frame, textvariable=total_count_var, font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(stats_frame, textvariable=digits_3_count_var, font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(stats_frame, textvariable=digits_4_count_var, font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(stats_frame, textvariable=digits_other_count_var, font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        
        # Function to update statistics
        def update_stats():
            all_ids = list(id_listbox.get(0, tk.END))
            total = len(all_ids)
            
            # Count by length
            digits_3 = sum(1 for id in all_ids if len(id) == 3)
            digits_4 = sum(1 for id in all_ids if len(id) == 4)
            digits_other = total - digits_3 - digits_4
            
            # Update variables
            total_count_var.set(f"Tổng số ID: {total}")
            digits_3_count_var.set(f"ID 3 chữ số: {digits_3}")
            digits_4_count_var.set(f"ID 4 chữ số: {digits_4}")
            digits_other_count_var.set(f"ID khác: {digits_other}")
        
        # Buttons for ID management
        button_frame = ttk.Frame(all_students_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # Function to add an ID
        def add_id():
            input_text = student_id_var.get().strip()
            if not input_text:
                id_status_var.set("ID không được để trống")
                return
                
            # Check if the input contains commas - multiple IDs
            if ',' in input_text:
                ids = [id.strip() for id in input_text.split(',') if id.strip()]
                
                # Validate and process each ID
                valid_ids = []
                invalid_ids = []
                duplicate_ids = []
                semantic_duplicates = []
                
                current_ids = list(id_listbox.get(0, tk.END))
                normalized_current_ids = {self.normalize_student_id(id): id for id in current_ids}
                
                for student_id in ids:
                    is_valid, error_msg = self.validate_student_id(student_id)
                    if not is_valid:
                        invalid_ids.append(student_id)
                        continue
                    
                    # Check for exact duplicates
                    if student_id in current_ids:
                        duplicate_ids.append(student_id)
                        continue
                    
                    # Check for semantic duplicates (same ID with different zeros)
                    normalized_id = self.normalize_student_id(student_id)
                    if normalized_id in normalized_current_ids and student_id != normalized_current_ids[normalized_id]:
                        semantic_duplicates.append(f"{student_id} (trùng với {normalized_current_ids[normalized_id]})")
                        continue
                    
                    valid_ids.append(student_id)
                    # Add to tracking collections
                    current_ids.append(student_id)
                    normalized_current_ids[normalized_id] = student_id
                
                # Add to listbox
                for id in valid_ids:
                    id_listbox.insert(tk.END, id)
                
                # Update statistics
                update_stats()
                
                # Show report on what happened
                msg = []
                if valid_ids:
                    msg.append(f"Đã thêm {len(valid_ids)} ID")
                if invalid_ids:
                    msg.append(f"{len(invalid_ids)} ID không hợp lệ")
                if duplicate_ids:
                    msg.append(f"{len(duplicate_ids)} ID trùng lặp")
                if semantic_duplicates:
                    msg.append(f"{len(semantic_duplicates)} ID trùng lặp (khác số 0)")
                    
                    # Show detailed message about semantic duplicates
                    if len(semantic_duplicates) <= 3:  # Only show details for a few duplicates
                        detail_msg = "Trùng lặp: " + ", ".join(semantic_duplicates)
                        messagebox.showinfo("ID trùng lặp (khác số 0)", detail_msg)
                
                id_status_var.set(" - ".join(msg))
                student_id_var.set("")  # Clear entry
                student_id_entry.focus()  # Keep focus on entry field
                return
                
            # Single ID processing
            student_id = input_text
            is_valid, error_msg = self.validate_student_id(student_id)
            if not is_valid:
                id_status_var.set(error_msg)
                return
                
            # Check if ID already exists in the list (exact match)
            current_ids = list(id_listbox.get(0, tk.END))
            if student_id in current_ids:
                id_status_var.set("ID này đã tồn tại trong danh sách!")
                return
            
            # Check for semantic duplicates (same number with different leading zeros)
            normalized_id = self.normalize_student_id(student_id)
            for existing_id in current_ids:
                if self.normalize_student_id(existing_id) == normalized_id:
                    id_status_var.set(f"ID này đã tồn tại với dạng {existing_id}!")
                    return
                
            id_listbox.insert(tk.END, student_id)
            student_id_var.set("")  # Clear entry
            id_status_var.set("Đã thêm 1 ID")  # Clear error message
            student_id_entry.focus()  # Keep focus on entry field for fast input
            
            # Update statistics
            update_stats()
            
        # Function to remove selected ID
        def remove_id():
            selected = id_listbox.curselection()
            if not selected:
                id_status_var.set("Vui lòng chọn ID để xóa")
                return
                
            id_listbox.delete(selected)
            id_status_var.set("")
            student_id_entry.focus()  # Return focus to entry field
            
            # Update statistics
            update_stats()
            
        # Add ID button
        add_btn = ttk.Button(button_frame, text="Thêm ID", command=add_id, width=15)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Remove ID button
        remove_btn = ttk.Button(button_frame, text="Xóa ID đã chọn", command=remove_id, width=15)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear all button
        def clear_all():
            id_listbox.delete(0, tk.END)
            id_status_var.set("")
            student_id_entry.focus()
            
            # Update statistics
            update_stats()
            
        clear_btn = ttk.Button(button_frame, text="Xóa tất cả", command=clear_all, width=15)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Create a hidden entry field to store the comma-separated list for compatibility
        student_ids_entry = ttk.Entry(all_students_frame)
        student_ids_entry.pack_forget()  # Hide this widget
        
        # Update the hidden entry when needed for compatibility with existing code
        def update_hidden_entry():
            all_ids = list(id_listbox.get(0, tk.END))
            student_ids_entry.delete(0, tk.END)
            student_ids_entry.insert(0, ",".join(all_ids))
            
        # Bind events to update hidden entry
        id_listbox.bind('<<ListboxSelect>>', lambda e: update_hidden_entry())
        
        # Override the process_single_class to update the hidden entry first
        original_process = self.process_single_class
        def wrapped_process(tab_id):
            update_hidden_entry()
            original_process(tab_id)
        self.process_single_class = wrapped_process
        
        # Load existing IDs if available (paste them to the listbox)
        def load_ids_from_clipboard():
            try:
                clipboard = self.root.clipboard_get()
                ids = [id.strip() for id in clipboard.split(',') if id.strip()]
                
                # Validate each ID
                valid_ids = []
                invalid_ids = []
                duplicate_ids = []
                semantic_duplicates = []
                
                current_ids = list(id_listbox.get(0, tk.END))
                normalized_current_ids = {self.normalize_student_id(id): id for id in current_ids}
                
                for student_id in ids:
                    is_valid, _ = self.validate_student_id(student_id)
                    if not is_valid:
                        invalid_ids.append(student_id)
                        continue
                    
                    # Check for exact duplicates
                    if student_id in current_ids:
                        duplicate_ids.append(student_id)
                        continue
                    
                    # Check for semantic duplicates (same ID with different zeros)
                    normalized_id = self.normalize_student_id(student_id)
                    if normalized_id in normalized_current_ids and student_id != normalized_current_ids[normalized_id]:
                        semantic_duplicates.append(f"{student_id} (trùng với {normalized_current_ids[normalized_id]})")
                        continue
                    
                    valid_ids.append(student_id)
                    # Add to tracking collections
                    current_ids.append(student_id)
                    normalized_current_ids[normalized_id] = student_id
                
                # Add to listbox
                for id in valid_ids:
                    id_listbox.insert(tk.END, id)
                
                # Update hidden entry
                update_hidden_entry()
                
                # Update statistics
                update_stats()
                
                # Show report on what happened
                msg = []
                if valid_ids:
                    msg.append(f"Đã nạp {len(valid_ids)} ID")
                if invalid_ids:
                    msg.append(f"{len(invalid_ids)} ID không hợp lệ")
                if duplicate_ids:
                    msg.append(f"{len(duplicate_ids)} ID trùng lặp")
                if semantic_duplicates:
                    msg.append(f"{len(semantic_duplicates)} ID trùng lặp (khác số 0)")
                    
                    # Show detailed message about semantic duplicates
                    if len(semantic_duplicates) <= 3:  # Only show details for a few duplicates
                        detail_msg = "Trùng lặp: " + ", ".join(semantic_duplicates)
                        messagebox.showinfo("ID trùng lặp (khác số 0)", detail_msg)
                
                id_status_var.set(" - ".join(msg))
                student_id_entry.focus()
            except Exception as e:
                id_status_var.set("Không thể nạp từ clipboard")
                logging.error(f"Error loading from clipboard: {str(e)}")
        
        paste_btn = ttk.Button(button_frame, text="Nạp từ clipboard", command=load_ids_from_clipboard, width=15)
        paste_btn.pack(side=tk.LEFT, padx=5)
        
        # Status label
        status_label = ttk.Label(class_frame, text="")
        status_label.pack(pady=5)
        
        return {
            'id': tab_id,
            'frame': class_frame,
            'lesson_type': lesson_type,
            'online_ids_entry': online_ids_entry,
            'student_ids_entry': student_ids_entry,
            'id_listbox': id_listbox,
            'status_label': status_label
        }

    def add_class_tab(self):
        tab_id = len(self.class_tabs) + 1
        tab_info = self.create_class_tab(tab_id)
        
        # Insert before report tab with default name
        self.notebook.insert(self.notebook.index(self.report_frame), tab_info['frame'], text=f"Lớp {tab_id}")
        
        # Store the tab name in the tab_info dictionary for later reference
        tab_info['name'] = f"Lớp {tab_id}"
        
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

        # Find the tab to delete
        deleted_index = None
        for i, tab in enumerate(self.class_tabs):
            if tab['id'] == tab_id:
                self.notebook.forget(tab['frame'])
                self.class_tabs.pop(i)
                deleted_index = i
                break
        
        if deleted_index is None:
            return
            
        # Renumber remaining tabs
        for i, tab in enumerate(self.class_tabs, 1):
            tab['id'] = i
            tab_index = self.notebook.index(tab['frame'])
            
            # Preserve custom names if they exist, otherwise use default naming
            if 'name' not in tab or tab['name'].startswith('Lớp '):
                # This is a default name, update it
                tab['name'] = f"Lớp {i}"
                self.notebook.tab(tab_index, text=tab['name'])
            # If it's a custom name, we keep it as is

    def setup_report_tab(self):
        report_container = ttk.Frame(self.report_frame)
        report_container.pack(fill=tk.BOTH, expand=True)
        
        self.report_text = tk.Text(report_container, height=15, width=60)
        scrollbar = ttk.Scrollbar(report_container, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=scrollbar.set)
        
        self.report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.report_text.config(state=tk.DISABLED)

    def normalize_student_id(self, student_id):
        """Normalize student ID by stripping leading zeros"""
        return student_id.lstrip('0') or '0'  # Return '0' if ID is all zeros
    
    def validate_student_id(self, student_id):
        student_id = student_id.strip()
        if not student_id.isdigit():
            return False, "Mã học sinh chỉ được chứa số"
        if len(student_id) < 2 or len(student_id) > 6:
            return False, "Độ dài mã học sinh không hợp lệ"
        return True, ""

    def check_for_duplicates(self, id_list):
        seen = set()
        duplicates = []
        normalized_ids = {}
        
        # First pass: build a mapping of normalized IDs to original IDs
        for id in id_list:
            normalized = self.normalize_student_id(id)
            if normalized in normalized_ids:
                # Already seen this normalized ID
                normalized_ids[normalized].append(id)
            else:
                normalized_ids[normalized] = [id]
        
        # Second pass: find duplicates
        for normalized, ids in normalized_ids.items():
            if len(ids) > 1:
                # If we have multiple IDs with the same normalized form
                duplicates.extend(ids)
        
        return duplicates

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
        
        # Check for duplicates (both exact and semantic)
        exact_duplicates = self.check_for_duplicates(student_list)
        if exact_duplicates:
            return False, f"Có ID trùng lặp: {', '.join(exact_duplicates)}"
            
        # Check for semantic duplicates
        normalized_ids = {}
        semantic_duplicates = []
        
        for id in student_list:
            normalized = self.normalize_student_id(id)
            if normalized in normalized_ids:
                # This is a semantic duplicate
                existing_id = normalized_ids[normalized]
                semantic_duplicates.append(f"{id} (trùng với {existing_id})")
            else:
                normalized_ids[normalized] = id
                
        if semantic_duplicates:
            return False, f"Có ID trùng lặp (khác số 0): {', '.join(semantic_duplicates)}"
            
        return True, student_list

    def process_all_classes(self):
        if not self.class_tabs:
            messagebox.showwarning("Cảnh báo", "Không có lớp nào để điểm danh!")
            return
            
        # First update all hidden entries from listboxes
        for tab in self.class_tabs:
            all_ids = list(tab['id_listbox'].get(0, tk.END))
            tab['student_ids_entry'].delete(0, tk.END)
            tab['student_ids_entry'].insert(0, ",".join(all_ids))
            
        all_failed_students = []
        
        # Validate all inputs first
        class_configs = []
        for tab in self.class_tabs:
            student_ids = tab['student_ids_entry'].get().strip()
            
            # Check if the student_ids is empty but the listbox has values
            if not student_ids and tab['id_listbox'].size() > 0:
                all_ids = list(tab['id_listbox'].get(0, tk.END))
                student_ids = ",".join(all_ids)
                tab['student_ids_entry'].delete(0, tk.END)
                tab['student_ids_entry'].insert(0, student_ids)
                
            # Double check if it's still empty
            if not student_ids:
                messagebox.showerror("Lỗi", f"Lỗi ở lớp {tab['id']}: Vui lòng nhập ít nhất một mã học sinh")
                return
                
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
        
        # Create a progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Tiến trình điểm danh")
        progress_window.geometry("400x300")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        
        # Overall progress frame
        overall_frame = ttk.LabelFrame(progress_window, text="Tiến trình tổng thể", padding=10)
        overall_frame.pack(fill=tk.X, padx=10, pady=10)
        
        overall_label = ttk.Label(overall_frame, text="0/0 lớp hoàn thành")
        overall_label.pack(pady=5)
        
        overall_progress = ttk.Progressbar(overall_frame, orient=tk.HORIZONTAL, length=350, mode='determinate')
        overall_progress.pack(fill=tk.X, pady=5)
        
        # Current class progress frame
        class_frame = ttk.LabelFrame(progress_window, text="Lớp đang xử lý", padding=10)
        class_frame.pack(fill=tk.X, padx=10, pady=10)
        
        class_label = ttk.Label(class_frame, text="Đang khởi tạo...")
        class_label.pack(pady=5)
        
        student_label = ttk.Label(class_frame, text="0/0 học sinh")
        student_label.pack(pady=5)
        
        class_progress = ttk.Progressbar(class_frame, orient=tk.HORIZONTAL, length=350, mode='determinate')
        class_progress.pack(fill=tk.X, pady=5)
        
        # Processing info
        log_frame = ttk.LabelFrame(progress_window, text="Thông tin xử lý", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        log_text = tk.Text(log_frame, height=5, width=40)
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_text.configure(yscrollcommand=log_scrollbar.set)
        
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Function to add log message with timestamp
        def add_log(message):
            timestamp = time.strftime("%H:%M:%S", time.localtime())
            log_text.config(state=tk.NORMAL)
            log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            log_text.see(tk.END)
            log_text.config(state=tk.DISABLED)
            
        # Process classes one by one to improve UI responsiveness
        total_classes = len(class_configs)
        overall_progress['maximum'] = total_classes
        
        def process_class_by_index(idx):
            if idx >= len(class_configs):
                # All classes have been processed
                add_log("Hoàn thành tất cả lớp!")
                overall_label.config(text=f"{total_classes}/{total_classes} lớp hoàn thành")
                
                # Update report
                self.update_report(all_failed_students)
                if all_failed_students:
                    self.notebook.select(self.report_frame)
                    
                # Close progress window after 2 seconds
                progress_window.after(2000, progress_window.destroy)
                return
                
            config = class_configs[idx]
            add_log(f"Bắt đầu xử lý Lớp {config['tab_id']}")
            class_label.config(text=f"Lớp {config['tab_id']}")
            
            worker = ClassAttendanceWorker(
                config['student_ids'],
                config['online_students'],
                config['lesson_type']
            )
            
            # Set existing browser if available
            if config['browser']:
                worker.driver = config['browser']
                worker.wait = WebDriverWait(config['browser'], 20)
            
            total_students = len(config['student_ids'])
            class_progress['maximum'] = total_students
            
            # Create generator for class processing
            process_generator = worker.process_class()
            
            def process_next_student():
                try:
                    # Get next progress update
                    current, total = next(process_generator)
                    
                    # Update progress
                    percent = int((current / total) * 100)
                    student_label.config(text=f"{current}/{total} học sinh ({percent}%)")
                    class_progress['value'] = current
                    
                    # Add log for every 5th student or last student
                    if current % 5 == 0 or current == total:
                        add_log(f"Đã xử lý {current}/{total} học sinh")
                    
                    # Process next student
                    progress_window.after(10, process_next_student)
                    
                except StopIteration as result:
                    # Processing is complete for current class
                    failed_students = worker.failed_students
                    
                    if failed_students:
                        # Add class information to each failed student
                        class_failed_students = [(student_id, error, config['tab_id']) for student_id, error in failed_students]
                        all_failed_students.extend(class_failed_students)
                        config['status_label'].config(text=f"Hoàn thành! Có {len(failed_students)} học sinh bị lỗi.")
                        add_log(f"Lớp {config['tab_id']} hoàn thành với {len(failed_students)} lỗi")
                    else:
                        config['status_label'].config(text="Hoàn thành điểm danh!")
                        add_log(f"Lớp {config['tab_id']} hoàn thành không có lỗi")
                    
                    # Update overall progress
                    overall_progress['value'] = idx + 1
                    overall_label.config(text=f"{idx + 1}/{total_classes} lớp hoàn thành")
                    
                    # Process next class
                    progress_window.after(500, lambda: process_class_by_index(idx + 1))
                    
                except Exception as e:
                    logging.error(f"Error in worker thread: {str(e)}")
                    config['status_label'].config(text="Lỗi rùi huhu")
                    all_failed_students.append((None, str(e), config['tab_id']))
                    add_log(f"Lỗi xử lý Lớp {config['tab_id']}: {str(e)}")
                    
                    # Process next class
                    progress_window.after(500, lambda: process_class_by_index(idx + 1))
            
            # Start processing the current class
            progress_window.after(10, process_next_student)
        
        # Start processing with the first class
        progress_window.after(100, lambda: process_class_by_index(0))

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

        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title(f"Tiến trình điểm danh - Lớp {tab_id}")
        progress_window.geometry("400x250")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(progress_window, text=f"Lớp {tab_id}", padding=10)
        progress_frame.pack(fill=tk.X, padx=10, pady=10)
        
        student_label = ttk.Label(progress_frame, text="0/0 học sinh")
        student_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=350, mode='determinate')
        progress_bar.pack(fill=tk.X, pady=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(progress_window, text="Thông tin xử lý", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        log_text = tk.Text(log_frame, height=5, width=40)
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_text.configure(yscrollcommand=log_scrollbar.set)
        
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Function to add log message with timestamp
        def add_log(message):
            timestamp = time.strftime("%H:%M:%S", time.localtime())
            log_text.config(state=tk.NORMAL)
            log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            log_text.see(tk.END)
            log_text.config(state=tk.DISABLED)
            
        # Create generator for processing
        add_log(f"Bắt đầu xử lý Lớp {tab_id}")
        process_generator = worker.process_class()
        
        total_students = len(result)
        progress_bar['maximum'] = total_students
        
        def process_next_student():
            try:
                # Get next progress update
                current, total = next(process_generator)
                
                # Update progress
                percent = int((current / total) * 100)
                student_label.config(text=f"{current}/{total} học sinh ({percent}%)")
                progress_bar['value'] = current
                
                # Add log for every 5th student or last student
                if current % 5 == 0 or current == total:
                    add_log(f"Đã xử lý {current}/{total} học sinh")
                
                # Process next student
                progress_window.after(10, process_next_student)
                
            except StopIteration:
                # Processing is complete
                failed_students = worker.failed_students
                
                if failed_students:
                    # Add class information to each failed student
                    class_failed_students = [(student_id, error, tab_id) for student_id, error in failed_students]
                    tab['status_label'].config(text=f"Hoàn thành! Có {len(failed_students)} học sinh bị lỗi.")
                    add_log(f"Hoàn thành với {len(failed_students)} lỗi")
                    self.update_report(class_failed_students)
                    self.notebook.select(self.report_frame)
                else:
                    tab['status_label'].config(text="Hoàn thành điểm danh!")
                    add_log("Hoàn thành không có lỗi")
                
                # Close progress window after 2 seconds
                progress_window.after(2000, progress_window.destroy)
                
            except Exception as e:
                logging.error(f"Error processing class: {str(e)}")
                tab['status_label'].config(text="Lỗi rùi huhu")
                self.update_report([(None, str(e), tab_id)])
                self.notebook.select(self.report_frame)
                add_log(f"Lỗi: {str(e)}")
                progress_window.after(2000, progress_window.destroy)
        
        # Start processing
        progress_window.after(100, process_next_student)

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
                logging.info(f"Browser opened for class tab {tab_id}")
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

    def rename_class_tab(self, tab_id):
        # Find the tab with the given ID
        tab = next((tab for tab in self.class_tabs if tab['id'] == tab_id), None)
        if not tab:
            return
            
        # Create a simple dialog to get the new name
        dialog = tk.Toplevel(self.root)
        dialog.title("Đổi tên lớp")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        
        # Make the dialog modal
        dialog.grab_set()
        
        # Get current tab label
        current_tab_index = self.notebook.index(tab['frame'])
        current_tab_text = self.notebook.tab(current_tab_index, "text")
        
        # Frame for the dialog content
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Label and entry for the new name
        ttk.Label(frame, text="Nhập tên mới cho lớp:", font=('Arial', 10)).pack(pady=(0, 10))
        
        new_name_var = tk.StringVar(value=current_tab_text)
        name_entry = ttk.Entry(frame, textvariable=new_name_var, width=30, font=('Arial', 10))
        name_entry.pack(pady=5)
        name_entry.select_range(0, tk.END)  # Select all text for easy editing
        name_entry.focus_set()  # Set focus to the entry
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Cancel button
        cancel_btn = ttk.Button(button_frame, text="Hủy", command=dialog.destroy, width=10)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Function to save the new name
        def save_name():
            new_name = new_name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Cảnh báo", "Tên lớp không được để trống!")
                return
                
            # Update the tab text
            tab_index = self.notebook.index(tab['frame'])
            self.notebook.tab(tab_index, text=new_name)
            
            # Store the new name in the tab info
            tab['name'] = new_name
            
            dialog.destroy()
        
        # Save button
        save_btn = ttk.Button(button_frame, text="Lưu", command=save_name, width=10)
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # Bind Enter key to save
        dialog.bind("<Return>", lambda event: save_name())
        
        # Center the dialog on the main window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Wait for the dialog to close
        self.root.wait_window(dialog)

if __name__ == "__main__":
    app = StudentAttendanceChecker()
    app.run()