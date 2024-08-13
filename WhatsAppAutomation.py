import tkinter as tk
from tkinter import messagebox, Frame, Label, Entry, StringVar
from tkinter import filedialog
from PIL import Image, ImageTk
import os
import json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
import time
from selenium.common.exceptions import NoAlertPresentException, TimeoutException, NoSuchElementException, WebDriverException
import logging
import pandas as pd
import sys

class WhatsAppAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Automation")
        self.root.geometry("800x600")
        self.logged_in_user = None
        self.users_file = self.resource_path('users.json')
        self.load_users()

        self.show_login_page()

    def load_users(self):
        if not os.path.exists(self.users_file):
            self.users = {"admin": {"password": "admin", "is_admin": True}}
            self.save_users()
        else:
            with open(self.users_file, 'r') as file:
                self.users = json.load(file)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def save_users(self):
        with open(self.users_file, 'w') as file:
            json.dump(self.users, file, indent=4)

    def show_login_page(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        file_path = self.resource_path("login_image.jpg")
        self.img_path = Image.open(file_path)
        self.image_resize = self.img_path.resize((350, 400))
        self.img = ImageTk.PhotoImage(self.image_resize)
        Label(self.root, image=self.img).place(x=30, y=40)

        self.frame1 = Frame(self.root, width=350, height=400, bg="white")
        self.frame1.place(x=415, y=40)

        Label(self.root, text="Sign-In", bg="white", fg="#05739c", font=("Microsoft JhengHei", 23, "bold")).place(x=525, y=70)

        self.user_name = Entry(self.root, width=30, font=("Microsoft JhengHei", 10), border=0, fg="#41484a")
        self.user_name.place(x=470, y=180)
        self.user_name.insert(0, "username")
        self.user_name.bind('<FocusIn>', self.on_enter_username)
        self.user_name.bind('<FocusOut>', self.on_leave_username)
        Frame(self.root, height=2, width=290, bg="black").place(x=470, y=200)

        self.password = Entry(self.root, width=30, font=("Microsoft JhengHei", 10), border=0, fg="#41484a", show="*")
        self.password.place(x=470, y=230)
        self.password.insert(0, "password")
        self.password.bind('<FocusIn>', self.on_enter_password)
        self.password.bind('<FocusOut>', self.on_leave_password)
        Frame(self.root, height=2, width=290, bg="black").place(x=470, y=250)

        tk.Button(self.root, text='Sign-in', fg="white", bg="#0eb1ed", width=35, border=0, height=2, font=("Microsoft JhengHei", 8, "bold"), command=self.login).place(x=485, y=290)

    def on_enter_username(self, event):
        if self.user_name.get() == 'username':
            self.user_name.delete(0, 'end')

    def on_leave_username(self, event):
        if self.user_name.get() == '':
            self.user_name.insert(0, 'username')

    def on_enter_password(self, event):
        if self.password.get() == 'password':
            self.password.delete(0, 'end')

    def on_leave_password(self, event):
        if self.password.get() == '':
            self.password.insert(0, 'password')

    def login(self):
        username = self.user_name.get()
        password = self.password.get()

        if username in self.users and self.users[username]["password"] == password:
            self.logged_in_user = username
            if self.users[username].get("is_admin"):
                self.show_options_page()
            else:
                self.show_main_page()
        else:
            messagebox.showerror("Error", "Invalid username or password.")

    def show_signup_page(self):
        for widget in self.root.winfo_children():
          widget.destroy()
        file_path = os.path.join(os.path.dirname(__file__), "login_image.jpg")
        self.img_path = Image.open(file_path)
        self.image_resize = self.img_path.resize((350, 400))
        self.img = ImageTk.PhotoImage(self.image_resize)
        Label(self.root, image=self.img).place(x=30, y=40)
        Label(self.root, text="Sign-Up", font=("Microsoft JhengHei", 23, "bold")).place(x=525, y=70)
    
        self.signup_username = Entry(self.root, width=30, font=("Microsoft JhengHei", 10), border=0, fg="#41484a")
        self.signup_username.place(x=470, y=180)
        self.signup_username.insert(0, "username")
        self.signup_username.bind('<FocusIn>', self.on_enter_signup_username)
        self.signup_username.bind('<FocusOut>', self.on_leave_signup_username)
        Frame(self.root, height=2, width=290, bg="black").place(x=470, y=200)

        self.signup_password = Entry(self.root, width=30, font=("Microsoft JhengHei", 10), border=0, fg="#41484a")
        self.signup_password.place(x=470, y=230)
        self.signup_password.insert(0, "password")
        self.signup_password.bind('<FocusIn>', self.on_enter_signup_password)
        self.signup_password.bind('<FocusOut>', self.on_leave_signup_password)
        Frame(self.root, height=2, width=290, bg="black").place(x=470, y=250)

        tk.Button(self.root, text="Create User", command=self.create_user).place(x=570, y=300)
        tk.Button(self.root, text="Back", command=self.show_login_page).place(x=650, y=300)

    def on_enter_signup_username(self, event):
        if self.signup_username.get() == 'username':
            self.signup_username.delete(0, 'end')

    def on_leave_signup_username(self, event):
        if self.signup_username.get() == '':
           self.signup_username.insert(0, 'username')

    def on_enter_signup_password(self, event):
       if self.signup_password.get() == 'password':
           self.signup_password.delete(0, 'end')
           self.signup_password.config(show="*")

    def on_leave_signup_password(self, event):
         if self.signup_password.get() == '':
           self.signup_password.insert(0, 'password')
           self.signup_password.config(show="")

    def create_user(self):
        username = self.signup_username.get()
        password = self.signup_password.get()

        if username in self.users:
            messagebox.showerror("Error", "Username already exists.")
        else:
            self.users[username] = {"password": password, "is_admin": False}
            self.save_users()
            messagebox.showinfo("Success", "User created successfully.")
            self.show_login_page()

    def show_options_page(self):
        for widget in self.root.winfo_children():
            widget.destroy()
       
        self.root.title("Options")
        self.root.config(bg="white")
       
        tk.Button(self.root, text="Create User",fg="white",bg="#0eb1ed",border=0,width=30,height=2,font=("Microsoft JhengHei",9,"bold"), command=self.show_signup_page).pack(pady=20)
        tk.Button(self.root, text="WhatsApp Automation",fg="white",bg="#0eb1ed",border=0,width=30,height=2,font=("Microsoft JhengHei",9,"bold"), command=self.show_main_page).pack(pady=20)

    def show_main_page(self):
        for widget in self.root.winfo_children():
           widget.destroy()
 
        self.root.geometry("1090x800")
    
    # Add title
        title_label = tk.Label(self.root, text="KRISHTEC WHATSAPP AUTOMATION", fg="blue",bg="white",font=("Microsoft JhengHei", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=20)
    
        left_frame = tk.Frame(self.root)
        left_frame.grid(row=1, column=0, padx=20, pady=20)

        right_frame = tk.Frame(self.root)
        right_frame.grid(row=1, column=1, padx=20, pady=20)

    # "Send Wishes" UI Elements
        tk.Label(left_frame, text="Send Wishes", font=("Microsoft JhengHei", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)
        tk.Label(left_frame, text="Select Excel File:").grid(row=1, column=0, padx=10, pady=10)
        self.excel_entry = tk.Entry(left_frame, width=50)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(left_frame, text="Browse", command=self.browse_excel).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(left_frame, text="Select Occasion:").grid(row=2, column=0, padx=10, pady=10)
        self.occasion_type = tk.StringVar()
        tk.OptionMenu(left_frame, self.occasion_type, "Pongal", "Diwali", "Christmas", "Engineering Day", "New Year").grid(row=2, column=1, padx=10, pady=10)

        tk.Label(left_frame, text="Select File Type:").grid(row=3, column=0, padx=10, pady=10)
        self.file_type = tk.StringVar()
        tk.OptionMenu(left_frame, self.file_type, "Image", "Video", "PDF").grid(row=3, column=1, padx=10, pady=10)

        tk.Button(left_frame, text="Upload File", command=self.upload_file).grid(row=4, column=1, padx=10, pady=10)
        tk.Button(left_frame, text="Send Wishes", command=self.send_wishes).grid(row=5, column=1, padx=10, pady=10)

    # "Send Message" UI Elements
        tk.Label(right_frame, text="Send Message", font=("Microsoft JhengHei", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)
        tk.Label(right_frame, text="Select Excel File:").grid(row=1, column=0, padx=10, pady=10)
        self.excel_entry_msg = tk.Entry(right_frame, width=50)
        self.excel_entry_msg.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(right_frame, text="Browse", command=self.browse_excel_msg).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(right_frame, text="Select Message Type:").grid(row=2, column=0, padx=10, pady=10)
        self.message_type = tk.StringVar()
        tk.OptionMenu(right_frame, self.message_type, "Text", "Image", "Video", "PDF").grid(row=2, column=1, padx=10, pady=10)

        tk.Button(right_frame, text="Upload File", command=self.upload_file_msg).grid(row=3, column=1, padx=10, pady=10)
        tk.Button(right_frame, text="Send Message", command=self.send_message).grid(row=4, column=1, padx=10, pady=10)

    def browse_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.excel_entry.delete(0, tk.END)
        self.excel_entry.insert(0, self.excel_path)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename()
        
    def send_wishes(self):
        if not self.excel_entry.get():
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        if not self.occasion_type.get():
            messagebox.showerror("Error", "Please select an occasion.")
            return

        if not self.file_type.get():
            messagebox.showerror("Error", "Please select a file type.")
            return

        # Read contacts from Excel file
        try:
            contacts_df = pd.read_excel(self.excel_entry.get())
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

        # Initialize WebDriver with ChromeDriverManager
        chrome_options = Options()
        chrome_options.add_argument("--incognito")
        chrome_options.add_argument("--disable-extensions")
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

        driver.get('https://web.whatsapp.com')

        messagebox.showinfo("Info", "Scan the QR code and press OK to continue.")

        # Wait for user to scan QR code
        while 'WhatsApp' not in driver.title:
            time.sleep(2)

        # Define wishes based on the occasion
        wishes = {
            "Pongal": "Happy Pongal, {name}!",
            "Diwali": "Happy Diwali, {name}!",
            "Christmas": "Merry Christmas, {name}!",
            "Engineering Day": "Happy Engineering Day, {name}!",
            "New Year": "Happy New Year, {name}!"
        }

        for index, row in contacts_df.iterrows():
            name = row['Names']
            number = row['Numbers']

            driver.get(f'https://web.whatsapp.com/send?phone={number}')

            try:
                alert = driver.switch_to.alert
                alert.dismiss()  
            except NoAlertPresentException:
                pass  

            try:
            
                if self.file_path:
                    attachment_box = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@title='Attach']"))
                    )
                    attachment_box.click()

                    file_input = None
                    if self.file_type.get() == "Image":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")
                    elif self.file_type.get() == "Video":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='video/mp4,video/3gpp,video/quicktime']")
                    elif self.file_type.get() == "PDF":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='*']")

                    file_input.send_keys(self.file_path)
                    time.sleep(2)
                    send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
                    send_button.click()
                    time.sleep(2)

                # Wait for message box to be available
                message_box = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Type a message"]'))
                )
                message_text = wishes[self.occasion_type.get()].format(name=name)
                message_box.send_keys(message_text)
                message_box.send_keys(Keys.ENTER)

            except (TimeoutException, NoSuchElementException, WebDriverException) as e:
                error_message = f"Failed to send message to {name} ({number}): {e}"
                logging.error(error_message)
                print(error_message)
                continue

            time.sleep(5)

        driver.quit()
        messagebox.showinfo("Info", "Wishes sent successfully!")


    def browse_excel_msg(self):
        self.excel_path_msg = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.excel_entry_msg.delete(0, tk.END)
        self.excel_entry_msg.insert(0, self.excel_path_msg)

    def upload_file_msg(self):
        self.file_path_msg = filedialog.askopenfilename()

    def send_message(self):
        if not self.excel_entry_msg.get():
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        if self.message_type.get() != "Text" and not self.file_path_msg:
            messagebox.showerror("Error", "Please upload a file for non-text messages.")
            return

        try:
            contacts_df = pd.read_excel(self.excel_entry_msg.get())
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

        chrome_options = Options()
        chrome_options.add_argument("--incognito")
        chrome_options.add_argument("--disable-extensions")
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

        driver.get('https://web.whatsapp.com')

        messagebox.showinfo("Info", "Scan the QR code and press OK to continue.")

        while 'WhatsApp' not in driver.title:
            time.sleep(2)

        for index, row in contacts_df.iterrows():
            name = row['Names']
            number = row['Numbers']

            driver.get(f'https://web.whatsapp.com/send?phone={number}')

            try:
                alert = driver.switch_to.alert
                alert.dismiss()
            except NoAlertPresentException:
                pass

            if self.message_type.get() == "Text":
                try:
                    message_box = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='6']"))
                    )
                    message_box.send_keys("Hello, this is an automated message.")
                    message_box.send_keys(Keys.ENTER)
                except Exception as e:
                    print(f"Failed to send message to {name} ({number}): {e}")
            else:
                try:
                    attachment_box = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@title='Attach']"))
                    )
                    attachment_box.click()

                    file_input = None
                    if self.message_type.get() == "Image":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")
                    elif self.message_type.get() == "Video":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='video/mp4,video/3gpp,video/quicktime']")
                    elif self.message_type.get() == "PDF":
                        file_input = driver.find_element(By.XPATH, "//input[@accept='*']")

                    file_input.send_keys(self.file_path_msg)
                    time.sleep(2)
                    send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
                    send_button.click()
                    time.sleep(2)
                except Exception as e:
                    print(f"Failed to send message to {name} ({number}): {e}")

            time.sleep(5)

        driver.quit()
        messagebox.showinfo("Info", "Messages sent successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAutomationApp(root)
    root.mainloop()
