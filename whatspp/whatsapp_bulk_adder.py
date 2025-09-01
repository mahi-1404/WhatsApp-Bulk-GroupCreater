import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

class WhatsAppAdderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Group Bulk Adder")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        self.file_path = ""
        self.contacts = []

        # ===== UI Design =====
        title = tk.Label(root, text="WhatsApp Group Bulk Adder", font=("Helvetica", 16, "bold"))
        title.pack(pady=10)

        file_frame = tk.Frame(root)
        file_frame.pack(pady=10)

        self.file_label = tk.Label(file_frame, text="No file selected", width=40, anchor="w")
        self.file_label.grid(row=0, column=0, padx=5)

        browse_button = tk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_button.grid(row=0, column=1)

        group_frame = tk.Frame(root)
        group_frame.pack(pady=10)

        tk.Label(group_frame, text="Group Name:", font=("Arial", 12)).grid(row=0, column=0, padx=5, sticky="e")
        self.group_name_entry = tk.Entry(group_frame, font=("Arial", 12), width=30)
        self.group_name_entry.grid(row=0, column=1, padx=5)

        start_button = tk.Button(root, text="Start Adding Members", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", command=self.start_automation)
        start_button.pack(pady=20)

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if self.file_path:
            filename = os.path.basename(self.file_path)
            self.file_label.config(text=filename)

    def read_contacts(self):
        if self.file_path.endswith(".csv"):
            df = pd.read_csv(self.file_path)
        else:
            df = pd.read_excel(self.file_path)
        self.contacts = df[['Name', 'Phone']].dropna().to_dict('records')

    def setup_driver(self):
        options = Options()
        options.add_argument("user-data-dir=./user_profile")  # Keep login session
        service = Service('chromedriver')  # Ensure chromedriver is in your PATH
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://web.whatsapp.com")
        input("Scan QR code in the browser and press Enter here to continue...")
        return driver

    def add_members_to_group(self, driver, group_name):
        try:
            search_box = driver.find_element(By.XPATH, '//div[@title="Search input textbox"]')
            search_box.click()
            search_box.send_keys(group_name)
            time.sleep(3)

            driver.find_element(By.XPATH, f'//span[@title="{group_name}"]').click()
            time.sleep(3)

            driver.find_element(By.XPATH, '//header').click()
            time.sleep(2)

            driver.find_element(By.XPATH, '//div[@title="Add participant"]').click()
            time.sleep(2)

            for contact in self.contacts:
                try:
                    number = contact['Phone']
                    search = driver.find_element(By.XPATH, '//div[@title="Type contact name"]')
                    search.clear()
                    search.send_keys(str(number))
                    time.sleep(2)
                    driver.find_element(By.XPATH, f'//span[@title="{number}"]').click()
                    print(f"Added: {number}")
                except Exception as e:
                    print(f"Failed to add {contact['Name']} - {e}")

            driver.find_element(By.XPATH, '//span[@data-icon="checkmark"]').click()
            print("Members added successfully.")
        except Exception as e:
            print("Error during group addition:", e)

        time.sleep(5)

    def start_automation(self):
        group_name = self.group_name_entry.get().strip()
        if not self.file_path or not group_name:
            messagebox.showerror("Missing Info", "Please select a file and enter a group name.")
            return

        try:
            self.read_contacts()
            driver = self.setup_driver()
            self.add_members_to_group(driver, group_name)
            driver.quit()
            messagebox.showinfo("Success", "Members added successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{e}")

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAdderApp(root)
    root.mainloop()
