import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os

class WhatsAppGroupCreator:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Group Creator")
        self.root.geometry("700x600")
        self.file_path = ""
        self.default_members = []

        # File Selection
        tk.Label(root, text="Select Excel/CSV File:", font=("Arial", 12)).pack(pady=5)
        file_frame = tk.Frame(root)
        file_frame.pack(pady=5)
        self.file_label = tk.Label(file_frame, text="No file selected", width=50, anchor="w")
        self.file_label.grid(row=0, column=0, padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=1)

        # Default Members Section
        tk.Label(root, text="Add Default Group Members (Name and Number):", font=("Arial", 12, "bold")).pack(pady=10)
        self.members_frame = tk.Frame(root)
        self.members_frame.pack()

        self.add_member_entry()  # Initial input row

        tk.Button(root, text="+ Add Member", command=self.add_member_entry, bg="#ccc").pack(pady=10)

        # Start button
        tk.Button(root, text="Start Group Creation", font=("Arial", 13), bg="#28a745", fg="white",
                  command=self.start_automation).pack(pady=20)

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if self.file_path:
            self.file_label.config(text=os.path.basename(self.file_path))

    def add_member_entry(self):
        frame = tk.Frame(self.members_frame)
        frame.pack(pady=2)

        name_entry = tk.Entry(frame, width=30)
        name_entry.grid(row=0, column=0, padx=5)
        name_entry.insert(0, "Name")

        number_entry = tk.Entry(frame, width=30)
        number_entry.grid(row=0, column=1, padx=5)
        number_entry.insert(0, "Phone Number")

        self.default_members.append((name_entry, number_entry))

    def get_default_members(self):
        members = []
        for name_entry, number_entry in self.default_members:
            name = name_entry.get().strip()
            number = number_entry.get().strip()
            if number:  # Name optional
                members.append(number)
        return members

    def read_contacts(self):
        if self.file_path.endswith(".csv"):
            df = pd.read_csv(self.file_path)
        else:
            df = pd.read_excel(self.file_path)

        required_columns = ['Reg No', 'Model', 'Ins Co', 'Contact No']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")

        return df[required_columns].dropna().to_dict('records')

    def setup_driver(self):
        options = Options()
        options.add_argument("user-data-dir=./user_profile")
        service = Service('chromedriver')  # Ensure chromedriver is in PATH
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://web.whatsapp.com")
        input("Scan the QR code in browser and press Enter here to continue...")
        return driver

    def create_group(self, driver, group_name, members):
        try:
            driver.get("https://web.whatsapp.com/send?phone=&text=hi")
            time.sleep(5)

            # Click Menu > New Group
            driver.find_element(By.XPATH, '//div[@title="Menu"]').click()
            time.sleep(1)
            driver.find_element(By.XPATH, '//div[text()="New group"]').click()
            time.sleep(2)

            # Add members
            for member in members:
                search = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                search.clear()
                search.send_keys(member)
                time.sleep(2)
                try:
                    driver.find_element(By.XPATH, f'//span[@title="{member}"]').click()
                except:
                    print(f"[!] Could not add member: {member}")
                time.sleep(1)

            driver.find_element(By.XPATH, '//span[@data-icon="arrow-forward"]').click()
            time.sleep(2)

            # Group name
            name_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="6"]')
            name_box.send_keys(group_name)
            time.sleep(1)
            driver.find_element(By.XPATH, '//span[@data-icon="checkmark"]').click()
            time.sleep(3)

            print(f"[+] Group created: {group_name}")
        except Exception as e:
            print(f"[!] Failed to create group {group_name}: {e}")

    def start_automation(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file.")
            return

        try:
            contacts = self.read_contacts()
            default_members = self.get_default_members()
            driver = self.setup_driver()

            for contact in contacts:
                reg = str(contact['Reg No'])
                model = str(contact['Model'])
                ins = str(contact['Ins Co'])
                phone = str(contact['Contact No'])

                last4 = reg[-4:]
                group_name = f"{last4} - {model} - {ins}"
                members = default_members + [phone]

                self.create_group(driver, group_name, members)

            driver.quit()
            messagebox.showinfo("Done", "All groups created.")
        except Exception as e:
            messagebox.showerror("Failed", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppGroupCreator(root)
    root.mainloop()
