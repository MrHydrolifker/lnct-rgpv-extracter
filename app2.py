import threading
import time
import tkinter as tk
from tkinter import scrolledtext, simpledialog
from PIL import Image
import pytesseract
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# ---------------------- SETUP ----------------------
driver = webdriver.Chrome()
driver.set_window_size(1200, 1000)
wait = WebDriverWait(driver, 20)

# ---------------------- FUNCTION TO PROCESS ROLL NUMBER ----------------------
def process_roll_number(roll_number, semester, text_area):
    try:
        print(f"\n🌐 Processing Roll No: {roll_number}  |  Semester: {semester}")
        driver.get("https://result.rgpv.ac.in/result/programselect.aspx?id=$%25")

        # Click the B.Tech label
        btech_label = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//label[@for='radlstProgram_1']"))
        )
        btech_label.click()
        print("✅ Selected B.Tech Program")

        # Wait for semester dropdown and select it before CAPTCHA (important)
        semester_dropdown = wait.until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_drpSemester"))
        )
        Select(semester_dropdown).select_by_visible_text(semester)
        print(f"📘 Selected Semester: {semester}")
        time.sleep(1)

        # Wait for CAPTCHA image and capture it
        captcha_img = wait.until(
            EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'CaptchaImage.axd')]"))
        )
        time.sleep(4)  # Give time for CAPTCHA to render fully
        captcha_img.screenshot("captcha.png")

        # OCR the CAPTCHA
        img = Image.open("captcha.png")
        captcha_text = pytesseract.image_to_string(img, config="--psm 7").strip()
        captcha_text = ''.join(c for c in captcha_text if c.isalnum())
        print("🔍 Extracted CAPTCHA:", captcha_text)

        # Fill roll number
        roll_box = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
        roll_box.clear()
        roll_box.send_keys(roll_number)
        print(f"✍️ Entered Roll Number: {roll_number}")
        time.sleep(1)

        # Fill CAPTCHA
        captcha_box = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
        captcha_box.clear()
        captcha_box.send_keys(captcha_text)
        print(f"✍️ Entered CAPTCHA Text: {captcha_text}")
        time.sleep(1.5)

        # Submit form
        submit_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='View Result']"))
        )
        submit_btn.click()
        print("🚀 Form Submitted")

        # Wait for page to load
        time.sleep(5)

        # Check for CAPTCHA alert
        try:
            alert = driver.switch_to.alert
            alert.accept()
            print("❌ CAPTCHA Alert: Retrying...")
            return False
        except:
            pass

        # Check for error message in page
        page_source = driver.page_source
        if "Invalid Captcha" in page_source or "Please enter correct captcha" in page_source:
            print("❌ CAPTCHA Failed (page message). Retrying...")
            return False

        # Extract and display result text
        result_text = driver.find_element(By.TAG_NAME, "body").text
        if "Enrollment No" not in result_text:
            print("⚠️ No valid result found, retrying...")
            return False

        print("📄 Result fetched successfully!")
        text_area.insert(tk.END, f"✅ Roll No: {roll_number} (Semester: {semester})\n{result_text}\n\n")
        text_area.yview(tk.END)
        return True

    except Exception as e:
        print("⚠️ Error while processing:", e)
        return False


# ---------------------- TKINTER UI ----------------------
window = tk.Tk()
window.title("RGPV Result Scraper - Stable Edition")
window.geometry("900x600")

text_area = scrolledtext.ScrolledText(window, width=110, height=30, wrap=tk.WORD)
text_area.pack(padx=10, pady=10)
# ---------------------- FUNCTION TO COPY ALL DATA ----------------------
def copy_all_data():
    data = text_area.get("1.0", tk.END)
    window.clipboard_clear()
    window.clipboard_append(data)
    window.update()  # make sure clipboard is updated
    text_area.insert(tk.END, "📋 All results copied to clipboard!\n")
    text_area.yview(tk.END)



# ---------------------- MAIN PROCESSING LOGIC ----------------------
def start_processing():
    # Ask for inputs
    constant_part = simpledialog.askstring("Input", "Enter constant part of Roll No (e.g. 0103AL231):", parent=window)
    start_roll = simpledialog.askinteger("Input", "Enter starting number (e.g. 1):", parent=window)
    end_roll = simpledialog.askinteger("Input", "Enter ending number (e.g. 84):", parent=window)
    semester = simpledialog.askstring("Input", "Enter Semester (e.g. 4):", parent=window)

    # Validation
    if not (constant_part and start_roll and end_roll and semester):
        text_area.insert(tk.END, "❌ Invalid input! Please fill all fields.\n")
        text_area.yview(tk.END)
        return

    # Threaded function
    def process_all():
        for i in range(start_roll, end_roll + 1):
            roll_number = f"{constant_part}{i:03d}"
            success = False
            retries = 0

            while not success and retries < 5:
                success = process_roll_number(roll_number, semester, text_area)
                if not success:
                    retries += 1
                    text_area.insert(tk.END, f"🔄 Retrying {roll_number} (Attempt {retries})...\n")
                    text_area.yview(tk.END)
                    time.sleep(3)

            if not success:
                text_area.insert(tk.END, f"❌ Failed after 5 attempts: {roll_number}\n\n")
                text_area.yview(tk.END)

        text_area.insert(tk.END, "\n✅ All roll numbers processed!\n")
        text_area.yview(tk.END)

    threading.Thread(target=process_all, daemon=True).start()
    text_area.insert(tk.END, "\n✅ All roll numbers processed!\n")
text_area.yview(tk.END)





# ---------------------- UI BUTTON ----------------------
# ---------------------- UI BUTTONS ----------------------
btn_frame = tk.Frame(window)
btn_frame.pack(pady=15)

start_button = tk.Button(btn_frame, text="🚀 Start Processing", command=start_processing, bg="#007BFF", fg="white", padx=10, pady=6)
start_button.grid(row=0, column=0, padx=10)

copy_button = tk.Button(btn_frame, text="📋 Copy All Data", command=copy_all_data, bg="#28A745", fg="white", padx=10, pady=6)
copy_button.grid(row=0, column=1, padx=10)



# ---------------------- MAINLOOP ----------------------
try:
    window.mainloop()
finally:
    driver.quit()
