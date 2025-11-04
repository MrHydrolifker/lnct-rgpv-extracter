from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
from PIL import Image
import pytesseract
import time
import tkinter as tk
from tkinter import scrolledtext
import threading

# ✅ Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\softwarelab153\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# ✅ Selenium Chrome Setup
driver = webdriver.Chrome()
driver.set_window_size(1200, 1000)
wait = WebDriverWait(driver, 20)

# ✅ Configurations
semester = "4"
start_enrollment = 1
end_enrollment = 5  # Adjust as needed
enrollment_prefix = "0103AL231"
MAX_CAPTCHA_RETRIES = 5

# ---------------------- TKINTER UI ----------------------
window = tk.Tk()
window.title("RGPV Result Scraper - Stable Edition")
window.geometry("900x600")

text_area = scrolledtext.ScrolledText(window, width=110, height=30, wrap=tk.WORD)
text_area.pack(padx=10, pady=10)

# ---------------------- Function to process a single roll number ----------------------
def process_roll_number(roll_number, semester):
    attempts = 0
    success = False

    while not success and attempts < MAX_CAPTCHA_RETRIES:
        attempts += 1
        try:
            # Open page fresh for each enrollment or retry
            driver.get("https://result.rgpv.ac.in/Result/ProgramSelect.aspx")
            time.sleep(2)

            # Select B.Tech
            btech_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='radlstProgram_1']")))
            btech_label.click()
            time.sleep(1)

            # Fill Enrollment Number
            roll_box = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtrollno")))
            roll_box.clear()
            roll_box.send_keys(roll_number)
            time.sleep(1)

            # Select Semester
            semester_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_drpSemester")))
            Select(semester_dropdown).select_by_visible_text(semester)
            time.sleep(1)

            # CAPTCHA
            captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'CaptchaImage.axd')]")))
            captcha_img.screenshot("captcha.png")
            img = Image.open("captcha.png").convert("L")
            img = img.point(lambda x: 0 if x < 140 else 255, '1')
            config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
            captcha_text = pytesseract.image_to_string(img, config=config).strip()
            captcha_text = ''.join(c for c in captcha_text if c.isalnum())

            # Fill CAPTCHA
            captcha_box = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TextBox1")))
            captcha_box.clear()
            captcha_box.send_keys(captcha_text)
            time.sleep(3)

            # Click View Result
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='View Result']")))
            submit_btn.click()
            time.sleep(3)

            # ---------------------- CAPTCHA / RESULT CHECK ----------------------
            time.sleep(5)  # Wait for page to load

            # Check for CAPTCHA alert
            try:
                alert = driver.switch_to.alert
                alert.accept()
                text_area.insert(tk.END, f"❌ CAPTCHA Alert: Retrying {roll_number} (Attempt {attempts})...\n")
                text_area.yview(tk.END)
                continue
            except NoAlertPresentException:
                pass

            # Check for error message in page
            page_source = driver.page_source
            if "Invalid Captcha" in page_source or "Please enter correct captcha" in page_source:
                text_area.insert(tk.END, f"❌ CAPTCHA Failed (page message). Retrying {roll_number} (Attempt {attempts})...\n")
                text_area.yview(tk.END)
                continue

            # Extract and display result text
            result_text = driver.find_element(By.TAG_NAME, "body").text
            if "Enrollment No" not in result_text:
                text_area.insert(tk.END, f"⚠️ No valid result found for {roll_number}. Retrying...\n")
                text_area.yview(tk.END)
                continue

            text_area.insert(tk.END, f"✅ Roll No: {roll_number} (Semester: {semester})\n{result_text}\n\n")
            text_area.yview(tk.END)
            success = True

        except Exception as e:
            text_area.insert(tk.END, f"⚠️ Error on attempt {attempts} for {roll_number}: {e}\n")
            text_area.yview(tk.END)
            time.sleep(3)

    if not success:
        text_area.insert(tk.END, f"❌ Failed to get result for {roll_number} after {MAX_CAPTCHA_RETRIES} attempts.\n\n")
        text_area.yview(tk.END)

# ---------------------- Main Processing Loop ----------------------
def start_processing():
    for i in range(start_enrollment, end_enrollment + 1):
        roll_number = f"{enrollment_prefix}{str(i).zfill(3)}"
        process_roll_number(roll_number, semester)

    text_area.insert(tk.END, "\n✅ All roll numbers processed!\n")
    text_area.yview(tk.END)

# ---------------------- Run Processing in Thread ----------------------
threading.Thread(target=start_processing, daemon=True).start()

window.mainloop()
driver.quit()
