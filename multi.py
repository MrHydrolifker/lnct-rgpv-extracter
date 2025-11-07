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

# ---------------------- GLOBALS ----------------------
lock = threading.Lock()
all_results = []  # shared list to store (roll_number, result_text)

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1200, 1000)
    return driver

# ---------------------- SINGLE ROLL PROCESS ----------------------
def process_roll_number(driver, roll_number, semester):
    wait = WebDriverWait(driver, 20)
    try:
        driver.get("https://result.rgpv.ac.in/Result/ProgramSelect.aspx")

        # Select B.Tech and semester
        wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='radlstProgram_1']"))).click()
        semester_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_drpSemester")))
        Select(semester_dropdown).select_by_visible_text(semester)
        time.sleep(1)

        # Capture CAPTCHA
        captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'CaptchaImage.axd')]")))
        time.sleep(3)
        captcha_img.screenshot("captcha.png")

        img = Image.open("captcha.png")
        captcha_text = pytesseract.image_to_string(img, config="--psm 7").strip()
        captcha_text = ''.join(c for c in captcha_text if c.isalnum())

        # Fill form
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno").send_keys(roll_number)
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1").send_keys(captcha_text)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='View Result']"))).click()
        time.sleep(4)

        # Check for CAPTCHA alerts or failure
        try:
            alert = driver.switch_to.alert
            alert.accept()
            return None
        except:
            pass

        page_source = driver.page_source
        if "Invalid Captcha" in page_source or "Please enter correct captcha" in page_source:
            return None

        result_text = driver.find_element(By.TAG_NAME, "body").text
        if "Enrollment No" not in result_text:
            return None

        return f"✅ Roll No: {roll_number}\n{result_text}\n\n"

    except Exception:
        return None

# ---------------------- RANGE PROCESS ----------------------
def process_range(driver, constant_part, start, end, semester, text_area):
    local_results = []
    for i in range(start, end + 1):
        roll_number = f"{constant_part}{i:03d}"
        success = False
        retries = 0

        while not success and retries < 5:
            result_text = process_roll_number(driver, roll_number, semester)
            if result_text:
                success = True
                local_results.append((roll_number, result_text))
                with lock:
                    text_area.after(0, text_area.insert, tk.END, result_text)
                    text_area.after(0, text_area.yview, tk.END)
            else:
                retries += 1
                with lock:
                    msg = f"🔄 Retrying {roll_number} (Attempt {retries})...\n"
                    text_area.after(0, text_area.insert, tk.END, msg)
                    text_area.after(0, text_area.yview, tk.END)
                time.sleep(3)

        if not success:
            with lock:
                msg = f"❌ Failed after 5 attempts: {roll_number}\n\n"
                text_area.after(0, text_area.insert, tk.END, msg)
                text_area.after(0, text_area.yview, tk.END)

    # Store thread results in global list safely
    with lock:
        all_results.extend(local_results)

# ---------------------- TKINTER UI ----------------------
window = tk.Tk()
window.title("RGPV Result Scraper - Parallel Edition (Ordered Output)")
window.geometry("900x600")

text_area = scrolledtext.ScrolledText(window, width=110, height=30, wrap=tk.WORD)
text_area.pack(padx=10, pady=10)

# ---------------------- COPY BUTTON ----------------------
def copy_all_data():
    data = text_area.get("1.0", tk.END)
    window.clipboard_clear()
    window.clipboard_append(data)
    window.update()
    text_area.insert(tk.END, "📋 All results copied to clipboard!\n")
    text_area.yview(tk.END)

# ---------------------- MAIN PROCESS ----------------------
def start_processing():
    constant_part = simpledialog.askstring("Input", "Enter constant part of Roll No (e.g. 0103AL231):", parent=window)
    start_roll = simpledialog.askinteger("Input", "Enter starting number (e.g. 1):", parent=window)
    end_roll = simpledialog.askinteger("Input", "Enter ending number (e.g. 84):", parent=window)
    semester = simpledialog.askstring("Input", "Enter Semester (e.g. 4):", parent=window)

    if not (constant_part and start_roll and end_roll and semester):
        text_area.insert(tk.END, "❌ Invalid input! Please fill all fields.\n")
        text_area.yview(tk.END)
        return

    total = end_roll - start_roll + 1
    chunk = max(1, total // 4)
    roll_ranges = [
        (start_roll, min(start_roll + chunk - 1, end_roll)),
        (start_roll + chunk, min(start_roll + 2 * chunk - 1, end_roll)),
        (start_roll + 2 * chunk, min(start_roll + 3 * chunk - 1, end_roll)),
        (start_roll + 3 * chunk, end_roll),
    ]

    text_area.insert(tk.END, "🚀 Launching 4 Chrome browsers...\n")
    text_area.yview(tk.END)

    drivers = [create_driver() for _ in range(4)]
    threads = []

    for idx, (s, e) in enumerate(roll_ranges):
        if s > e:
            continue
        t = threading.Thread(target=process_range, args=(drivers[idx], constant_part, s, e, semester, text_area))
        t.start()
        threads.append(t)

    def monitor_threads():
        for t in threads:
            t.join()

        # Sort results by roll number
        with lock:
            sorted_results = sorted(all_results, key=lambda x: x[0])

            with open("results.txt", "w", encoding="utf-8") as f:
                for _, result_text in sorted_results:
                    f.write(result_text)

        for d in drivers:
            d.quit()

        text_area.insert(tk.END, "\n✅ All roll numbers processed and saved (ordered) in results.txt\n")
        text_area.yview(tk.END)

    threading.Thread(target=monitor_threads, daemon=True).start()

# ---------------------- BUTTONS ----------------------
btn_frame = tk.Frame(window)
btn_frame.pack(pady=15)

start_button = tk.Button(btn_frame, text="🚀 Start Processing", command=start_processing, bg="#007BFF", fg="white", padx=10, pady=6)
start_button.grid(row=0, column=0, padx=10)

copy_button = tk.Button(btn_frame, text="📋 Copy All Data", command=copy_all_data, bg="#28A745", fg="white", padx=10, pady=6)
copy_button.grid(row=0, column=1, padx=10)

# ---------------------- MAIN LOOP ----------------------
try:
    window.mainloop()
finally:
    import psutil
    for proc in psutil.process_iter():
        if "chrome" in proc.name().lower() and "chromedriver" in proc.name().lower():
            proc.kill()
