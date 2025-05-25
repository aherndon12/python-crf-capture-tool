# ---------------------------
# Imports necessary libraries
# ---------------------------
import os
import re
import time
import warnings
import datetime
import pandas as pd
import img2pdf
from PyPDF2 import PdfMerger
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
from dotenv import load_dotenv

# Ignore non-critical warnings
warnings.filterwarnings("ignore")

# Load environment variables from .env file
load_dotenv()

# Get credentials from environment variables
username = os.getenv('IMEDIDATA_USERNAME')
password = os.getenv('IMEDIDATA_PASSWORD')

if not username or not password:
    raise ValueError("Environment variables IMEDIDATA_USERNAME and IMEDIDATA_PASSWORD must be set in credentials.env")

# ----------------------------------
# Authenticates the user in iMedidata
# ----------------------------------
def login(driver):
    """
    Logs into the iMedidata portal.
    1. Goes to the login page.
    2. Waits for the username/password fields, fills them with credentials.
    3. Handles potential 2FA (prompts user to continue manually).
    4. Waits until redirected to the main iMedidata page.
    """
    login_url = "https://login.imedidata.com/login"
    driver.get(login_url)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "session[username]")))
        driver.find_element(By.NAME, "session[username]").send_keys(username)
        driver.find_element(By.NAME, "session[password]").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "button[data-testid='sign_in_button']").click()

        # Wait briefly to check if 2FA is required
        time.sleep(5)

        # Check explicitly for 2FA presence
        if "2FA" in driver.page_source or "two-factor" in driver.current_url.lower():
            input("2FA detected. Complete authentication manually, then press Enter to continue.")
        
        # After login or manual 2FA, wait explicitly for the home page element
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "root")))

        print("Login successful.")
    except Exception as e:
        print(f"Login error: {e}")
        input("Complete manual login and press Enter.")

# -------------------------------
# Cleans up file names for saving
# -------------------------------
def sanitize_filename(name):
    """
    Replaces invalid filename characters with underscores and strips whitespace.
    Returns the sanitized filename string.
    """
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()

# -------------------------------------------------------
# Class that prepares pages and captures full-length shots
# -------------------------------------------------------
class PageCapturer:
    """
    A helper class that:
    1. Navigates to a page and removes certain elements (e.g., sidebars).
    2. Displays dropdown (select) options as text.
    3. Captures full-page screenshots by stitching multiple images.
    """

    def __init__(self, driver):
        self.driver = driver
        self.actual_screenshot_height = None

    def prepare_page(self, url):
        """
        Opens the given URL and removes certain UI elements (e.g., sidebars, badges).
        Also converts <select> dropdowns into text for easier reference in screenshots.
        """
        self.driver.get(url)
        time.sleep(3)
        self.driver.execute_script("""
            ['.mcc-sidebar-left', '._pendo-image', '._pendo-badge', '.sticky-bottom']
            .forEach(selector => document.querySelector(selector)?.remove());
        """)
        self.driver.execute_script("""
            document.querySelectorAll('select').forEach(select => {
                const options = Array.from(select.options)
                    .map(opt => opt.text.trim())
                    .filter(t => t !== '...');
                const span = document.createElement('span');
                span.style.cssText = 'color: red; font-size: 12px; margin-left: 10px;';
                span.textContent = 'Options: ' + options.join(' | ');
                select.parentNode.insertBefore(span, select.nextSibling);
            });
        """)
        time.sleep(1)
        self.driver.execute_script("document.documentElement.style.overflow = 'hidden';")

    def get_page_dimensions(self):
        """
        Returns total width and height of the webpage (for full-page screenshot).
        """
        return self.driver.execute_script("""
            return [
                Math.max(document.documentElement.scrollWidth, document.body.scrollWidth,
                         document.documentElement.offsetWidth, document.body.offsetWidth),
                Math.max(document.documentElement.scrollHeight, document.body.scrollHeight,
                         document.documentElement.offsetHeight, document.body.offsetHeight)
            ];
        """)

    def capture_full_page(self, crf_name, output_dir):
        """
        Scrolls through the web page in increments of the browser window height,
        captures screenshots, and stitches them into one tall image.
        Returns the final stitched PIL Image object.
        """
        full_width, full_height = self.get_page_dimensions()

        # Determine screenshot height if not already known
        if not self.actual_screenshot_height:
            test_path = os.path.join(output_dir, "test_screenshot.png")
            self.driver.save_screenshot(test_path)
            with Image.open(test_path) as img:
                self.actual_screenshot_height = img.height
            os.remove(test_path)

        num_full_scrolls = full_height // self.actual_screenshot_height
        remaining_height = full_height % self.actual_screenshot_height
        screenshots = []

        try:
            # Capture full scrolls
            for i in range(num_full_scrolls):
                self.driver.execute_script(f"window.scrollTo(0, {i * self.actual_screenshot_height})")
                time.sleep(0.5)
                temp_path = os.path.join(output_dir, f"temp_{crf_name}_{i}.png")
                self.driver.save_screenshot(temp_path)
                screenshots.append(temp_path)

            # Capture the leftover part
            if remaining_height > 0:
                self.driver.execute_script(f"window.scrollTo(0, {full_height - self.actual_screenshot_height})")
                time.sleep(0.5)
                final_path = os.path.join(output_dir, f"temp_{crf_name}_final.png")
                self.driver.save_screenshot(final_path)
                with Image.open(final_path) as img:
                    crop_y = img.height - remaining_height
                    cropped = img.crop((0, crop_y, img.width, img.height))
                    cropped.save(final_path)
                screenshots.append(final_path)

            # Stitch screenshots together
            if screenshots:
                final_image = Image.new('RGB', (full_width, full_height))
                y_offset = 0
                for idx, path in enumerate(screenshots):
                    with Image.open(path) as img:
                        if idx == len(screenshots) - 1 and remaining_height > 0:
                            img = img.crop((0, img.height - remaining_height, img.width, img.height))
                        paste_height = min(img.height, full_height - y_offset)
                        final_image.paste(img, (0, y_offset))
                        y_offset += paste_height
                        if y_offset >= full_height:
                            break
                return final_image.crop((0, 0, full_width, full_height))
        finally:
            # Clean up temp images
            for path in screenshots:
                if os.path.exists(path):
                    os.remove(path)

        return None

# -----------------------------------------------
# Main function that reads URLs, captures pages,
# converts them to PDFs, merges PDFs, and cleans up
# -----------------------------------------------
def main():
    """
    Main execution flow:
    1. Installs/updates ChromeDriver automatically.
    2. Creates a Chrome profile, logs into iMedidata.
    3. Reads a list of URLs from an Excel file, then captures each page to PNG.
    4. Converts PNG files to individual PDFs.
    5. Merges all PDFs into one combined PDF.
    6. Cleans up individual PDFs and finishes.
    """
    chromedriver_autoinstaller.install()

    options = Options()
    user_data_dir = os.path.join(os.path.dirname(__file__), 'chromeprofile')
    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--hide-scrollbars")
    options.add_argument("--force-device-scale-factor=1")

    driver = webdriver.Chrome(options=options)
    login(driver)  # Call the login function right after credentials

    today = datetime.datetime.now().strftime("%d%b%Y")
    output_dir = f"output_{today}"
    os.makedirs(output_dir, exist_ok=True)

    df = pd.read_excel('URLs.xlsx')
    pdf_files = []
    capturer = PageCapturer(driver)

    for _, row in df.iterrows():
        crf_name = sanitize_filename(row['CRF'])
        url = row['URL']
        png_path = os.path.join(output_dir, f'{crf_name}.png')
        pdf_path = os.path.join(output_dir, f'{crf_name}.pdf')

        if os.path.exists(pdf_path):
            pdf_files.append(pdf_path)
            continue

        try:
            print(f"Processing: {crf_name}")
            capturer.prepare_page(url)
            stitched = capturer.capture_full_page(crf_name, output_dir)
            if stitched:
                stitched.save(png_path)
                with open(pdf_path, "wb") as f:
                    f.write(img2pdf.convert(png_path))
                pdf_files.append(pdf_path)
                print(f"Created: {pdf_path}")
            else:
                print(f"Failed to capture: {crf_name}")
        except Exception as e:
            print(f"Error processing {crf_name}: {str(e)}")
            continue

    if pdf_files:
        merger = PdfMerger()
        for pdf in pdf_files:
            bookmark_name = os.path.splitext(os.path.basename(pdf))[0]  # Extract filename without extension for bookmark
            merger.append(pdf, outline_item=bookmark_name)
        merged_pdf_path = os.path.join(output_dir, "Rave EDC - CRF Casebook.pdf")
        merger.write(merged_pdf_path)
        merger.close()
        print(f"Merged PDF created successfully at {merged_pdf_path}")

    # Cleanup individual PDFs (optional)
    for pdf in pdf_files:
        if os.path.exists(pdf):
            os.remove(pdf)

    driver.quit()
    print("Process completed successfully!")

if __name__ == "__main__":
    main()