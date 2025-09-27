from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os  # For log suppression and paths

from PIL import Image, ImageDraw, ImageFont  # For generating JPG with colors (pip install pillow)

import os
import time
from mgz.summary import Summary
from datetime import datetime, date, timedelta

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from threading import Timer

# Configuration - Edit if needed
REPLAY_BASE_DIR = os.path.expanduser('~') + r'\Games\Age of Empires 2 DE' # Base directory for AoE II DE
PROFILE_ID = None # Set this to your profile ID (e.g., '76561199310445090') if auto-detection doesn't pick the right one
# Find the savegame directory
if PROFILE_ID is None:
    profile_dirs = [d for d in os.listdir(REPLAY_BASE_DIR) if d.isdigit() and len(d) > 5] # Profile IDs are long numbers
    if len(profile_dirs) == 0:
        raise ValueError("Could not find any profile directory in Games\Age of Empires 2 DE")
    elif len(profile_dirs) > 1:
        # Auto-select the profile with the most recently modified savegame folder
        profile_times = []
        for pd in profile_dirs:
            sg_path = os.path.join(REPLAY_BASE_DIR, pd, 'savegame')
            if os.path.exists(sg_path):
                mod_time = os.path.getmtime(sg_path)
                profile_times.append((pd, mod_time))
        if profile_times:
            profile_times.sort(key=lambda x: x[1], reverse=True) # Newest first
            PROFILE_ID = profile_times[0][0]
        else:
            raise ValueError("No valid savegame folders found in profiles. Set PROFILE_ID manually.")
    else:
        PROFILE_ID = profile_dirs[0]
# Safeguard: Ensure PROFILE_ID is a string
PROFILE_ID = str(PROFILE_ID)
REPLAY_DIR = os.path.join(REPLAY_BASE_DIR, PROFILE_ID, 'savegame')

PDF_FOLDER = os.path.join(os.path.dirname(__file__), 'chat_logs')
os.makedirs(PDF_FOLDER, exist_ok=True)

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        suffix = 'th'
    else:
        suffix = ['th', 'st', 'nd', 'rd', 'th'][min(n % 10, 4)]
    return str(n) + suffix

def get_todays_replays():
    now = datetime.now()
    if now.hour >= 7:
        start_time = now.replace(hour=7, minute=0, second=0, microsecond=0)
    else:
        start_time = (now - timedelta(days=1)).replace(hour=7, minute=0, second=0, microsecond=0)
  
    todays_replays = []
    for f in os.listdir(REPLAY_DIR):
        if f.endswith('.aoe2record'):
            file_path = os.path.join(REPLAY_DIR, f)
            mod_time = os.path.getmtime(file_path)
            file_datetime = datetime.fromtimestamp(mod_time)
            if file_datetime >= start_time:
                todays_replays.append(file_path)
  
    # Sort by modification time, oldest first
    todays_replays.sort(key=os.path.getmtime)
    return todays_replays

def format_replay_info(replay_path, summary=None):
    basename = os.path.basename(replay_path)
    try:
        # Extract date and time from filename (e.g., "@2025.09.22 185103")
        parts = basename.split('@')[1].split()[0:2]
        date_str = parts[0]
        time_str = parts[1]
        dt = datetime.strptime(f"{date_str} {time_str}", "%Y.%m.%d %H%M%S")
    except Exception:
        # Fallback to modification time for solo games or invalid names
        dt = datetime.fromtimestamp(os.path.getmtime(replay_path))
    formatted_date = dt.strftime("%B %d, %Y")
    formatted_time = dt.strftime("%H:%M")
  
    duration_str = "Unknown duration"
    if summary:
        try:
            duration_ms = summary.get_duration()
            if duration_ms:
                # Adjust for 1.7x game speed to get real-time duration
                real_duration_sec = (duration_ms / 1000) / 1.7
                duration_str = time.strftime('%H:%M:%S', time.gmtime(real_duration_sec))
        except Exception:
            pass
  
    return f"{formatted_date} at {formatted_time} (Duration: {duration_str})"

def extract_all_chat(replay_path):
    retries = 10
    for attempt in range(retries):
        try:
            with open(replay_path, 'rb') as f:
                data = f.read()
            if len(data) == 0:
                raise ValueError("Empty file")
            from io import BytesIO
            summary = Summary(BytesIO(data))
          
            # Get players
            players = summary.get_players()
          
            # Color mapping (standard AoE II DE player colors by color_id)
            color_map = {
                0: 'Blue',
                1: 'Red',
                2: 'Green',
                3: 'Yellow',
                4: 'Teal',
                5: 'Purple',
                6: 'Gray',
                7: 'Orange'
            }
          
            # Get chats
            chats = summary.get_chat()
            chat_lines = []  # List of (line, color_name)
            for chat in chats:
                player = next((p for p in players if p['number'] == chat['player_number']), None)
                if player:
                    player_name = player['name']
                    color_id = player.get('color_id', -1)
                    color = color_map.get(color_id, 'Unknown')
                else:
                    player_name = 'Unknown'
                    color = 'Unknown'
              
                timestamp_ms = chat.get('timestamp', 0)
                timestamp_sec = timestamp_ms / 1000
                timestamp = time.strftime('%H:%M:%S', time.gmtime(timestamp_sec))
                line = f"[{timestamp}] {player_name} ({color}): {chat['message']}"
                chat_lines.append((line, color))
          
            if not chat_lines:
                chat_lines.append(("Only boring people in this game, there was no chat", 'White'))
          
            return chat_lines, summary
        except Exception as e:
            if attempt == retries - 1:
                return [(f"Error parsing replay {replay_path}: {e}", 'White')], None
            time.sleep(5)

# RGB colors for PDF (original ANSI-inspired colors, adjusted for visibility on black)
color_rgb = {
    'Blue': (85, 85, 255),
    'Red': (255, 85, 85),
    'Green': (85, 255, 85),
    'Yellow': (255, 255, 85),
    'Teal': (85, 255, 255),
    'Purple': (255, 85, 255),
    'Gray': (170, 170, 170),
    'Orange': (255, 165, 0),
    'Unknown': (255, 255, 255),
    'White': (255, 255, 255)
}

# Function to generate JPG for a single replay
def generate_jpg_for_replay(replay_path):
    chat_lines, summary = extract_all_chat(replay_path)
    replay_info = format_replay_info(replay_path, summary)
    lines = [replay_info] + [line for line, _ in chat_lines]

    # Settings for HD image
    font_size = 20
    font_path = "C:/Windows/Fonts/arial.ttf"  # Use a TrueType font for better quality; adjust path if needed
    font = ImageFont.truetype(font_path, font_size)
    padding = 20
    # Calculate line height based on font
    sample_bbox = font.getbbox("Ap")
    line_height = (sample_bbox[3] - sample_bbox[1]) + 10  # Height plus spacing

    # Calculate max width based on longest text
    max_text_width = max([font.getbbox(text)[2] - font.getbbox(text)[0] for text in lines])
    width = max_text_width + 2 * padding

    # Calculate height
    height = padding + line_height + line_height  # For info plus space
    height += len(chat_lines) * line_height + padding

    # Create image
    img = Image.new('RGB', (width, height), color=(0, 0, 0))
    draw = ImageDraw.Draw(img)

    y = padding
    draw.text((padding, y), replay_info, fill=(255, 255, 255), font=font)
    y += line_height * 2

    for (text, color_name) in chat_lines:
        r, g, b = color_rgb.get(color_name, (255, 255, 255))
        draw.text((padding, y), text, fill=(r, g, b), font=font)
        y += line_height

    # Get day of week and nth
    dt = datetime.fromtimestamp(os.path.getmtime(replay_path))
    day_of_week = dt.strftime("%A")
    todays_replays = get_todays_replays()
    if replay_path in todays_replays:
        nth = ordinal(todays_replays.index(replay_path) + 1)
    else:
        nth = ordinal(len(todays_replays) + 1)

    jpg_file = os.path.join(PDF_FOLDER, f"{day_of_week} {nth} game.jpg")
    img.save(jpg_file, quality=95)  # Higher quality for clarity
    return jpg_file

# Function to send JPG to WhatsApp
def send_to_whatsapp(jpg_path):
    # User inputs for WhatsApp
    group_name = "Gandhicide"  # Exact group name
    session_dir = "C:\\Users\\guyac\\Documents\\Age-of-Empires-II-Chat-Extractor to Whatsapp\\whatsapp_session"  # Session folder path
    headless_mode = True  # Set to False for first run only (to scan QR)

    # Setup Chrome options
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={session_dir}")  # Persists session
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')  # Avoid detection
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36')  # Fake user agent
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('--log-level=3')  # Suppress non-fatal browser logs (info, warnings)
    options.add_argument('--disable-gpu')  # Suppress GPU/media errors (AMD, video encoder)
    options.add_argument('--disable-accelerated-video-decode')  # Suppress media foundation errors
    if headless_mode:
        options.add_argument('--headless=new')  # Modern headless mode

    # Auto-install and use chromedriver with log suppression
    service = Service(ChromeDriverManager().install(), log_output=os.devnull)  # Discard all ChromeDriver logs
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get('https://web.whatsapp.com')
        
        # Wait for WhatsApp to load (search box as indicator)
        wait_time = 60 if not headless_mode else 30
        search_box_locator = (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
        WebDriverWait(driver, wait_time).until(EC.presence_of_element_located(search_box_locator))
        
        # Search for the group
        search_box = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(search_box_locator))
        search_box.click()
        search_box.send_keys(group_name)
        time.sleep(2)  # Wait for search results
        
        # Click the group (assumes first result is correct; use contains for partial match if needed)
        group_result = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//span[contains(@title, "{group_name}")]'))
        )
        group_result.click()
        
        # Wait for chat to load (message box)
        message_box_locator = (By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(message_box_locator))
        
        # Send the JPG attachment
        # Click attachment button
        attach_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="plus-rounded"]'))
        )
        attach_btn.click()
        time.sleep(2)  # Brief delay for menu to open
        
        # Find the image input and send file path
        image_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        image_input.send_keys(jpg_path)
        time.sleep(10)  # Increased delay to allow preview to load
        
        # Click send in the preview
        send_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@aria-label="Send"]'))
        )
        send_btn.click()
        
        print("File sent successfully!")
        print("Still listening for new files...")
        time.sleep(2)  # Brief delay to ensure send

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        driver.quit()  # Close the browser process

# File watcher handler
class ReplayHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_processed_time = 0
        self.timer = None
        self.pending_event = None

    def on_modified(self, event):
        if self.timer:
            self.timer.cancel()
        self.pending_event = event
        self.timer = Timer(30.0, self._debounced_process)
        self.timer.start()

    def _debounced_process(self):
        if self.pending_event:
            self.process(self.pending_event, 'modified')
            self.pending_event = None
        self.timer = None

    def process(self, event, event_type):
        if event.is_directory:
            return
        if event.src_path.endswith('.aoe2record'):
            print(f"A new file is spotted: {event.src_path} (event: {event_type}).")
            # Wait for file to be fully written
            size = -1
            while True:
                new_size = os.path.getsize(event.src_path)
                if new_size == size and new_size > 0:
                    break
                size = new_size
                time.sleep(2)
            print(f"Processing replay file: {event.src_path} (event: {event_type})")
            jpg_path = generate_jpg_for_replay(event.src_path)
            send_to_whatsapp(jpg_path)
            self.last_processed_time = time.time()

if __name__ == "__main__":
    event_handler = ReplayHandler()
    observer = Observer()
    observer.schedule(event_handler, path=REPLAY_DIR, recursive=False)
    observer.start()
    print(f"Watching for new replay files in {REPLAY_DIR}...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("Stopped watching for new files.")
    observer.join()