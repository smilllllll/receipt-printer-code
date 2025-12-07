import time
from mss import mss
from PIL import Image, ImageOps, ImageDraw, ImageEnhance
from datetime import datetime
import os
import cv2
import numpy as np
from escpos.printer import Serial
import win32api
import pytesseract
import keyboard

# === CONFIGURATION ===
WIDTH = 496
HEIGHT = 279
FPS = 2               # capture rate
SAVE_TO_FOLDER = False
SAVE_FOLDER = "screenshots"

COM_PORT = "COM3"
BAUDRATE = 9600

FRAME_INTERVAL = 2 / FPS

# Print speed estimate (for sync)
PRINT_SPEED_MM_S = 100   # assume ~100 mm/s feed rate
DPI = 180
MM_PER_INCH = 25.4

if SAVE_TO_FOLDER:
    os.makedirs(SAVE_FOLDER, exist_ok=True)

print(f"Starting screenshot bot: {WIDTH}x{HEIGHT}px at {FPS} FPS")
print("Hotkeys: CTRL+T toggle auto-print, CTRL+P print once, CTRL+O OCR, CTRL+Q quit")

sct = mss()

try:
    printer = Serial(devfile=COM_PORT, baudrate=BAUDRATE, timeout=1)
except Exception as e:
    print(f"ERROR: Could not open printer on {COM_PORT}: {e}")
    printer = None


def set_print_density(printer, density=1, break_time=0):
    if not (0 <= density <= 31):
        raise ValueError("Density must be between 0 and 31")
    if not (0 <= break_time <= 31):
        raise ValueError("Break time must be between 0 and 31")

    cmd = bytes([
        0x1D, 0x28, 0x45,
        0x03, 0x00,
        0x00,
        density,
        break_time
    ])
    printer._raw(cmd)


def estimate_feed_time(image_height_px, dpi=DPI, speed_mm_s=PRINT_SPEED_MM_S):
    mm_per_px = MM_PER_INCH / dpi
    height_mm = image_height_px * mm_per_px
    return height_mm / speed_mm_s

def get_image_brightness(pil_image):
    gray = pil_image.convert("L")  # grayscale
    np_gray = np.array(gray, dtype=np.uint8)
    return np_gray.mean()


def print_to_thermal_printer(pil_image, printer, linebreak=True):
    if printer is None:
        print("Printer is not initialized.")
        return
    temp_path = "temp_print.png"

    # ---- Dynamic Brightness Detection ----
    avg_brightness = get_image_brightness(pil_image)
    print(f"[BRIGHTNESS] Average = {avg_brightness:.1f}")

    if avg_brightness < 50:
        brightness_factor = 7
        contrast_factor = 5
        print("[BRIGHTNESS] Very dark → high boost")
    elif avg_brightness < 90:
        brightness_factor = 1.8
        contrast_factor = 1.15
        print("[BRIGHTNESS] Slightly dark → moderate boost")
    else:
        brightness_factor = 1.6
        contrast_factor = 1.2
        print("[BRIGHTNESS] Normal brightness → no boost")

    enhancer = ImageEnhance.Brightness(pil_image)
    bright = enhancer.enhance(brightness_factor)

    enhancer = ImageEnhance.Contrast(bright)
    adjusted = enhancer.enhance(contrast_factor)

    bw_img = adjusted.convert("1")
    bw_img = ImageOps.expand(bw_img, border=(0, 0, 0, 30), fill=255)
    bw_img.save(temp_path)

    try:
        set_print_density(printer, density=1)
        est_time = estimate_feed_time(bw_img.height)
        print(f"[SYNC] Printing... est. {est_time:.2f}s")

        printer.image(temp_path)
        printer.text("\n")
        if linebreak:
            printer.text("\n\n\n\n\n\n")

        time.sleep(est_time)
        print("Done printing.")
    except Exception as e:
        print(f"ERROR printing: {e}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)



def add_cursor_to_image(pil_image, monitor):
    x, y = win32api.GetCursorPos()
    mon_x, mon_y, mon_w, mon_h = monitor["left"], monitor["top"], monitor["width"], monitor["height"]
    scale_x = WIDTH / mon_w
    scale_y = HEIGHT / mon_h
    new_x = int((x - mon_x) * scale_x)
    new_y = int((y - mon_y) * scale_y)

    draw = ImageDraw.Draw(pil_image)
    cursor_radius = 10
    draw.ellipse(
        (new_x - cursor_radius, new_y - cursor_radius, new_x + cursor_radius, new_y + cursor_radius),
        fill="red"
    )
    return pil_image


def extract_text_from_screenshot(pil_image, lang="eng"):
    try:
        text = pytesseract.image_to_string(pil_image, lang=lang)
        print("=== OCR RESULT ===")
        print(text.strip())
        print("==================")

        coords_line = None
        for line in text.splitlines():
            if line.strip().startswith("Block:"):
                coords_line = line.strip()
                break

        if coords_line:
            print(f"[FOUND] {coords_line}")
            if printer:
                printer.text("\n\n\n")
                printer.text(coords_line)
                printer.text("\n\n\n")
            return coords_line
        else:
            print("[NOT FOUND] No coords detected.")
            return None
    except Exception as e:
        print(f"ERROR during OCR: {e}")
        return None


# === GLOBAL STATE ===
printing_enabled = False
running = True


# === HOTKEY HANDLERS ===
def toggle_printing():
    global printing_enabled
    printing_enabled = not printing_enabled
    print(f"Auto-printing {'ENABLED' if printing_enabled else 'DISABLED'}")


def manual_print():
    global last_img
    if last_img is not None:
        print("Manual print triggered.")
        print_to_thermal_printer(last_img, printer, linebreak=True)


def ocr_action():
    global last_full_img
    if last_full_img is not None:
        extract_text_from_screenshot(last_full_img, lang="eng")


def quit_program():
    global running
    running = False
    print("Quitting...")


# === REGISTER HOTKEYS ===
keyboard.add_hotkey("ctrl+t", toggle_printing)
keyboard.add_hotkey("ctrl+p", manual_print)
keyboard.add_hotkey("ctrl+o", ocr_action)
keyboard.add_hotkey("ctrl+q", quit_program)


# === MAIN LOOP ===
last_img = None
last_full_img = None

try:
    next_frame_time = time.time()
    while running:
        start_time = time.time()

        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)

        last_full_img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

        img_resized = last_full_img.resize((WIDTH, HEIGHT), Image.LANCZOS)

        # --- CROP BOTTOM 20px ---
        top_crop = 5
        bottom_crop = 10
        crop_height = HEIGHT - top_crop - bottom_crop

        img_cropped = img_resized.crop((0, top_crop, WIDTH, top_crop + crop_height))

        img_with_cursor = add_cursor_to_image(img_cropped.copy(), monitor)
        last_img = img_with_cursor

        img_np = np.array(img_with_cursor, dtype=np.uint8)
        cv2.imshow("Latest Screenshot", img_np)

        if SAVE_TO_FOLDER:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_path = os.path.join(SAVE_FOLDER, f"mc_{timestamp}.png")
            img_with_cursor.save(file_path)
            print(f"Saved: {file_path}")

        if printing_enabled:
            print_to_thermal_printer(img_with_cursor, printer, linebreak=False)

        next_frame_time += FRAME_INTERVAL
        sleep_time = next_frame_time - time.time()
        if sleep_time > 0:
            time.sleep(sleep_time)
        else:
            next_frame_time = time.time()

        if cv2.waitKey(1) & 0xFF == 27:
            break

except KeyboardInterrupt:
    print("\nScreenshot bot stopped.")

finally:
    if printer:
        try:
            printer.close()
        except Exception:
            pass
    cv2.destroyAllWindows()
