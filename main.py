import time
from PIL import Image, ImageOps, ImageEnhance
from escpos.printer import Serial
import os
import keyboard

# CONFIGURATION
WIDTH = 496
HEIGHT = 279

COM_PORT = "COM3"
BAUDRATE = 9600
THING_IMAGE_PATH = "mark.png"

PRINT_SPEED_MM_S = 100
DPI = 180
MM_PER_INCH = 25.4

BRIGHTNESS = 1.6
CONTRAST = 1.6
USE_DITHER = True


def set_print_density(printer, density=1, break_time=0):
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


def load_static_print_image():
    if not os.path.exists(THING_IMAGE_PATH):
        print(f"ERROR: {THING_IMAGE_PATH} not found!")
        return None

    img = Image.open(THING_IMAGE_PATH).convert("RGB")
    img = img.resize((WIDTH, HEIGHT), Image.LANCZOS)
    return img


# PRINT IMAGE ON PRINTER
def print_static_image():
    static_img = load_static_print_image()
    if static_img is None:
        return

    temp_path = "temp_print.png"

    # Brightness + Contrast
    static_img = ImageEnhance.Brightness(static_img).enhance(BRIGHTNESS)
    static_img = ImageEnhance.Contrast(static_img).enhance(CONTRAST)

    # Image dithering
    if USE_DITHER:
        bw_img = static_img.convert("1", dither=Image.FLOYDSTEINBERG)
    else:
        bw_img = static_img.convert("1", dither=Image.NONE)

    # Image padding
    bw_img = ImageOps.expand(bw_img, border=(0, 0, 0, 30), fill=255)

    bw_img.save(temp_path)

    try:
        set_print_density(printer, density=1)

        est_time = estimate_feed_time(bw_img.height)
        print(f"Printing... est {est_time:.2f}s")

        printer.image(temp_path)
        printer.text("\n\n\n\n\n\n")

        time.sleep(est_time)
        print("Done printing.")

    except Exception as e:
        print(f"ERROR printing: {e}")

    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


# PRINTER INIT
try:
    printer = Serial(devfile=COM_PORT, baudrate=BAUDRATE, timeout=1)
except Exception as e:
    print(f"ERROR: Could not open printer: {e}")
    printer = None


# HOTKEYS
def hotkey_print():
    if printer:
        print_static_image()
    else:
        print("Printer not initialized.")

keyboard.add_hotkey("ctrl+p", hotkey_print)
print("Press CTRL + P to print the image. CTRL + Q to quit.")

keyboard.wait("ctrl+q")

# CLEANUP
if printer:
    try:
        printer.close()
    except:
        pass

print("Exited.")

# more examples:
# printer.text("...") will print out text on the printer
# printer.cut() will cut the paper
# see for full documentation: https://python-escpos.readthedocs.io/en/latest
