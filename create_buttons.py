#!/usr/bin/env python3
"""
Create professional button images for QAA Quote System
"""

from PIL import Image, ImageDraw, ImageFont
import os

# Professional colors
NAVY = (27, 54, 93)
NAVY_DARK = (20, 40, 70)
SKY_BLUE = (0, 163, 224)
GREEN = (40, 167, 69)
GRAY = (100, 100, 100)
WHITE = (255, 255, 255)

def create_gradient_button(text, size, color1, color2, output_path):
    """Create a professional gradient button"""
    img = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Create gradient fill
    for y in range(size[1]):
        ratio = y / size[1]
        r = int(color1[0] * (1 - ratio) + color2[0] * ratio)
        g = int(color1[1] * (1 - ratio) + color2[1] * ratio)
        b = int(color1[2] * (1 - ratio) + color2[2] * ratio)
        draw.line([(0, y), (size[0], y)], fill=(r, g, b))

    # Round corners by masking
    mask = Image.new('L', size, 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle([(0, 0), (size[0]-1, size[1]-1)], radius=6, fill=255)
    img.putalpha(mask)

    # Redraw on masked image
    img = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([(0, 0), (size[0]-1, size[1]-1)], radius=6, fill=color1, outline=color2, width=1)

    # Add subtle gradient overlay
    for y in range(size[1]):
        ratio = y / size[1]
        alpha = int(30 * ratio)
        draw.line([(1, y), (size[0]-2, y)], fill=(0, 0, 0, alpha))

    # Add text with shadow
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 13)
    except:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size[0] - text_width) // 2
    y = (size[1] - text_height) // 2 - 1

    # Shadow
    draw.text((x+1, y+1), text, fill=(0, 0, 0, 100), font=font)
    # Main text
    draw.text((x, y), text, fill=WHITE, font=font)

    img.save(output_path, 'PNG')
    print(f"  Created: {os.path.basename(output_path)}")

def create_all_buttons(output_dir):
    """Create all professional buttons"""

    print("Creating professional buttons...")

    # Main menu navigation buttons (larger)
    menu_buttons = [
        ("500HR SERVICE", (180, 50), NAVY, NAVY_DARK, "btn_500hr.png"),
        ("CUSTOMER PROPERTY", (180, 50), NAVY, NAVY_DARK, "btn_custprop.png"),
        ("EXCHANGE / SALE", (180, 50), NAVY, NAVY_DARK, "btn_exchange.png"),
        ("QUOTE TRACKER", (180, 50), SKY_BLUE, (0, 130, 180), "btn_tracker.png"),
        ("PO TRACKING", (180, 50), SKY_BLUE, (0, 130, 180), "btn_po_tracking.png"),
        ("INFOR XA DATA", (180, 50), GRAY, (70, 70, 70), "btn_infor_xa.png"),
        ("QUICK REFERENCE", (180, 50), GRAY, (70, 70, 70), "btn_quick_ref.png"),
    ]

    # Action buttons (standard size)
    action_buttons = [
        ("Generate PDF", (140, 38), NAVY, NAVY_DARK, "btn_generate_pdf.png"),
        ("Send Email", (140, 38), NAVY, NAVY_DARK, "btn_email.png"),
        ("Clear Form", (120, 38), GRAY, (70, 70, 70), "btn_clear.png"),
        ("← Menu", (100, 38), SKY_BLUE, (0, 130, 180), "btn_menu.png"),
        ("Update Date", (120, 38), NAVY, NAVY_DARK, "btn_update.png"),
        ("+ Reminder", (120, 38), (230, 126, 34), (200, 100, 25), "btn_add_reminder.png"),
        ("Close Quote", (120, 38), GREEN, (30, 140, 55), "btn_close_quote.png"),
        ("Update Timestamp", (150, 38), NAVY, NAVY_DARK, "btn_timestamp.png"),
    ]

    for text, size, c1, c2, filename in menu_buttons + action_buttons:
        create_gradient_button(text, size, c1, c2, os.path.join(output_dir, filename))

    print("\nAll buttons created!")

if __name__ == "__main__":
    create_all_buttons("/home/user/Test")
