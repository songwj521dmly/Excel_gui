from PIL import Image, ImageDraw, ImageFilter
import math

def create_gradient(width, height, start_color, end_color):
    base = Image.new('RGBA', (width, height), start_color)
    top = Image.new('RGBA', (width, height), end_color)
    mask = Image.new('L', (width, height))
    mask_data = []
    for y in range(height):
        for x in range(width):
            # Diagonal gradient
            p = (x + y) / (width + height)
            mask_data.append(int(255 * p))
    mask.putdata(mask_data)
    base.paste(top, (0, 0), mask)
    return base

def create_icon():
    # Define sizes for the icon
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []

    # Colors
    # Excel Green Gradient
    color1 = (33, 115, 70, 255)   # Darker Green #217346
    color2 = (60, 168, 100, 255)  # Lighter Green
    
    grid_color = (255, 255, 255, 40) # Faint white for grid
    text_color = (255, 255, 255, 255)
    shadow_color = (0, 0, 0, 80)

    for size in sizes:
        w, h = size
        
        # 1. Create Gradient Background
        img = create_gradient(w, h, color2, color1)
        draw = ImageDraw.Draw(img)

        # 2. Draw Subtle Grid (Spreadsheet look)
        # Draw 2 vertical and 2 horizontal lines
        cell_w = w / 3
        cell_h = h / 3
        
        for i in range(1, 3):
            # Vertical
            x = int(i * cell_w)
            draw.line([(x, 0), (x, h)], fill=grid_color, width=max(1, int(w/64)))
            # Horizontal
            y = int(i * cell_h)
            draw.line([(0, y), (w, y)], fill=grid_color, width=max(1, int(h/64)))

        # 3. Draw "E" with Shadow
        # Dimensions
        # Adjusting proportions to separate the bars of 'E'
        margin = int(w * 0.20)     # Reduced margin to give more vertical space
        thickness = int(w * 0.12)  # Reduced thickness to open up gaps
        
        # Helper to draw E rects
        def draw_e(offset_x, offset_y, color):
            # Vertical
            draw.rectangle([margin + offset_x, margin + offset_y, margin + thickness + offset_x, h - margin + offset_y], fill=color)
            # Top
            draw.rectangle([margin + offset_x, margin + offset_y, w - margin + offset_x, margin + thickness + offset_y], fill=color)
            # Middle
            mid_y = int(h / 2) - int(thickness / 2)
            draw.rectangle([margin + offset_x, mid_y + offset_y, w - margin + offset_x, mid_y + thickness + offset_y], fill=color)
            # Bottom
            draw.rectangle([margin + offset_x, h - margin - thickness + offset_y, w - margin + offset_x, h - margin + offset_y], fill=color)

        # Draw Shadow
        shadow_offset = max(1, int(w / 32))
        draw_e(shadow_offset, shadow_offset, shadow_color)
        
        # Draw Main Text
        draw_e(0, 0, text_color)
        
        # 4. Add a border
        draw.rectangle([0, 0, w-1, h-1], outline=(255, 255, 255, 100), width=max(1, int(w/128)))

        images.append(img)

    # Save as ICO
    images[0].save('app_icon.ico', format='ICO', sizes=sizes, append_images=images[1:])
    print("app_icon.ico created successfully.")

if __name__ == "__main__":
    create_icon()
