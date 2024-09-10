from PIL import Image
import os

def resize_icon(input_path, output_path, sizes):
    original = Image.open(input_path)
    original = original.convert("RGBA")
    
    icons = []
    for size in sizes:
        icon = original.copy()
        icon.thumbnail((size, size), Image.LANCZOS)
        icons.append(icon)
    
    icons[0].save(output_path, format='ICO', sizes=[(icon.width, icon.height) for icon in icons])
    print(f"Icon saved as {output_path}")

if __name__ == "__main__":
    input_icon = "app_icon.ico"  # Your original icon
    output_icon = "app_icon_multi.ico"  # The new multi-size icon
    sizes = [16, 32, 48, 256]
    
    resize_icon(input_icon, output_icon, sizes)