from PIL import Image
import os

input_file = "image.jpg"   # Your input image
output_file = "image_resized.jpg"

# Resize
img = Image.open(input_file)
img = img.resize((590, 750), Image.LANCZOS)

# Find JPEG quality that gives ~480 KB
target_size = 200 * 1024  # 200 KB

low, high = 1, 95
best_quality = 95

while low <= high:
    mid = (low + high) // 2
    img.save(output_file, "JPEG", quality=mid, optimize=True)
    size = os.path.getsize(output_file)

    if size > target_size:
        high = mid - 1
    else:
        best_quality = mid
        low = mid + 1

img.save(output_file, "JPEG", quality=best_quality, optimize=True)

print(f"Saved as {output_file}")
print(f"Quality: {best_quality}")
print(f"Size: {os.path.getsize(output_file)/1024:.1f} KB")