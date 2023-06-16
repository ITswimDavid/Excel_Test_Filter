from openpyxl import load_workbook
from PIL import Image
from io import BytesIO
import matplotlib.pyplot as plt

# Load the workbook and select the worksheet
wb = load_workbook('your_file.xlsx')
ws = wb.active

# Loop through the worksheets images
for i, img in enumerate(ws.images):
    image = Image.open(BytesIO(img.image))

    # Show the image
    plt.figure()
    plt.imshow(image)
    plt.title(f'Image {i + 1}')
    plt.axis('off')
    plt.show()
