from PIL import Image, ImageDraw, ImageFont

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import pandas as pd
import os

import tempfile

def whiteBox(image):
    width, height = image.size  
    draw = ImageDraw.Draw(image)  # Use ImageDraw instead of putpixel()

    topLeftBoxCornerX = int(width / 10)
    topLeftBoxCornerY = int(8 * (height / 10))  
    bottomRightBoxCornerX = int(width - width / 10)
    bottomRightBoxCornerY = int(9 * (height / 10))  

    # Use a rectangle instead of slow putpixel()
    draw.rectangle([topLeftBoxCornerX, topLeftBoxCornerY, bottomRightBoxCornerX, bottomRightBoxCornerY], fill=(255, 255, 255))

def dynamicallyCrop(image):
    width, height = image.size
    if height > width:
        left = 0 
        upper = width / 2
        right = width
        lower = width / 2
    else:
        left = width / 2 - height / 2
        upper = 0
        right = width / 2 + height / 2
        lower = height
    return image.crop((left, upper, right, lower))

def getNames():
    df = pd.read_excel('FloorStuff.xlsx')  
    return df["First Name"].dropna().tolist()  # Convert to a list

from PIL import ImageDraw, ImageFont

def displayName(image):
    width, height = image.size
    draw = ImageDraw.Draw(image)
    
    text = "Ian (RA)"
    black = (0, 0, 0)  

    # Load font with adjustable size
    try:
        font = ImageFont.truetype("911porschav3.ttf", int(width/10))  # Adjust size here
    except IOError:
        font = ImageFont.load_default()  # Default font (size is small & fixed)

    # Get text size
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # Center text
    position = ((width - text_width) // 2, (7.9 * height // 10) + ((height // 10 - text_height) // 2))

    # Draw text
    draw.text(position, text, fill=black, font=font)

def addPicture(pdf_canvas, image, position, page_width, page_height):
    # Get the image's width and height
    img_width, img_height = image.size

    # Resize the image if necessary (you can adjust the size here)
    max_width = 260  # Maximum width for the image
    max_height = 260  # Maximum height for the image
    aspect_ratio = img_height / img_width

    if img_width > max_width:
        img_width = max_width
        img_height = int(img_width * aspect_ratio)
    
    if img_height > max_height:
        img_height = max_height
        img_width = int(img_height / aspect_ratio)

    # Position the image based on the input position argument
    if position == 1:  # Top-Left corner
        x_pos, y_pos = 0, page_height - img_height
    elif position == 2:  # Top-Right corner
        x_pos, y_pos = page_width - img_width, page_height - img_height
    elif position == 3:  # Middle-left
        x_pos, y_pos = 0, (page_height - img_height) // 2
    elif position == 4:  # Middle-right
        x_pos, y_pos = page_width - img_width, (page_height - img_height) // 2
    elif position == 5:  # Bottom-left
        x_pos, y_pos = 0, 0
    elif position == 6:  # Bottom-right
        x_pos, y_pos = page_width - img_width, 0
    else:
        raise ValueError("Invalid position number. Choose from 1, 2, 3, 4, 5, 6.")

    # Save the image as a temporary file so it can be added to the PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_image:
        temp_image_path = temp_image.name
        image.save(temp_image_path)  # Save the image as a temporary file

    # Draw the image at the specified position on the PDF canvas
    pdf_canvas.drawImage(temp_image_path, x_pos, y_pos, width=img_width, height=img_height)

    # Optionally, delete the temporary image file after drawing it on the PDF
    os.remove(temp_image_path)

def main():
    folder_path = r"C:\Users\IanKe\Downloads\Door-Dec-Creator\Porsche Pictures"  # Path to your folder
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    # Create the PDF canvas outside the loop (to keep adding to the same PDF)
    pdf_file = "DoorDecs.pdf"
    c = canvas.Canvas(pdf_file, pagesize=letter)

    page_width, page_height = letter  # The width and height of the page in points

    i = 0  # Start at 0 to properly cycle through 1-6

    for file in files:
        image_path = os.path.join(folder_path, file)  # Get full path to the image file

        with Image.open(image_path) as image:
            image = dynamicallyCrop(image)  # Crop the image as needed
            whiteBox(image)  # Add the white box
            displayName(image)  # Add the name text

            # Ensure position cycles from 1 to 6
            position = (i % 6) + 1  
            addPicture(c, image, position, page_width, page_height)

            i += 1  # Increment the counter for the next image

            # Start a new page after every 6 images
            if position == 6:
                c.showPage()

    # Save the PDF after all images are added
    c.save()

    print(f"PDF created: {pdf_file}")

if __name__ == "__main__":
    main()