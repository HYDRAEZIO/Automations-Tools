import os
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch

def create_pdf_from_images(directory_path, output_pdf):
    # Set up the PDF file
    c = canvas.Canvas(output_pdf, pagesize=A4)
    width, height = A4

    # Get list of all image files in the directory
    image_files = [f for f in os.listdir(directory_path) if f.endswith(('.png', '.jpg', '.jpeg'))]
    image_files.sort()  # Optional: Sort the images alphabetically

    i = 0
    while i < len(image_files):
        # First image
        image_file1 = image_files[i]
        image_path1 = os.path.join(directory_path, image_file1)
        img1 = Image.open(image_path1)
        img1_width, img1_height = img1.size

        # Calculate scaling factor to fit the images within half the A4 page height each
        scale_factor = min((width - 2 * inch) / img1_width, (height / 2 - 1.5 * inch) / img1_height)
        new_width = img1_width * scale_factor
        new_height = img1_height * scale_factor

        # Place the first image at the top half of the page
        x_position1 = (width - new_width) / 2
        y_position1 = height - new_height - inch  # 1 inch from the top for spacing
        c.drawImage(image_path1, x_position1, y_position1, new_width, new_height)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(width / 2, y_position1 - 20, image_file1)

        # Check if there is a second image for the same page
        if i + 1 < len(image_files):
            image_file2 = image_files[i + 1]
            image_path2 = os.path.join(directory_path, image_file2)
            img2 = Image.open(image_path2)
            img2_width, img2_height = img2.size

            # Apply the same scale factor to the second image
            new_width2 = img2_width * scale_factor
            new_height2 = img2_height * scale_factor

            # Place the second image in the bottom half of the page
            x_position2 = (width - new_width2) / 2
            y_position2 = (height / 2) - new_height2 - inch / 2
            c.drawImage(image_path2, x_position2, y_position2, new_width2, new_height2)
            c.setFont("Helvetica-Bold", 14)
            c.drawCentredString(width / 2, y_position2 - 20, image_file2)

        # Add a new page for the next pair of images
        c.showPage()
        
        # Increment by 2 to process the next pair of images
        i += 2

    # Save the PDF
    c.save()
    print(f"PDF created successfully: {output_pdf}")

# Specify the directory containing images and output PDF file name
directory_path = "D:/Reports/Broker/October"
output_pdf = "D:/Reports/Broker/October/Report.pdf"

# Run the function
create_pdf_from_images(directory_path, output_pdf)