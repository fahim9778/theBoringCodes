import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from PIL import Image


# current flashcard groups
current_flashcard_groups = 17  # Change this as needed

current_directory = os.path.dirname(os.path.realpath(__file__))

# Create the parent folder if it doesn't exist
parent_folder = os.path.join(current_directory, "Gregmat Quant Flashcards")
if not os.path.exists(parent_folder):
    raise Exception(f"The directory {parent_folder} does not exist.")

# Set the name of the output PDF file
output_file = os.path.join(current_directory, "Gregmat Quant Flashcards Combined.pdf")

# Create a new PDF file
c = canvas.Canvas(output_file, pagesize=landscape(letter))


# Loop over the flashcard groups
for i in range(1, current_flashcard_groups+1):  # Adjust the range as needed
    # Specify the directory where the images are located
    directory = f"{parent_folder}/Flashcard_Group_{i}"
    if not os.path.exists(directory):
        raise Exception(f"The directory {directory} does not exist.")

    # Get a list of all the image files in the directory
    images = [f for f in os.listdir(directory) if f.endswith('.png')]

    # Sort the images by their number to ensure they're in the correct order
    images.sort(key=lambda f: int(f.split('_')[1].split('.')[0]))
    
    # Loop over the images and add them to the PDF
    for image_file in images:
        # Open the image file
        image = Image.open(os.path.join(directory, image_file))

        # Get the dimensions of the image
        width, height = image.size

        # Calculate the width and height in points
        width_pt = width * 0.75
        height_pt = height * 0.75

        # Calculate the aspect ratio of the image
        aspect = height / width

        # Calculate the width and height in points, maintaining aspect ratio
        width_pt = landscape(letter)[0]
        height_pt = width_pt * aspect

        # Calculate the vertical position for the image
        vertical_position = (landscape(letter)[1] - height_pt) / 2

        # Add the image to the PDF, resizing it to fit the page, and centering it vertically
        c.drawImage(os.path.join(directory, image_file), 0, vertical_position, width=width_pt, height=height_pt)

        # Move to the next page
        c.showPage()

# Save the PDF
c.save()

print(f"PDF saved as {output_file}.")
