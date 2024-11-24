from PIL import Image, ImageEnhance
import pytesseract
import numpy as np
import cv2
import pyperclip  # To copy the text to the clipboard

# Set Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def process_image_and_copy_to_clipboard(image_path):
    """
    Processes an image for OCR, extracts the text, formats it for Excel,
    and copies the formatted text to the clipboard.

    """
    try:
        # Load the image
        image = Image.open(image_path)

        # Convert to grayscale
        grayscale = image.convert("L")

        # Enhance contrast
        enhancer = ImageEnhance.Contrast(grayscale)
        contrast_image = enhancer.enhance(3.0)

        # Convert to binary
        img_array = np.array(contrast_image)
        _, binary_image = cv2.threshold(img_array, 150, 255, cv2.THRESH_BINARY)

        # Convert back to PIL Image for OCR
        preprocessed_image = Image.fromarray(binary_image)

        # Perform OCR
        extracted_text = pytesseract.image_to_string(preprocessed_image)

        # Format the text for Excel (tab-delimited)
        formatted_text = extracted_text.replace("\n", "\t")

        # Copy the formatted text to the clipboard
        pyperclip.copy(formatted_text)
        print("Text has been copied to the clipboard in a format suitable for Excel.")
    except Exception as e:
        print(f"An error occurred: {e}")
