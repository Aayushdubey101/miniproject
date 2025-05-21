import qrcode
from PIL import Image
import os

def generate_qr_code(data, output_path="qrcode.png", size=10, border=4):
    """
    Generate a QR code from the given data and save it as a PNG image.
    
    Args:
        data (str): The text/URL to encode in the QR code
        output_path (str): Path where the QR code image will be saved
        size (int): Size of the QR code (1 to 40)
        border (int): Border size of the QR code
        
    Returns:
        str: Path to the saved QR code image
    """
    if not data:
        raise ValueError("Data cannot be empty")
    
    # Create QR code instance
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=size,
        border=border,
    )
    
    # Add data to the QR code
    qr.add_data(data)
    qr.make(fit=True)
    
    # Create an image from the QR code
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Save the image
    img.save(output_path)
    
    return output_path