import cv2
import numpy as np
from pyzbar.pyzbar import decode
from PIL import Image

def scan_qr_from_image(image_path):
    """
    Scan QR code from an image file.
    
    Args:
        image_path (str): Path to the image containing QR code
        
    Returns:
        str: Decoded QR code data or None if no QR code found
    """
    try:
        # Read the image
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"Could not read image from {image_path}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Decode QR code
        decoded_objects = decode(gray)
        
        # Return results
        if decoded_objects:
            return decoded_objects[0].data.decode('utf-8')
        else:
            return None
    except Exception as e:
        print(f"Error scanning QR code: {e}")
        return None

def scan_qr_from_webcam():
    """
    Scan QR code from webcam in real-time.
    
    Returns:
        str: Decoded QR code data or None if user cancels
    """
    # Initialize webcam
    cap = cv2.VideoCapture(0)
    
    if not cap.isOpened():
        print("Error: Could not open webcam")
        return None
    
    print("Webcam QR scanner started. Press 'q' to quit.")
    
    while True:
        # Read frame from webcam
        ret, frame = cap.read()
        
        if not ret:
            print("Error: Could not read frame from webcam")
            break
        
        # Convert to grayscale
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # Decode QR code
        decoded_objects = decode(gray)
        
        # Draw rectangle around QR code and display data
        for obj in decoded_objects:
            # Get QR code data
            qr_data = obj.data.decode('utf-8')
            
            # Draw rectangle
            points = obj.polygon
            if len(points) > 4:
                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                cv2.polylines(frame, [hull], True, (0, 255, 0), 2)
            else:
                cv2.polylines(frame, [np.array([point for point in points], dtype=np.int32)], True, (0, 255, 0), 2)
            
            # Display data
            cv2.putText(frame, qr_data, (obj.rect.left, obj.rect.top - 10), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
            
            # Release webcam and close windows
            cap.release()
            cv2.destroyAllWindows()
            
            return qr_data
        
        # Display the frame
        cv2.imshow("QR Code Scanner", frame)
        
        # Check for 'q' key to quit
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    # Release webcam and close windows
    cap.release()
    cv2.destroyAllWindows()
    
    return None
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from PIL import Image

def scan_qr_from_image(image_path):
    """
    Scan QR code from an image file.
    
    Args:
        image_path (str): Path to the image containing QR code
        
    Returns:
        str: Decoded QR code data or None if no QR code found
    """
    try:
        # Read the image
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"Could not read image from {image_path}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Decode QR code
        decoded_objects = decode(gray)
        
        # Return results
        if decoded_objects:
            return decoded_objects[0].data.decode('utf-8')
        else:
            return None
    except Exception as e:
        print(f"Error scanning QR code: {e}")
        return None

def scan_qr_from_webcam():
    """
    Scan QR code from webcam in real-time.
    
    Returns:
        str: Decoded QR code data or None if user cancels
    """
    # Initialize webcam
    cap = cv2.VideoCapture(0)
    
    if not cap.isOpened():
        print("Error: Could not open webcam")
        return None
    
    print("Webcam QR scanner started. Press 'q' to quit.")
    
    while True:
        # Read frame from webcam
        ret, frame = cap.read()
        
        if not ret:
            print("Error: Could not read frame from webcam")
            break
        
        # Convert to grayscale
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # Decode QR code
        decoded_objects = decode(gray)
        
        # Draw rectangle around QR code and display data
        for obj in decoded_objects:
            # Get QR code data
            qr_data = obj.data.decode('utf-8')
            
            # Draw rectangle
            points = obj.polygon
            if len(points) > 4:
                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                cv2.polylines(frame, [hull], True, (0, 255, 0), 2)
            else:
                cv2.polylines(frame, [np.array([point for point in points], dtype=np.int32)], True, (0, 255, 0), 2)
            
            # Display data
            cv2.putText(frame, qr_data, (obj.rect.left, obj.rect.top - 10), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
            
            # Release webcam and close windows
            cap.release()
            cv2.destroyAllWindows()
            
            return qr_data
        
        # Display the frame
        cv2.imshow("QR Code Scanner", frame)
        
        # Check for 'q' key to quit
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    # Release webcam and close windows
    cap.release()
    cv2.destroyAllWindows()
    
    return None