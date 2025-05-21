import os
import argparse
import pyperclip
from qr_generator import generate_qr_code
from qr_scanner import scan_qr_from_image, scan_qr_from_webcam

def main():
    parser = argparse.ArgumentParser(description="QR Code Generator and Scanner")
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # Generate QR code command
    generate_parser = subparsers.add_parser("generate", help="Generate a QR code")
    generate_parser.add_argument("data", help="Text or URL to encode in the QR code")
    generate_parser.add_argument("-o", "--output", default="qrcode.png", help="Output file path (default: qrcode.png)")
    generate_parser.add_argument("-s", "--size", type=int, default=10, help="Size of QR code (default: 10)")
    generate_parser.add_argument("-b", "--border", type=int, default=4, help="Border size (default: 4)")
    
    # Scan QR code from image command
    scan_image_parser = subparsers.add_parser("scan-image", help="Scan QR code from an image")
    scan_image_parser.add_argument("image_path", help="Path to the image containing QR code")
    scan_image_parser.add_argument("-c", "--copy", action="store_true", help="Copy result to clipboard")
    scan_image_parser.add_argument("-o", "--output", help="Save result to file")
    
    # Scan QR code from webcam command
    scan_webcam_parser = subparsers.add_parser("scan-webcam", help="Scan QR code from webcam")
    scan_webcam_parser.add_argument("-c", "--copy", action="store_true", help="Copy result to clipboard")
    scan_webcam_parser.add_argument("-o", "--output", help="Save result to file")
    
    # GUI command
    gui_parser = subparsers.add_parser("gui", help="Launch GUI")
    
    args = parser.parse_args()
    
    if args.command == "generate":
        try:
            output_path = generate_qr_code(args.data, args.output, args.size, args.border)
            print(f"QR code generated successfully and saved to {output_path}")
            
            # Open the generated QR code
            abs_path = os.path.abspath(output_path)
            print(f"QR code available at: {abs_path}")
            os.startfile(abs_path)  # Opens the file with the default application
        except Exception as e:
            print(f"Error generating QR code: {e}")
    
    elif args.command == "scan-image":
        try:
            result = scan_qr_from_image(args.image_path)
            if result:
                print(f"QR code content: {result}")
                
                # Copy to clipboard if requested
                if args.copy:
                    pyperclip.copy(result)
                    print("Result copied to clipboard")
                
                # Save to file if requested
                if args.output:
                    with open(args.output, "w") as f:
                        f.write(result)
                    print(f"Result saved to {args.output}")
            else:
                print("No QR code found in the image")
        except Exception as e:
            print(f"Error scanning QR code: {e}")
    
    elif args.command == "scan-webcam":
        try:
            print("Opening webcam for QR code scanning...")
            result = scan_qr_from_webcam()
            if result:
                print(f"QR code content: {result}")
                
                # Copy to clipboard if requested
                if args.copy:
                    pyperclip.copy(result)
                    print("Result copied to clipboard")
                
                # Save to file if requested
                if args.output:
                    with open(args.output, "w") as f:
                        f.write(result)
                    print(f"Result saved to {args.output}")
            else:
                print("No QR code scanned or scanning cancelled")
        except Exception as e:
            print(f"Error scanning QR code: {e}")
    
    elif args.command == "gui":
        try:
            from gui import run_gui
            run_gui()
        except ImportError:
            print("GUI module not found. Make sure you have implemented the GUI.")
    
    else:
        parser.print_help()

if __name__ == "__main__":
    main()