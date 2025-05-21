import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pyperclip
from PIL import Image, ImageTk
import threading
import tempfile

from qr_generator import generate_qr_code
from qr_scanner import scan_qr_from_image, scan_qr_from_webcam

class QRCodeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Set app theme colors - mobile-inspired color scheme
        self.bg_color = "#f8f9fa"  # Light background
        self.accent_color = "#007bff"  # Primary blue
        self.text_color = "#212529"  # Dark text
        self.success_color = "#28a745"  # Success green
        self.error_color = "#dc3545"  # Error red
        self.secondary_color = "#6c757d"  # Secondary gray
        
        self.title("QR Code App")
        
        # Get screen width and height
        screen_width = self.winfo_screenwidth()
        
        # Convert 6cm to pixels (assuming 96 DPI)
        # 1 inch = 2.54 cm, 1 inch = 96 pixels
        width_in_pixels = int((11 / 2.54) * 96)
        
        # Set app dimensions to full screen height with 6cm width
        self.geometry(f"{width_in_pixels}x{self.winfo_screenheight()}")
        self.resizable(False, True)  # Allow only vertical resizing
        self.configure(bg=self.bg_color)
        
        # Set app icon if available
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
            if os.path.exists(icon_path):
                self.iconphoto(True, tk.PhotoImage(file=icon_path))
        except:
            pass
        
        # Apply a theme
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except:
            pass
        
        # Configure styles for mobile-like appearance
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color)
        self.style.configure("TLabelframe", background=self.bg_color, foreground=self.text_color)
        self.style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.text_color)
        
        # Mobile-style buttons (rounded with accent color)
        self.style.configure("Mobile.TButton", 
                           background=self.accent_color, 
                           foreground="white", 
                           font=("Arial", 12, "bold"),
                           padding=10)
        self.style.map("Mobile.TButton", 
                     background=[("active", "#0069d9"), ("pressed", "#0062cc")],
                     foreground=[("active", "white"), ("pressed", "white")])
        
        # Secondary button style
        self.style.configure("Secondary.TButton", 
                           background=self.secondary_color, 
                           foreground="white", 
                           font=("Arial", 12, "bold"),
                           padding=10)
        self.style.map("Secondary.TButton", 
                     background=[("active", "#5a6268"), ("pressed", "#545b62")],
                     foreground=[("active", "white"), ("pressed", "white")])
        
        # Create header - mobile app style
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=10, pady=(15, 5))
        
        header_label = ttk.Label(header_frame, text="QR Code App", 
                                font=("Arial", 20, "bold"), foreground=self.accent_color)
        header_label.pack()
        
        # Create tabs with mobile-style
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.generator_tab = ttk.Frame(self.notebook)
        self.scanner_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.generator_tab, text="Generate")
        self.notebook.add(self.scanner_tab, text="Scan")
        
        # Setup tabs
        self.setup_generator_tab()
        self.setup_scanner_tab()
        
        # Create footer - mobile app style
        footer_frame = ttk.Frame(self)
        footer_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        footer_label = ttk.Label(footer_frame, text="© 2023 QR Code App", 
                                font=("Arial", 8), foreground=self.secondary_color)
        footer_label.pack(side="right")
    
    def setup_generator_tab(self):
        # Mobile-style container with padding
        container = ttk.Frame(self.generator_tab)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Input section - mobile style
        ttk.Label(container, text="Enter Text or URL:", 
                font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        
        # Text input - mobile style with rounded corners effect and auto-adjusting width
        input_frame = ttk.Frame(container, borderwidth=1, relief="solid")
        input_frame.pack(fill="x", pady=(0, 15))
        
        # Create a scrollbar for horizontal scrolling
        x_scrollbar = ttk.Scrollbar(input_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")
        
        # Modified text input with horizontal scrollbar and wrap=none to allow horizontal scrolling
        self.text_input = tk.Text(input_frame, height=4, font=("Arial", 12),
                                bg="white", fg=self.text_color, relief="flat",
                                highlightthickness=0, wrap="none",
                                xscrollcommand=x_scrollbar.set)
        self.text_input.pack(fill="both", expand=True, padx=2, pady=2)
        x_scrollbar.config(command=self.text_input.xview)
        
        # Options section - combined size and border in one row
        options_frame = ttk.Frame(container)
        options_frame.pack(fill="x", pady=(0, 15))
        
        # Combined size and border options in one row
        combined_frame = ttk.Frame(options_frame)
        combined_frame.pack(fill="x", pady=5)
        
        # Size option
        size_label = ttk.Label(combined_frame, text="Size:", font=("Arial", 12))
        size_label.pack(side="left", padx=(0, 5))
        
        self.size_var = tk.IntVar(value=10)
        size_spinbox = ttk.Spinbox(combined_frame, from_=1, to=40, textvariable=self.size_var, 
                                  width=3, font=("Arial", 12))
        size_spinbox.pack(side="left", padx=(0, 15))
        
        # Border option - in same row
        border_label = ttk.Label(combined_frame, text="Border:", font=("Arial", 12))
        border_label.pack(side="left", padx=(0, 5))
        
        self.border_var = tk.IntVar(value=4)
        border_spinbox = ttk.Spinbox(combined_frame, from_=0, to=10, textvariable=self.border_var, 
                                    width=3, font=("Arial", 12))
        border_spinbox.pack(side="left")
        
        # Generate button - mobile style (full width)
        generate_button = ttk.Button(container, text="GENERATE QR CODE", 
                                   command=self.generate_qr, style="Mobile.TButton")
        generate_button.pack(fill="x", pady=15)
        
        # Result section - mobile style
        result_label = ttk.Label(container, text="QR Code Result:", font=("Arial", 12, "bold"))
        result_label.pack(anchor="w", pady=(5, 10))
        
        # QR code image container - mobile style
        image_container = ttk.Frame(container, borderwidth=1, relief="solid")
        image_container.pack(pady=(0, 15))
        
        # QR code image
        self.qr_image_label = ttk.Label(image_container, background="white")
        self.qr_image_label.pack(padx=2, pady=2)
        
        # Save button - mobile style
        save_button = ttk.Button(container, text="SAVE QR CODE", 
                               command=self.save_qr, style="Secondary.TButton")
        save_button.pack(fill="x")
        
        # Status label - mobile style
        self.generator_status = ttk.Label(container, text="", font=("Arial", 10), wraplength=350)
        self.generator_status.pack(pady=10)
    
    def setup_scanner_tab(self):
        # Mobile-style container with padding
        container = ttk.Frame(self.scanner_tab)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure grid
        container.columnconfigure(0, weight=1)  # Make column expandable
        
        # Scan options - mobile style with grid
        ttk.Label(container, text="Scan QR Code:", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(5, 10))
        
        # Scan from image button - mobile style (full width) with grid
        scan_image_button = ttk.Button(container, text="SCAN FROM IMAGE", 
                                     command=self.scan_from_image, style="Mobile.TButton")
        scan_image_button.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        # Scan from webcam button - mobile style (full width) with grid
        scan_webcam_button = ttk.Button(container, text="SCAN FROM CAMERA", 
                                      command=self.scan_from_webcam, style="Mobile.TButton")
        scan_webcam_button.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        
        # Result section - mobile style with grid
        ttk.Label(container, text="Scan Result:", font=("Arial", 12, "bold")).grid(row=3, column=0, sticky="w", pady=(5, 5))
        
        # Result text - mobile style with rounded corners effect and grid
        result_frame = ttk.Frame(container, borderwidth=1, relief="solid")
        result_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 15))
        
        # Configure result_frame to expand
        container.rowconfigure(4, weight=1)
        
        # Add horizontal scrollbar for result text
        x_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")
        
        self.result_text = tk.Text(result_frame, height=6, width=30, font=("Arial", 12),
                                 bg="white", fg=self.text_color, relief="flat",
                                 highlightthickness=0, wrap="none",
                                 xscrollcommand=x_scrollbar.set)
        self.result_text.pack(fill="both", expand=True, padx=2, pady=2)
        x_scrollbar.config(command=self.result_text.xview)
        
        # Action buttons - mobile style with grid
        copy_button = ttk.Button(container, text="COPY TO CLIPBOARD", 
                                command=self.copy_result, style="Secondary.TButton")
        copy_button.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        
        save_result_button = ttk.Button(container, text="SAVE TO FILE", 
                                       command=self.save_result, style="Secondary.TButton")
        save_result_button.grid(row=6, column=0, sticky="ew")
        
        # Status label - mobile style with grid
        self.scanner_status = ttk.Label(container, text="", font=("Arial", 10), wraplength=350)
        self.scanner_status.grid(row=7, column=0, pady=10)
    
    def generate_qr(self):
        # Get input text
        text = self.text_input.get("1.0", "end-1c").strip()
        
        if not text:
            self.generator_status.config(text="Error: Input text cannot be empty", foreground=self.error_color)
            return
        
        try:
            # Create temporary file for QR code
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_file:
                temp_path = temp_file.name
            
            # Generate QR code
            size = self.size_var.get()
            border = self.border_var.get()
            generate_qr_code(text, temp_path, size, border)
            
            # Display QR code
            self.display_qr_image(temp_path)
            
            # Save path for later use
            self.current_qr_path = temp_path
            
            self.generator_status.config(text="QR code generated successfully", foreground=self.success_color)
        except Exception as e:
            self.generator_status.config(text=f"Error: {str(e)}", foreground=self.error_color)
    
    def display_qr_image(self, image_path):
        # Open and resize image
        img = Image.open(image_path)
        img = img.resize((250, 250), Image.LANCZOS)
        
        # Convert to PhotoImage
        photo = ImageTk.PhotoImage(img)
        
        # Update label
        self.qr_image_label.config(image=photo)
        self.qr_image_label.image = photo  # Keep a reference
    
    def save_qr(self):
        if not hasattr(self, 'current_qr_path'):
            self.generator_status.config(text="Error: No QR code generated yet", foreground=self.error_color)
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Copy the temporary file to the selected location
            img = Image.open(self.current_qr_path)
            img.save(file_path)
            
            self.generator_status.config(text=f"QR code saved to {file_path}", foreground=self.success_color)
        except Exception as e:
            self.generator_status.config(text=f"Error saving QR code: {str(e)}", foreground=self.error_color)
    
    def scan_from_image(self):
        # Ask for image file
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            # Scan QR code
            result = scan_qr_from_image(file_path)
            
            if result:
                # Display result
                self.result_text.delete("1.0", "end")
                self.result_text.insert("1.0", result)
                self.scanner_status.config(text="QR code scanned successfully", foreground=self.success_color)
            else:
                self.scanner_status.config(text="No QR code found in the image", foreground=self.error_color)
        except Exception as e:
            self.scanner_status.config(text=f"Error scanning QR code: {str(e)}", foreground=self.error_color)
    
    def scan_from_webcam(self):
        # Disable buttons during scanning
        self.scanner_status.config(text="Opening camera... Press 'q' to cancel", foreground=self.accent_color)
        self.update()
        
        # Run webcam scanning in a separate thread to avoid freezing the GUI
        def scan_thread():
            try:
                result = scan_qr_from_webcam()
                
                # Update GUI in the main thread
                self.after(0, lambda: self.handle_webcam_result(result))
            except Exception as e:
                self.after(0, lambda: self.scanner_status.config(text=f"Error: {str(e)}", foreground=self.error_color))
        
        threading.Thread(target=scan_thread, daemon=True).start()
    
    def handle_webcam_result(self, result):
        if result:
            # Display result
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", result)
            self.scanner_status.config(text="QR code scanned successfully", foreground=self.success_color)
        else:
            self.scanner_status.config(text="No QR code scanned or scanning cancelled", foreground=self.error_color)
    
    def copy_result(self):
        # Get result text
        result = self.result_text.get("1.0", "end-1c").strip()
        
        if not result:
            self.scanner_status.config(text="Error: No result to copy", foreground=self.error_color)
            return
        
        # Copy to clipboard
        pyperclip.copy(result)
        self.scanner_status.config(text="Result copied to clipboard", foreground=self.success_color)
    
    def save_result(self):
        # Get result text
        result = self.result_text.get("1.0", "end-1c").strip()
        
        if not result:
            self.scanner_status.config(text="Error: No result to save", foreground=self.error_color)
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Save to file
            with open(file_path, "w") as f:
                f.write(result)
            
            self.scanner_status.config(text=f"Result saved to {file_path}", foreground=self.success_color)
        except Exception as e:
            self.scanner_status.config(text=f"Error saving result: {str(e)}", foreground=self.error_color)

def run_gui():
    app = QRCodeApp()
    app.mainloop()

if __name__ == "__main__":
    run_gui()