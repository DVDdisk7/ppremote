import customtkinter as ctk
import threading
import socket
import qrcode
from PIL import Image, ImageDraw
import os
import sys
import queue

from web_server import run_server, stop_server

class QueueLogger:
    def __init__(self, q):
        self.q = q
    def write(self, text):
        self.q.put(text)
    def flush(self):
        pass
    def isatty(self):
        return False

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PowerPoint Remote Host")
        
        icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), 'icon.ico')
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
            
        self.geometry("500x200")
        self.resizable(False, False)
        
        self.protocol('WM_DELETE_WINDOW', self.on_closing)

        # Default Port
        self.port = 5432
        self.server_thread = None

        # Top container for Left (Controls) and Right (QR Code)
        self.top_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_frame.pack(fill="x", padx=10, pady=10)

        # Left Frame
        self.left_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        self.left_frame.pack(side="left", fill="both", expand=True, padx=10)

        # Right Frame
        self.right_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        self.right_frame.pack(side="right", fill="both", padx=10)

        # Title
        self.title_label = ctk.CTkLabel(self.left_frame, text="PowerPoint Remote", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(anchor="w", pady=(0, 5))

        # IP Info
        self.ip_address = self.get_local_ip()
        self.ip_label = ctk.CTkLabel(self.left_frame, text=f"Local IP: {self.ip_address}", font=ctk.CTkFont(size=14))
        self.ip_label.pack(anchor="w", pady=(0, 5))

        # Group Frame for Port and Logs Toggle
        self.group_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.group_frame.pack(anchor="w", pady=(0, 10))

        # Port
        self.port_label = ctk.CTkLabel(self.group_frame, text="Port:")
        self.port_label.pack(side="left", padx=(0, 5))
        self.port_entry = ctk.CTkEntry(self.group_frame, width=50, height=24)
        self.port_entry.insert(0, str(self.port))
        self.port_entry.pack(side="left", padx=(0, 15))

        # Logs toggle
        self.logs_visible = False
        self.toggle_logs_btn = ctk.CTkButton(self.group_frame, text="Show Logs", command=self.toggle_logs, fg_color="transparent", border_width=1, text_color="gray", width=80, height=24)
        self.toggle_logs_btn.pack(side="left")

        # Control Buttons
        self.start_btn = ctk.CTkButton(self.left_frame, text="Start Server", command=self.toggle_server, fg_color="green", hover_color="darkgreen", width=120)
        self.start_btn.pack(anchor="w", pady=(0, 5))

        self.status_label = ctk.CTkLabel(self.left_frame, text="Server is stopped.", text_color="gray")
        self.status_label.pack(anchor="w", pady=(0, 0))

        # QR Code Display
        self.qr_label = ctk.CTkLabel(self.right_frame, text="")
        self.qr_label.pack(pady=10)
        
        self.log_box = ctk.CTkTextbox(self, height=150, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_box.configure(state="disabled")

        self.log_queue = queue.Queue()
        self.logger = QueueLogger(self.log_queue)
        
        sys.stdout = self.logger
        sys.stderr = self.logger
        self.process_log_queue()
        
        # Initial QR render
        self.update_qr()

        # Update QR on port change
        self.port_entry.bind("<KeyRelease>", lambda event: self.update_qr())

    def toggle_logs(self):
        self.logs_visible = not self.logs_visible
        if self.logs_visible:
            self.geometry("500x380")
            self.toggle_logs_btn.configure(text="Hide Logs")
            self.log_box.pack(pady=(0,10), padx=20, fill="both", expand=True)
        else:
            self.geometry("500x200")
            self.toggle_logs_btn.configure(text="Show Logs")
            self.log_box.pack_forget()

    def process_log_queue(self):
        while not self.log_queue.empty():
            try:
                text = self.log_queue.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", text)
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
            except queue.Empty:
                break
        self.after(100, self.process_log_queue)

    def get_local_ip(self):
        try:
            # Create a socket to find local IP
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    def update_qr(self):
        try:
            self.port = int(self.port_entry.get())
        except ValueError:
            self.port = 5432
            
        url = f"http://{self.ip_address}:{self.port}"
        qr = qrcode.QRCode(version=1, box_size=8, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        # Save temporarily to load in ctk
        temp_dir = os.environ.get('TEMP', os.path.dirname(__file__))
        img_path = os.path.join(temp_dir, "ppremote_qr.png")
        img.save(img_path)
        
        my_image = ctk.CTkImage(light_image=Image.open(img_path),
                                dark_image=Image.open(img_path),
                                size=(150, 150))
        self.qr_label.configure(image=my_image)

    def check_port(self, port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            return s.connect_ex(('0.0.0.0', port)) == 0

    def toggle_server(self):
        if self.server_thread is None or not self.server_thread.is_alive():
            try:
                self.port = int(self.port_entry.get())
            except ValueError:
                self.port = 5432
                
            if self.check_port(self.port):
                self.status_label.configure(text=f"Error: Port {self.port} is in use!", text_color="red")
                return

            self.update_qr()
            self.status_label.configure(text=f"Server running at http://{self.ip_address}:{self.port}", text_color="green")
            self.start_btn.configure(text="Stop Server", fg_color="red", hover_color="darkred")
            self.port_entry.configure(state="disabled")
            
            # Start web server in a daemon thread
            self.server_thread = threading.Thread(target=run_server, args=("0.0.0.0", self.port), daemon=True)
            self.server_thread.start()
        else:
            stop_server()
            self.status_label.configure(text="Server is stopped.", text_color="gray")
            self.start_btn.configure(text="Start Server", fg_color="green", hover_color="darkgreen")
            self.port_entry.configure(state="normal")

    def on_closing(self):
        stop_server()
        try:
            from web_server import ppt
            ppt.cleanup_cache()
        except Exception as e:
            print("Error during cleanup:", e)
        self.quit()
        os._exit(0) # Forcefully kill any remaining daemon threads like uvicorn

if __name__ == "__main__":
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
