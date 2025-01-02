import cv2
from docx import Document
from docx.shared import Inches
from tkinter import Tk, filedialog, Button, Label, messagebox
import numpy as np
from io import BytesIO
from PIL import Image

def extract_screenshots_and_create_document(video_path, interval, document_name):
    """Extracts frames from the video and directly inserts them into a Word document."""
    # Initialize Word document
    doc = Document()
    # doc.add_heading('Video Screenshots', level=1)

    # Open the video
    video = cv2.VideoCapture(video_path)
    fps = int(video.get(cv2.CAP_PROP_FPS))
    frame_interval = interval * fps
    count = 0
    success, frame = video.read()

    # Process video frames
    while success:
        frame_number = int(video.get(cv2.CAP_PROP_POS_FRAMES))
        if frame_number % frame_interval == 0:
            # Convert frame to an image in memory
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil_image = Image.fromarray(frame_rgb)
            buffer = BytesIO()
            pil_image.save(buffer, format="JPEG")
            buffer.seek(0)

            # Add the image to the Word document
            # doc.add_paragraph(f'Image {count + 1}')
            doc.add_picture(buffer, width=Inches(7))
            count += 1

        success, frame = video.read()

    video.release()

    # Save the Word document
    doc.save(document_name)
    print(f"Document saved as {document_name}")

def select_video():
    """Select a video file and process it."""
    video_path = filedialog.askopenfilename(
        title="Select Video File",
        filetypes=[("MP4 Files", "*.mp4"), ("All Files", "*.*")]
    )
    if not video_path:
        messagebox.showwarning("No File", "No video file selected!")
        return

    # Ask for output file name and location
    document_path = filedialog.asksaveasfilename(
        title="Save Word Document As",
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")]
    )
    if not document_path:
        messagebox.showwarning("No File", "No save location selected!")
        return

    # Process the video
    extract_screenshots_and_create_document(video_path, interval, document_path)

    messagebox.showinfo("Success", f"Processing complete. Document saved at:\n{document_path}")

# Tkinter Frontend
root = Tk()
root.title("Video Screenshot Extractor")
root.geometry("400x200")

interval = 5  # Seconds

Label(root, text="Video Screenshot Extractor", font=("Arial", 16)).pack(pady=10)
Button(root, text="Select Video and Process", command=select_video).pack(pady=10)
Button(root, text="Exit", command=root.quit).pack(pady=10)

root.mainloop()
