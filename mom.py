import cv2
import os
from docx import Document
from docx.shared import Inches
from tkinter import Tk, filedialog

def clear_folder(folder_path):
    """Deletes all files in the specified folder."""
    if os.path.exists(folder_path):
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print(f"Cleared folder: {folder_path}")
    else:
        os.makedirs(folder_path)
        print(f"Created folder: {folder_path}")

def extract_screenshots(video_path, interval, output_folder):
    """Extracts screenshots from a video at specified intervals."""
    os.makedirs(output_folder, exist_ok=True)

    # Load video
    video = cv2.VideoCapture(video_path)
    fps = int(video.get(cv2.CAP_PROP_FPS))
    frame_interval = interval * fps
    count = 0
    success, frame = video.read()

    # Extract frames
    while success:
        frame_number = int(video.get(cv2.CAP_PROP_POS_FRAMES))
        if frame_number % frame_interval == 0:
            file_path = os.path.join(output_folder, f'screenshot_{count}.jpg')
            cv2.imwrite(file_path, frame)
            print(f"Saved {file_path}")
            count += 1
        success, frame = video.read()

    video.release()

def create_document_with_screenshots(output_folder, document_name):
    """Creates a Word document with screenshots."""
    doc = Document()
    doc.add_heading('Video Screenshots', level=1)

    for file_name in sorted(os.listdir(output_folder)):
        if file_name.endswith('.jpg'):
            file_path = os.path.join(output_folder, file_name)
            doc.add_paragraph(f'Image: {file_name}')
            # Increased the image size
            doc.add_picture(file_path, width=Inches(6))

    doc.save(document_name)
    print(f"Document saved as {document_name}")

def select_video_file():
    """Opens a file dialog to select a video file."""
    root = Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(
        title="Select Video File",
        filetypes=[("MP4 Files", "*.mp4"), ("All Files", "*.*")]
    )
    root.destroy()  # Close tkinter window
    return file_path

# Parameters
output_folder = 'screenshots'
document_name = 'VideoScreenshots.docx'
interval = 5  # Interval in seconds

# Main Process
video_path = select_video_file()  # Drag-and-drop functionality to select video
if video_path:
    clear_folder(output_folder)  # Clear screenshots folder before starting
    extract_screenshots(video_path, interval, output_folder)
    create_document_with_screenshots(output_folder, document_name)
else:
    print("No video file selected. Exiting.")
