import cv2
import os
# from moviepy.editor import VideoFileClip
from docx import Document
from docx.shared import Inches

def extract_screenshots(video_path, interval, output_folder):
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Load video
    video = cv2.VideoCapture(video_path)
    fps = int(video.get(cv2.CAP_PROP_FPS))
    frame_count = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
    duration = frame_count // fps

    # Calculate frame interval
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
    # Create a Word document
    doc = Document()
    doc.add_heading('Video Screenshots', level=1)

    for file_name in sorted(os.listdir(output_folder)):
        if file_name.endswith('.jpg'):
            file_path = os.path.join(output_folder, file_name)
            doc.add_paragraph(f'Image: {file_name}')
            doc.add_picture(file_path, width=Inches(4))

    doc.save(document_name)
    print(f"Document saved as {document_name}")

# Parameters
video_path = '/Users/namangoyal/Downloads/Spike Prime Soccer Robot With Shooting Arm With Building Instructions.mp4'  # Path to the MP4 file
interval = 5  # Interval in seconds
output_folder = 'screenshots'
document_name = 'VideoScreenshots.docx'

# Process
extract_screenshots(video_path, interval, output_folder)
create_document_with_screenshots(output_folder, document_name)
