import cv2
from docx import Document
from docx.shared import Inches
from tkinter import Tk, filedialog, Button, Label, Entry, StringVar, messagebox, Frame
import numpy as np
from io import BytesIO
from PIL import Image
import os
import tempfile
import threading
import re
import subprocess
import sys

def extract_screenshots_and_create_document(video_path, interval, document_name, start_time=0, end_time=None):
    """Extracts frames from the video and directly inserts them into a Word document."""
    # Initialize Word document
    doc = Document()
    # doc.add_heading('Video Screenshots', level=1)

    # Open the video
    video = cv2.VideoCapture(video_path)
    fps = int(video.get(cv2.CAP_PROP_FPS))
    total_frames = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
    frame_interval = interval * fps
    count = 0
    
    # Calculate start and end frames
    start_frame = int(start_time * fps)
    if end_time is not None:
        end_frame = min(int(end_time * fps), total_frames)
    else:
        end_frame = total_frames
    
    # Set video position to start frame
    video.set(cv2.CAP_PROP_POS_FRAMES, start_frame)
    success, frame = video.read()

    # Process video frames
    while success and video.get(cv2.CAP_PROP_POS_FRAMES) <= end_frame:
        frame_number = int(video.get(cv2.CAP_PROP_POS_FRAMES))
        relative_frame = frame_number - start_frame
        
        if relative_frame % frame_interval == 0:
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
    return count

def validate_youtube_url(url):
    """Validate and normalize YouTube URL."""
    # Standard YouTube watch URL with query parameters
    watch_pattern = r'(?:https?:\/\/)?(?:www\.)?youtube\.com\/watch\?(?:[^&]+&)*v=([0-9A-Za-z_-]{11})(?:&[^&]+)*'
    match = re.search(watch_pattern, url)
    if match:
        video_id = match.group(1)
        return f'https://www.youtube.com/watch?v={video_id}', video_id
    
    # Other common YouTube URL patterns
    other_patterns = [
        r'(?:https?:\/\/)?(?:www\.)?youtu\.be\/([0-9A-Za-z_-]{11})',
        r'(?:https?:\/\/)?(?:www\.)?youtube\.com\/embed\/([0-9A-Za-z_-]{11})',
        r'(?:https?:\/\/)?(?:www\.)?youtube\.com\/v\/([0-9A-Za-z_-]{11})'
    ]
    
    for pattern in other_patterns:
        match = re.search(pattern, url)
        if match:
            video_id = match.group(1)
            return f'https://www.youtube.com/watch?v={video_id}', video_id
    
    # If we get here, no valid YouTube URL pattern was found
    return None, None

def download_youtube_video(youtube_url, temp_dir):
    """Download YouTube video to a temporary file using yt-dlp."""
    try:
        # Validate and normalize the URL
        normalized_url, video_id = validate_youtube_url(youtube_url)
        if not normalized_url:
            return None, "Invalid YouTube URL format. Please use a standard YouTube URL like https://www.youtube.com/watch?v=VIDEO_ID or https://youtu.be/VIDEO_ID"
        
        print(f"Attempting to download video ID: {video_id}")
        
        # Create output filename
        temp_file = os.path.join(temp_dir, f"youtube_video_{video_id}.mp4")
        
        # Build yt-dlp command
        command = [
            sys.executable, '-m', 'yt_dlp',
            '--format', 'mp4',
            '--output', temp_file,
            '--no-playlist',
            '--no-warnings',
            normalized_url
        ]
        
        # Run yt-dlp as a subprocess
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False
        )
        
        # Check if download was successful
        if process.returncode != 0:
            error_msg = process.stderr if process.stderr else "Unknown error occurred"
            print(f"Download error: {error_msg}")
            return None, f"Failed to download video: {error_msg}"
        
        # Check if file exists
        if not os.path.exists(temp_file):
            return None, "Download completed but file was not created"
        
        # Try to get title from output or just use video ID
        title_match = re.search(r'Destination: .*\[info\] (.+?)\.', process.stdout)
        title = title_match.group(1) if title_match else f"YouTube Video {video_id}"
        
        return temp_file, title
    except Exception as e:
        error_message = str(e)
        print(f"YouTube download error: {error_message}")
        return None, f"Failed to download video: {error_message}"

def select_video():
    """Select a video file and process it."""
    video_path = filedialog.askopenfilename(
        title="Select Video File",
        filetypes=[("MP4 Files", "*.mp4"), ("All Files", "*.*")]
    )
    if not video_path:
        messagebox.showwarning("No File", "No video file selected!")
        return

    process_selected_video(video_path)

def process_selected_video(video_path):
    """Process a selected video file."""
    # Ask for output file name and location
    document_path = filedialog.asksaveasfilename(
        title="Save Word Document As",
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")]
    )
    if not document_path:
        messagebox.showwarning("No File", "No save location selected!")
        return

    # Get interval value
    try:
        interval_value = int(interval_var.get())
        if interval_value <= 0:
            raise ValueError("Interval must be positive")
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter a valid positive number for interval")
        return
    
    # Get start time in seconds
    try:
        start_min = int(start_min_var.get()) if start_min_var.get() else 0
        start_sec = int(start_sec_var.get()) if start_sec_var.get() else 0
        if start_min < 0 or start_sec < 0 or start_sec >= 60:
            raise ValueError("Invalid start time")
        start_time = start_min * 60 + start_sec
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter valid numbers for start time")
        return
        
    # Get end time in seconds
    try:
        if end_min_var.get() or end_sec_var.get():
            end_min = int(end_min_var.get()) if end_min_var.get() else 0
            end_sec = int(end_sec_var.get()) if end_sec_var.get() else 0
            if end_min < 0 or end_sec < 0 or end_sec >= 60:
                raise ValueError("Invalid end time")
            end_time = end_min * 60 + end_sec
            if end_time <= start_time:
                raise ValueError("End time must be greater than start time")
        else:
            end_time = None
    except ValueError as e:
        messagebox.showerror("Invalid Input", f"Please enter valid numbers for end time: {str(e)}")
        return

    # Update status
    status_var.set("Processing video... Please wait")
    root.update()

    # Process the video in a separate thread
    def process_thread():
        try:
            frames_count = extract_screenshots_and_create_document(
                video_path, 
                interval_value, 
                document_path,
                start_time,
                end_time
            )
            root.after(0, lambda: messagebox.showinfo("Success", 
                                               f"Processing complete!\n{frames_count} screenshots captured.\nDocument saved at:\n{document_path}"))
            root.after(0, lambda: status_var.set("Ready"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
            root.after(0, lambda: status_var.set("Error occurred"))

    threading.Thread(target=process_thread).start()

def process_youtube_link():
    """Download YouTube video and process it."""
    youtube_url = youtube_url_var.get().strip()
    if not youtube_url:
        messagebox.showwarning("No URL", "Please enter a YouTube URL!")
        return
    
    # Validate URL format first to provide immediate feedback
    normalized_url, _ = validate_youtube_url(youtube_url)
    if not normalized_url:
        messagebox.showerror("Invalid URL", "Please enter a valid YouTube URL\nExample formats:\n- https://www.youtube.com/watch?v=VIDEO_ID\n- https://youtu.be/VIDEO_ID")
        return
    
    # Update status
    status_var.set("Downloading YouTube video... Please wait")
    root.update()
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    
    # Run in a separate thread to keep UI responsive
    def download_and_process():
        try:
            # Download the video
            video_path, video_title = download_youtube_video(youtube_url, temp_dir)
            
            if not video_path:
                root.after(0, lambda: messagebox.showerror("Download Error", f"Failed to download video: {video_title}"))
                root.after(0, lambda: status_var.set("Download failed"))
                return
            
            root.after(0, lambda: status_var.set(f"Downloaded: {video_title}\nSelecting save location..."))
            
            # Ask for output file name and location
            def select_output():
                document_path = filedialog.asksaveasfilename(
                    title="Save Word Document As",
                    defaultextension=".docx",
                    filetypes=[("Word Document", "*.docx")]
                )
                if not document_path:
                    messagebox.showwarning("No File", "No save location selected!")
                    status_var.set("Ready")
                    return
                
                # Get interval value
                try:
                    interval_value = int(interval_var.get())
                    if interval_value <= 0:
                        raise ValueError("Interval must be positive")
                except ValueError:
                    messagebox.showerror("Invalid Input", "Please enter a valid positive number for interval")
                    status_var.set("Ready")
                    return
                
                # Get start time in seconds
                try:
                    start_min = int(start_min_var.get()) if start_min_var.get() else 0
                    start_sec = int(start_sec_var.get()) if start_sec_var.get() else 0
                    if start_min < 0 or start_sec < 0 or start_sec >= 60:
                        raise ValueError("Invalid start time")
                    start_time = start_min * 60 + start_sec
                except ValueError:
                    messagebox.showerror("Invalid Input", "Please enter valid numbers for start time")
                    status_var.set("Ready")
                    return
                    
                # Get end time in seconds
                try:
                    if end_min_var.get() or end_sec_var.get():
                        end_min = int(end_min_var.get()) if end_min_var.get() else 0
                        end_sec = int(end_sec_var.get()) if end_sec_var.get() else 0
                        if end_min < 0 or end_sec < 0 or end_sec >= 60:
                            raise ValueError("Invalid end time")
                        end_time = end_min * 60 + end_sec
                        if end_time <= start_time:
                            raise ValueError("End time must be greater than start time")
                    else:
                        end_time = None
                except ValueError as e:
                    messagebox.showerror("Invalid Input", f"Please enter valid numbers for end time: {str(e)}")
                    status_var.set("Ready")
                    return
                
                # Process the downloaded video
                status_var.set(f"Processing video: {video_title}...")
                
                try:
                    frames_count = extract_screenshots_and_create_document(
                        video_path, 
                        interval_value, 
                        document_path,
                        start_time,
                        end_time
                    )
                    messagebox.showinfo("Success", 
                                       f"Processing complete!\n{frames_count} screenshots captured.\nDocument saved at:\n{document_path}")
                except Exception as e:
                    messagebox.showerror("Processing Error", f"An error occurred while processing: {str(e)}")
                
                # Clean up
                try:
                    os.remove(video_path)
                    os.rmdir(temp_dir)
                except:
                    pass
                
                status_var.set("Ready")
            
            root.after(0, select_output)
            
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
            root.after(0, lambda: status_var.set("Error occurred"))

    threading.Thread(target=download_and_process).start()

# Tkinter Frontend
root = Tk()
root.title("Video Screenshot Extractor")
root.geometry("600x400")  # Slightly wider to accommodate new time fields

# Variables
interval_var = StringVar(value="5")  
start_min_var = StringVar(value="0")
start_sec_var = StringVar(value="0")
end_min_var = StringVar(value="")
end_sec_var = StringVar(value="")
youtube_url_var = StringVar()
status_var = StringVar(value="Ready")

# Main frame
main_frame = Frame(root, padx=20, pady=20)
main_frame.pack(fill="both", expand=True)

# Title
Label(main_frame, text="Video Screenshot Extractor", font=("Arial", 16, "bold")).pack(pady=10)

# Interval setting
interval_frame = Frame(main_frame)
interval_frame.pack(fill="x", pady=5)
Label(interval_frame, text="Interval between screenshots (seconds):").pack(side="left")
Entry(interval_frame, textvariable=interval_var, width=5).pack(side="left", padx=5)

# Trim settings
trim_frame = Frame(main_frame)
trim_frame.pack(fill="x", pady=10)

# Start time
start_time_frame = Frame(trim_frame)
start_time_frame.pack(side="left", padx=(0, 20))
Label(start_time_frame, text="Start time:").pack(side="left")

start_min_frame = Frame(start_time_frame)
start_min_frame.pack(side="left", padx=5)
Entry(start_min_frame, textvariable=start_min_var, width=3).pack(side="left")
Label(start_min_frame, text="min").pack(side="left")

start_sec_frame = Frame(start_time_frame)
start_sec_frame.pack(side="left")
Entry(start_sec_frame, textvariable=start_sec_var, width=3).pack(side="left")
Label(start_sec_frame, text="sec").pack(side="left")

# End time
end_time_frame = Frame(trim_frame)
end_time_frame.pack(side="left")
Label(end_time_frame, text="End time (optional):").pack(side="left")

end_min_frame = Frame(end_time_frame)
end_min_frame.pack(side="left", padx=5)
Entry(end_min_frame, textvariable=end_min_var, width=3).pack(side="left")
Label(end_min_frame, text="min").pack(side="left")

end_sec_frame = Frame(end_time_frame)
end_sec_frame.pack(side="left")
Entry(end_sec_frame, textvariable=end_sec_var, width=3).pack(side="left")
Label(end_sec_frame, text="sec").pack(side="left")

# Video selection
video_frame = Frame(main_frame)
video_frame.pack(fill="x", pady=10)
Button(video_frame, text="Select Video File", command=select_video, width=20).pack(pady=5)

# YouTube URL
youtube_frame = Frame(main_frame)
youtube_frame.pack(fill="x", pady=5)
Label(youtube_frame, text="YouTube URL:").pack(anchor="w")
Entry(youtube_frame, textvariable=youtube_url_var, width=60).pack(fill="x", pady=5)
Button(youtube_frame, text="Process YouTube Link", command=process_youtube_link, width=20).pack(pady=5)

# Status bar
status_frame = Frame(main_frame)
status_frame.pack(fill="x", side="bottom", pady=10)
Label(status_frame, textvariable=status_var, bd=1, relief="sunken", anchor="w").pack(fill="x")

root.mainloop()