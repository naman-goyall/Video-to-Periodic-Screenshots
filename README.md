# Video Screenshot Extractor

This program processes an MP4 video, extracts periodic frames in memory, and generates a Word document with the screenshots embedded—without saving the images to disk.

## Features

1. Extracts frames from a video at regular intervals and embeds them directly into a Word document.
2. Eliminates the need for temporary image storage on disk.
3. Lets you select a video file and output location through a graphical interface.
4. Allows customization of the interval between frames and the size of images in the document.

---

## Prerequisites

### Python Dependencies
Install the required Python libraries using the following command:

```bash
pip install opencv-python pillow python-docx
System-Level Tools
FFmpeg: Required for video processing with OpenCV.
macOS: Install FFmpeg using Homebrew:
bash
Copy code
brew install ffmpeg
Ubuntu/Linux: Install via apt:
bash
Copy code
sudo apt update
sudo apt install ffmpeg
Windows:
Download FFmpeg from FFmpeg's official website.
Add FFmpeg to your system's PATH environment variable.
Usage
Run the script:

bash
Copy code
python script_name.py
A file dialog will appear:

Step 1: Select an MP4 video file.
Step 2: Choose where to save the Word document.
The program will process the video, extract screenshots at the specified interval, and save the Word document.

Locate the generated Word document at the chosen location.

Customization
Interval Between Screenshots
Modify the interval variable in the script (default: 5 seconds).

python
Copy code
interval = 5  # Change to your desired interval
Image Size in Document
Adjust the image size in the Word document by modifying the width in the doc.add_picture call (default: 6 inches).

python
Copy code
doc.add_picture(buffer, width=Inches(6))
Notes
Screenshots are processed in memory, so no temporary files are saved.
Ensure that the program has read/write permissions for the chosen output location.
The program supports .mp4 files by default. Extend support by modifying the filedialog.askopenfilename call in the script.
Example File Structure
bash
Copy code
project/
│
├── script_name.py         # Main Python script
├── requirements.txt       # Python dependencies
├── README.md              # Instructions and setup guide
FAQ
1. I get a "No ffmpeg exe could be found" error. What should I do?
Ensure FFmpeg is installed and added to your system's PATH. Follow the installation steps in the System-Level Tools section.
2. Can I use this with other video formats?
Currently, the program supports .mp4 files. You can extend support by modifying the filedialog.askopenfilename call to include other formats.
License
This project is open-source and available under the MIT License.

yaml
Copy code

---

### What’s New in This Version of the README:
1. Highlighted the **in-memory processing** feature.
2. Removed references to temporary file storage or `screenshots` folders.
3. Updated usage instructions to reflect the new functionality.
4. Clarified customization options for interval and image size.

Save this as `README.md` in your project folder. Let me know if you’d like any further edits!







