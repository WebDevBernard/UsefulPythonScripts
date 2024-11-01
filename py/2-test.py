import os
import re

def rename_folders_with_video_names(parent_folder_path):
    # Loop through all items in the parent folder
    for root, dirs, files in os.walk(parent_folder_path):
        for dir_name in dirs:
            folder_path = os.path.join(root, dir_name)

            # Skip specific folders
            if '$RECYCLE.BIN' in folder_path or ".Trash-1000" in folder_path or folder_path == 'D:\\System Volume Information':
                continue

            # List all video files in the current folder
            video_files = [f for f in os.listdir(folder_path) if f.endswith(('.mp4', '.mkv', 'avi'))]

            # Check if any video file matches the S02E01 pattern
            if any(re.search(r'S\d{2}E\d{2}', f) for f in video_files):
                print(f"Ignoring folder '{folder_path}' due to SxxExx pattern.")
                continue  # Skip this folder

            if video_files:
                # Get the first video file
                video_file = video_files[0]

                # Regex to extract the movie name and year, while ignoring additional text
                match = re.search(r'^(.*?)(?:[\[\(]\s*(\d{4})\s*[\]\)])?.*$', video_file)
                if match:
                    video_name = match.group(1).strip().replace('.', ' ')  # Replace dots with space in the title
                    year = match.group(2)  # The four-digit year
                    new_folder_name = f"{video_name.title()} {year}".strip() if year else video_name.title().strip()

                    # Create the new folder path
                    new_folder_path = os.path.join(root, new_folder_name)

                    # Check if the new folder path already exists
                    if os.path.exists(new_folder_path):
                        print(f"Folder '{new_folder_path}' already exists. Skipping rename.")
                    else:
                        # Rename the folder
                        os.rename(folder_path, new_folder_path)
                        print(f"Renamed folder '{folder_path}' to: '{new_folder_name}'")
                else:
                    print(f"No matching pattern found in the video file name: {video_file}.")
            else:
                print(f"No .mp4 or .mkv files found in the folder: {folder_path}.")

# Usage
folder_path = "D:\\"  # Change this to your folder path
rename_folders_with_video_names(folder_path)

print(folder_path)
