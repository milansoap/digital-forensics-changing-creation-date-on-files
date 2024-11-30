import os
import time
import ctypes
from PIL import Image
from ctypes import wintypes

def update_jpg_metadata_creation_date(file_path, creation_date):
    temp_file = file_path + ".temp"

    # Open the image
    image = Image.open(file_path)
    exif_data = image.info.get("exif")

    # Skip EXIF update if no EXIF data is present
    if not exif_data:
        print("No EXIF metadata found. Skipping metadata update.")
        return

    # Load existing EXIF data and update DateTimeOriginal
    exif_dict = image.getexif()
    exif_dict[36867] = creation_date  # 36867 is the EXIF tag for DateTimeOriginal

    # Save the updated image
    image.save(temp_file, exif=exif_dict.tobytes())

    # Replace the original file
    os.replace(temp_file, file_path)
    print("Metadata creation date updated in EXIF.")

def set_file_creation_date(file_path, creation_time):
    # Convert creation time to Windows FILETIME format
    creation_timestamp = time.mktime(time.strptime(creation_time, "%Y-%m-%d %H:%M:%S"))
    creation_ft = int(creation_timestamp * 10**7 + 116444736000000000)
    creation_time_struct = wintypes.FILETIME(creation_ft & 0xFFFFFFFF, creation_ft >> 32)

    # Open the file handle
    file_handle = ctypes.windll.kernel32.CreateFileW(
        file_path,
        0x40000000,  # GENERIC_WRITE
        0,
        None,
        3,  # OPEN_EXISTING
        0x02000000,  # FILE_FLAG_BACKUP_SEMANTICS
        None,
    )
    if file_handle == -1:
        raise OSError("Failed to open file handle.")

    # Set the file creation time
    success = ctypes.windll.kernel32.SetFileTime(
        file_handle,
        ctypes.pointer(creation_time_struct),  # Creation time
        None,  # Access time (unchanged)
        None,  # Modification time (unchanged)
    )
    ctypes.windll.kernel32.CloseHandle(file_handle)

    if not success:
        raise OSError("Failed to update file creation time.")

    print("File creation date updated.")

if __name__ == "__main__":
    file_path = "jpg.jpg"  # Path to the JPG file

    # Set metadata creation date: 2024-09-15 10:23:00
    creation_date_metadata = "2024:09:15 10:23:00"
    update_jpg_metadata_creation_date(file_path, creation_date_metadata)

    # Set file system creation date: 2024-09-15 10:23:00
    creation_date_fs = "2024-09-15 10:23:00"
    set_file_creation_date(file_path, creation_date_fs)
