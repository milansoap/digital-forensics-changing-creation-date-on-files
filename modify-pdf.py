import os
import time
import ctypes
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, TextStringObject
from ctypes import wintypes

def update_pdf_creation_date(file_path, creation_date):
    temp_file = file_path + '.temp'

    # Open the PDF and read metadata
    with open(file_path, 'rb') as pdf_file:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()

        # Copy pages
        for page in reader.pages:
            writer.add_page(page)

        # Copy existing metadata
        metadata = reader.metadata or {}

        # Update only the creation date in metadata
        metadata[NameObject("/CreationDate")] = TextStringObject(creation_date)
        writer.add_metadata(metadata)

        # Write to a temporary file
        with open(temp_file, 'wb') as temp_pdf:
            writer.write(temp_pdf)

    # Preserve original timestamps
    original_access_time = os.path.getatime(file_path)
    original_modification_time = os.path.getmtime(file_path)

    # Replace the original file
    os.replace(temp_file, file_path)

    # Restore timestamps
    os.utime(file_path, (original_access_time, original_modification_time))
    print("Metadata creation date updated.")

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
    file_path = "pdf.pdf"  # Path to the PDF file

    # Set metadata creation date: 2024-09-15 10:23
    creation_date_metadata = "D:20240915102300"
    update_pdf_creation_date(file_path, creation_date_metadata)

    # Set file system creation date: 2024-09-15 10:23
    creation_date_fs = "2024-09-15 10:23:00"
    set_file_creation_date(file_path, creation_date_fs)
