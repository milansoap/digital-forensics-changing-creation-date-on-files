import os
import time
from datetime import datetime
import ctypes
from ctypes import wintypes
import zipfile
import xml.etree.ElementTree as ET
import os

def edit_core_metadata(file_path, creation_date):
    # Create a temporary file path for the new ZIP
    temp_file_path = file_path + '.temp'
    
    with zipfile.ZipFile(file_path, 'r') as zf_in, zipfile.ZipFile(temp_file_path, 'w') as zf_out:
        # Iterate over each file in the original ZIP
        for item in zf_in.infolist():
            # Read file content
            content = zf_in.read(item.filename)
            
            # Modify core.xml if it's the target file
            if item.filename == "docProps/core.xml":
                # Parse the XML content
                root = ET.fromstring(content)
                
                # Find the dcterms:created element and update its text
                created_element = root.find('.//{http://purl.org/dc/terms/}created')
                if created_element is not None:
                    created_element.text = creation_date
                else:
                    # Create the element if it doesn't exist
                    dcterms_ns = "{http://purl.org/dc/terms/}"
                    xsi_ns = "{http://www.w3.org/2001/XMLSchema-instance}"
                    created_element = ET.Element(dcterms_ns + "created")
                    created_element.set(xsi_ns + "type", "dcterms:W3CDTF")
                    created_element.text = creation_date
                    root.append(created_element)
                
                # Generate the updated XML content
                content = ET.tostring(root, encoding='utf-8').decode('utf-8')
                # Ensure proper XML declaration
                content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + content
            
            # Write the file to the new ZIP archive
            zf_out.writestr(item, content, zf_in.getinfo(item.filename).compress_type)
    
    # Replace the original ZIP file with the new one
    import os
    os.replace(temp_file_path, file_path)

# Function to set file creation time on Windows
def set_creation_time(file_path, creation_time):
    creation_timestamp = time.mktime(time.strptime(creation_time, "%Y-%m-%d %H:%M:%S"))
    creation_ft = int(creation_timestamp * 10**7 + 116444736000000000)
    creation_time_struct = wintypes.FILETIME(creation_ft & 0xFFFFFFFF, creation_ft >> 32)

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

    ctypes.windll.kernel32.SetFileTime(
        file_handle,
        ctypes.pointer(creation_time_struct),  # Creation time
        None,  # Access time (unchanged)
        None,  # Modification time (unchanged)
    )
    ctypes.windll.kernel32.CloseHandle(file_handle)

# Main script execution
if __name__ == "__main__":
    file_path = "excel2.xlsx"
    creation_date = "2024-09-15T10:23:23Z"
    creation_time_fs = "2024-09-15 10:23:23"

    # Preserve original timestamps
    original_access_time = os.path.getatime(file_path)
    original_modification_time = os.path.getmtime(file_path)

    # Step 1: Edit core.xml metadata in-place
    edit_core_metadata(file_path, creation_date)

    # Step 2: Restore access and modification timestamps
    os.utime(file_path, (original_access_time, original_modification_time))

    # Step 3: Update creation time
    set_creation_time(file_path, creation_time_fs)

    # Step 4: Ensure access and modification times are preserved again
    os.utime(file_path, (original_access_time, original_modification_time))

    print("ZIP Modify Date preserved. Creation date updated successfully!")