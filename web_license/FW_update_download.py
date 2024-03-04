import requests
from bs4 import BeautifulSoup

# forfile system operation
import os

# for delete function
import shutil

# fmt: off

# 全局變數用來存儲文件名和文件ID
file_names = []
file_ids = []
file_list = {}

# 240303 => still need : the create and delay folder for the directly,
# so the update firmware will be able to use

# this is version 1 only have download function
# def download_file_from_drive(file_id, destination_path):
#     """
#     Downloads a file from Google Drive with the given file ID.

#     Args:
#         file_id (str): The ID of the Google Drive file.
#         destination_path (str): The destination path on the local machine to save the file.

#     Returns:
#         None
#     """
#     # Google Drive export link for downloading files
#     export_link = f"https://drive.google.com/uc?id={file_id}"

#     # Send a GET request to the export link
#     response = requests.get(export_link)

#     # Check if the request was successful
#     if response.status_code == 200:
#         # Save the downloaded file to the destination path
#         with open(destination_path, "wb") as f:
#             f.write(response.content)
#         print(f"File downloaded successfully to {destination_path}")
#     else:
#         print("Failed to download file")
def download_file_from_drive(file_id, destination_path):
    """
    Downloads a file from Google Drive with the given file ID.

    Args:
        file_id (str): The ID of the Google Drive file.
        destination_path (str): The destination path on the local machine to save the file.

    Returns:
        None
    """
    # Create the directory if it does not exist
    os.makedirs(os.path.dirname(destination_path), exist_ok=True)

    # Google Drive export link for downloading files
    export_link = f"https://drive.google.com/uc?id={file_id}"

    # Send a GET request to the export link
    response = requests.get(export_link)

    # Check if the request was successful
    if response.status_code == 200:
        # Save the downloaded file to the destination path
        with open(destination_path, "wb") as f:
            f.write(response.content)
        print(f"File downloaded successfully to {destination_path}")
    else:
        print("Failed to download file")

def delete_folder(folder_path):
    """
    Deletes a folder and all its contents.

    Args:
        folder_path (str): The path of the folder to delete.

    Returns:
        None
    """
    shutil.rmtree(folder_path)
    print(f"Folder deleted: {folder_path}")

def parse_drive_link(drive_link):
    """
    Parses a Google Drive share link to extract file name and file ID.

    Args:
        drive_link (str): The Google Drive share link.

    Returns:
        dict: A dictionary mapping file names to file IDs.
    """
    # Extract file ID from the link
    file_id = drive_link.split("/")[5]

    # Get file name using the provided function
    file_name = get_file_name_from_link(drive_link)

    # 將文件名和文件ID作為鍵值對添加到全局字典中
    file_list[file_name] = file_id

    file_names.append(file_name)
    file_ids.append(file_id)

    return file_list


def get_file_name_from_link(drive_link):
    """
    Retrieves the name of a file from a Google Drive share link.

    Args:
        drive_link (str): The Google Drive share link.

    Returns:
        str: The name of the file.
    """
    # Send a GET request to the share link
    response = requests.get(drive_link)

    # Parse the HTML content of the response
    soup = BeautifulSoup(response.content, "html.parser")

    # Find the element containing the file name
    title_element = soup.find("title")

    # Extract and return the file name
    if title_element:
        full_file_name = title_element.text.strip()
        file_name_parts = full_file_name.split(" - ")
        if len(file_name_parts) > 1:
            file_name = file_name_parts[0]
        else:
            file_name = full_file_name
        return file_name
    else:
        return None



# Testing the function
if __name__ == "__main__":

    # testing for get file name
    drive_link = (
        "https://drive.google.com/file/d/1E-PHn9YsPZFqaD_SKbFzNUwvV_21jab7/view"
    )
    file_name0 = get_file_name_from_link(drive_link)
    print("File Name:", file_name0)

    # testing for used previous file name and get file ID
    drive_link = (
        "https://drive.google.com/file/d/1E-PHn9YsPZFqaD_SKbFzNUwvV_21jab7/view"
    )
    file_list0 = parse_drive_link(drive_link)
    file_id0 = file_list0[file_name0]
    print(file_id0)
    print(file_list0)

    print(f"the final name: {file_names}")
    print(f"the final id: {file_ids}")
    print(f"the final name: {file_list}")

    # download file to

    file_id = "1E-PHn9YsPZFqaD_SKbFzNUwvV_21jab7"  # Replace with the ID of the Google Drive file
    destination_path = (
        f"C:/wave_form_raw/{file_name0}"  # Replace with your desired destination path
    )
    download_file_from_drive(file_id, destination_path)
    print(f"down load -{file_name0}- done ")

    # testing version 2 => download, use and delete




    file_id = "1E-PHn9YsPZFqaD_SKbFzNUwvV_21jab7"  # Replace with the ID of the Google Drive file
    destination_path = (
        f"C:/g_tmp_pico/{file_name0}"  # Replace with your desired destination path
    )
    # Download the file
    download_file_from_drive(file_id, destination_path)

    # Use the file...
    input()

    # Delete the folder after using the file
    folder_path = os.path.dirname(destination_path)
    delete_folder(folder_path)
