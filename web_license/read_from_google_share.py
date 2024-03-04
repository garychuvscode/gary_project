import requests
import csv


def read_file_from_share_link(share_link):
    try:
        # 從共享連結中提取出檔案 ID
        file_id = share_link.split("/")[5]

        # 使用檔案 ID 組合出下載檔案的連結
        download_link = f"https://drive.google.com/uc?id={file_id}"

        # 發送 HTTP 請求以獲取檔案內容
        response = requests.get(download_link)
        response.raise_for_status()  # 如果請求失敗，則引發異常

        # 返回檔案內容
        return response.text
    except Exception as e:
        print(f"Error reading file from share link: {e}")
        return None


def search_license(file_content, search_item):
    """
    Search for the license value corresponding to the search_item in the 'mail' column.

    Args:
        file_content (str): The content of the CSV file.
        search_item (str): The item to search for in the 'mail' column.

    Returns:
        str: The license value ('Y' or 'N') corresponding to the search_item.
            Returns None if the search_item is not found.
    """
    try:
        # Convert file content string to a CSV reader object
        csv_reader = csv.reader(file_content.split("\n"), delimiter=",")

        # Skip the header row
        next(csv_reader)

        # Iterate through each row in the CSV file
        for row in csv_reader:
            # Check if the search_item is in the 'mail' column
            print(f"now is: {row}")
            print(f"now2 is: {row[2]}")
            print(f"now4 is: {row[4]}")
            if row[2] == search_item:
                # Return the value of the 'license' column for the matching row
                print(f"going to return: {row[4]}")
                return row[4]

        # If the search_item is not found, return None
        print(f"not found")
        return None

    except Exception as e:
        print(f"error: {e}")


# # 測試函數 txt file
# share_link = (
#     "https://drive.google.com/file/d/1E8ppUcJtJMpNwVL0WaEXYQ9iNl3onVWC/view?usp=sharing"
# )

# 測試函數 excel file
share_link = "https://docs.google.com/spreadsheets/d/1GdKn9zvlgpnC1ghrHsUVPrl-DLsmkgpC/edit?usp=sharing&ouid=104829976301101083195&rtpof=true&sd=true"

# 測試函數 excel file (CSV)
share_link = "https://drive.google.com/file/d/1E-PHn9YsPZFqaD_SKbFzNUwvV_21jab7/view?usp=drive_link"


file_content = read_file_from_share_link(share_link)
if file_content is not None:
    print("File content:")
    print(file_content)
else:
    print("Failed to read the file.")


search_license(file_content, "account@example.com")
search_license(file_content, "avon@asus.com.tw")
