# EE-Cruise-Report-Automation

This guide provides instructions to add shore scientists to the cruise report table using the provided Python and Google Appscript code.

## Instructions

### 1. Download the Chat Data
   - Visit `chatroom@rooms.exdata.tgfoe.org` to download the GFOE Chat data.

### 2. Prepare Your Files
   - Move the chat data into a new folder named `chats`.
   - Ensure this folder is in the same location as the `offshore_scientists.py` script.

### 3. Run the Script
   - Open your terminal or command prompt.
   - Navigate to where `offshore_scientists.py` is located.
   - Type and enter `python offshore_scientists.py`.

### 4. Upload the Result to Google Drive
   - Find the `researchers_output.txt` file in your working directory.
   - Upload this file to your Google Drive.
   - Extract the file ID from the shareable link.

### 5. Update the Google Appscript
   - Open your cruise report's Google Appscript.
   - Replace the existing ID in line 22 with the one you extracted.
   - Save the Appscript.

### 6. Generate the Table
   - In the Google Appscript, find the `Generation Menu`.
   - Select `Add Shore Scientists to Table`.
   - The table in your cruise report will auto-populate based on the specified dates.
