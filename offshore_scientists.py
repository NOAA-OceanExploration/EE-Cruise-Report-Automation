import os
import re

# Define the directory path
directory_path = "chats"

# Define unwanted names
unwanted_names = ["iscwatch", "iscwatch2", "admin", "okexnav", "chat-admin"]


# Function to handle edge cases
def format_name(name):
    # Remove spaces before and after the name
    name = name.strip()
    # Remove spaces within the name and convert to lowercase
    return name.replace(" ", "").lower()


# Function to extract researchers from the content
def extract_researchers(content):
    # Regular expression to match researcher names
    pattern = r"\[(\d{2}:\d{2}:\d{2})\] <(.*?)>"
    matches = re.findall(pattern, content)
    researchers = set()
    for _, name in matches:
        if name not in unwanted_names:
            researchers.add(format_name(name))
    return researchers


# Function to process each file
def process_file(file_path):
    with open(file_path, 'r') as file:
        content = file.read()
        # Split the content by "ROV Descending" to get individual dives
        dives = content.split("ROV Descending")
        results = {}
        for dive in dives:
            # Extract the date from the file name
            date = os.path.basename(file_path).split('.')[0]
            researchers = extract_researchers(dive)
            if researchers:
                results[date] = researchers
        return results


# Main function
def main():
    all_results = {}
    for filename in os.listdir(directory_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(directory_path, filename)
            results = process_file(file_path)
            all_results.update(results)

    # Write the results to a text file
    with open("researchers_output.txt", 'w') as output_file:
        for date, researchers in all_results.items():
            output_file.write(f"{date}:\n")
            for researcher in researchers:
                output_file.write(f"{researcher}\n")
            output_file.write("\n")


if __name__ == "__main__":
    main()
