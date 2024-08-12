from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# SharePoint site details
site_url = "https://your_domain.sharepoint.com/sites/your_site_name"
username = "your_username@your_domain.com"
password = "your_password"

# Authenticate and create a client context
credentials = UserCredential(username, password)
ctx = ClientContext(site_url).with_credentials(credentials)

# Function to print items in a list
def print_list_items(list_title):
    target_list = ctx.web.lists.get_by_title(list_title)
    items = target_list.items.get().execute_query()
    print(f"\nItems in the '{list_title}' list:")
    for item in items:
        print(f"- {item.properties['Title']}")

# Example 1: Get and print all lists
lists = ctx.web.lists.get().execute_query()
print("Lists in the site:")
for list_obj in lists:
    print(f"- {list_obj.properties['Title']}")

# Example 2: Get items from a specific list
print_list_items("Your List Title")

# Example 3: Add a new item to a list
target_list = ctx.web.lists.get_by_title("Your List Title")
item_properties = {'Title': 'New Item from Python Client'}
target_list.add_item(item_properties).execute_query()
print("\nNew item added to the list.")

# Print the updated list
print_list_items("Your List Title")

# Example 4: Upload a file to a document library
doc_library = ctx.web.lists.get_by_title("Documents")
file_path = "/path/to/your/local/file.txt"
with open(file_path, "rb") as file_content:
    file_name = "uploaded_file.txt"
    target_folder = doc_library.root_folder
    target_file = target_folder.upload_file(file_name, file_content).execute_query()
print(f"\nFile uploaded: {target_file.serverRelativeUrl}")

# Example 5: Download a file
file_url = "/sites/your_site_name/Shared Documents/example.docx"
file = File.open_binary(ctx, file_url)
with open("downloaded_file.docx", "wb") as local_file:
    local_file.write(file.content)
print("\nFile downloaded successfully.")
