import win32com.client
def get_folder(path):
    """
    path: list of folder names from root to target
    Example: ["Mailbox - John Doe", "Inbox", "Focused"]
    """
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = ns.Folders[path[0]]
    for name in path[1:]:
        folder = folder.Folders[name]
    return folder

# Example usage:
focused_folder = get_folder(["Mailbox - John Doe", "Inbox", "Focused"])
for msg in focused_folder.Items:
    print(msg.Subject)
