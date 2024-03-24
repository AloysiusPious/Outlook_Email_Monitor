import os

# Predefined list of directories to delete
directories_to_delete = ['Charts', 'Reports','Critical_Alerts','tmp','Incidents']

# Delete directories and their contents
for directory in directories_to_delete:
    try:
        # Remove files in the directory
        for root, dirs, files in os.walk(directory):
            for file in files:
                os.remove(os.path.join(root, file))
        # Remove the directory itself
        os.rmdir(directory)
        print(f'Deleted directory and its contents: {directory}')
    except FileNotFoundError:
        print(f'Directory not found: {directory}')
    except PermissionError:
        print(f'Permission denied for directory: {directory}')
    except Exception as e:
        print(f'Error deleting directory: {directory}, {e}')