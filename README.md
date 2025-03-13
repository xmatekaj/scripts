# Advanced Context Menu Manager - User Manual

## Introduction

The Advanced Context Menu Manager is a comprehensive tool for organizing and managing scripts in your Windows context menu. It allows you to create categories, subcategories, and add scripts that can be executed by right-clicking on files or directories.

## Key Features

- Create hierarchical categories and subcategories in the context menu
- Add Python scripts to be launched from the context menu
- Choose whether scripts appear when right-clicking on files, directories, or both
- Edit script content directly within the application
- Customize script icons for better visual recognition
- Import/export configuration for backup or sharing

## Installation and Setup

1. Download the `advanced_context_menu_manager.py` script
2. **Important**: Run the script as administrator (right-click → Run as administrator)
   - Administrator privileges are required to modify the Windows registry
3. On first run, the application will create its directory structure in `%APPDATA%\AdvancedScriptLauncher`

## Main Interface

The interface is divided into two main sections:

### Left Panel: Category and Script Tree
- Shows all categories, subcategories, and scripts in a hierarchical tree view
- Right-click on items to access context menu options
- Use toolbar buttons to add, edit, or delete items

### Right Panel: Details Area
- Shows details of the currently selected item
- Edit properties of categories or scripts
- For scripts, you can also edit the script content directly

## Managing Categories

### Adding a Category
1. Select a parent item in the tree where you want to add the category
2. Click the "Add Category" button or right-click and select "Add Category"
3. Enter a name for the category
4. Select a parent category (or "Root" for top-level categories)
5. Click "Add Category"

### Editing a Category
1. Select the category in the tree
2. The category details will appear in the right panel
3. Modify the name or parent category
4. Click "Save Changes"

### Deleting a Category
1. Select the category in the tree
2. Click the "Delete" button or right-click and select "Delete"
3. Confirm the deletion
   - **Note**: Deleting a category will also delete all its subcategories and scripts

## Managing Scripts

### Adding a Script
1. Select a category in the tree where you want to add the script
2. Click the "Add Script" button or right-click and select "Add Script"
3. Browse to the Python script file (.py)
4. Enter a name for the script in the context menu
5. Select the category for the script
6. Choose the context(s):
   - Files: Script will appear when right-clicking on files
   - Directories: Script will appear when right-clicking on directories
7. Optionally select an icon file (.ico)
8. Click "Add Script"

### Editing a Script
1. Select the script in the tree
2. The script details will appear in the right panel
3. Modify any properties (name, category, context, icon)
4. Edit the script content directly in the text area
5. Click "Save Changes"

### Deleting a Script
1. Select the script in the tree
2. Click the "Delete" button or right-click and select "Delete"
3. Confirm the deletion

## Using Scripts from the Context Menu

Once you've added scripts, they will appear in your Windows context menu when you right-click on:
- Files (if the script is configured for file context)
- Directories (if the script is configured for directory context)

The scripts will be organized according to your defined categories and subcategories.

## Additional Functions

### Setting Python Path
1. Go to "File" → "Set Python Path"
2. Browse to your Python executable (python.exe)
3. Click "Open"
4. Choose whether to update all registry entries with the new path

### Exporting Configuration
1. Go to "Tools" → "Export Configuration"
2. Choose a location to save the configuration file
3. Click "Save"

### Importing Configuration
1. Go to "Tools" → "Import Configuration"
2. Browse to a previously exported configuration file
3. Click "Open"
4. Confirm the import

### Opening Scripts Directory
1. Go to "Tools" → "Open Scripts Directory"
- This opens the directory where all your scripts are stored

## Script Writing Guidelines

When writing scripts for use with the Advanced Context Menu Manager:

1. Scripts receive the target path as the first command line argument (`sys.argv[1]`)
2. Use a `main()` function as the entry point for your script
3. Check if the target is a file or directory using `os.path.isdir()`
4. Include appropriate error handling and user feedback
5. For a consistent user experience, keep console windows open until the user presses Enter

### Example Script Template:

```python
import os
import sys

def main():
    # Check if we have a target path
    if len(sys.argv) < 2:
        print("No target specified")
        input("Press Enter to exit...")
        return
    
    target_path = sys.argv[1]
    
    # Check if target is a file or directory
    if os.path.isdir(target_path):
        print(f"Target is a directory: {target_path}")
        # Process directory
    else:
        print(f"Target is a file: {target_path}")
        # Process file
    
    # Keep console window open
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
```

## Troubleshooting

### Registry Issues
- If context menu items don't appear, try reopening Explorer or restarting your computer
- Make sure the Python path is correctly set in the application
- Run the application as administrator

### Script Execution Issues
- Check the Python path in the application settings