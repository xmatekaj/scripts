import os
import sys
import winreg
import subprocess
import shutil
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import uuid

class ContextMenuRegistry:
    """Class for managing Windows registry entries for context menu items"""
    
    @staticmethod
    def add_menu_item(context_type, menu_path, command, icon=None):
        """
        Add a menu item to the Windows context menu
        
        Args:
            context_type (str): Type of context ('directory', 'file', '*', etc.)
            menu_path (str): Path in context menu (e.g., 'ScriptTools\\PDF Tools')
            command (str): Command to execute
            icon (str, optional): Path to icon file
            
        Returns:
            bool: Success or failure
        """
        try:
            # Determine the registry key based on context type
            if context_type == 'directory':
                base_key = r'Directory\shell'
            elif context_type == 'file':
                base_key = r'*\shell'
            else:
                base_key = rf'{context_type}\shell'
                
            # Create registry keys for path components
            key_path = base_key
            path_components = menu_path.split('\\')
            
            # Process all components except the last (which is the command name)
            for i, component in enumerate(path_components[:-1]):
                # Create a key for this submenu and set MUIVerb for its display name
                submenu_key_path = f"{key_path}\\{component}"
                submenu_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, submenu_key_path)
                winreg.SetValueEx(submenu_key, "MUIVerb", 0, winreg.REG_SZ, component)
                winreg.SetValueEx(submenu_key, "subcommands", 0, winreg.REG_SZ, "")
                
                # If this is the first component, set the shell key
                if i == 0:
                    shell_key_path = f"{submenu_key_path}\\shell"
                    shell_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shell_key_path)
                    winreg.CloseKey(shell_key)
                
                winreg.CloseKey(submenu_key)
                
                # Update key_path to include this component and its shell subkey
                key_path = f"{submenu_key_path}\\shell"
            
            # Create the final command key with the last path component
            command_key_path = f"{key_path}\\{path_components[-1]}"
            command_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key_path)
            winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, path_components[-1])
            
            # Set icon if provided
            if icon:
                winreg.SetValueEx(command_key, "Icon", 0, winreg.REG_SZ, icon)
            
            # Create command subkey
            cmd_subkey = winreg.CreateKey(command_key, "command")
            winreg.SetValueEx(cmd_subkey, "", 0, winreg.REG_SZ, command)
            
            winreg.CloseKey(cmd_subkey)
            winreg.CloseKey(command_key)
            
            return True
        except Exception as e:
            print(f"Error adding menu item: {e}")
            return False
    
    @staticmethod
    def remove_menu_item(context_type, menu_path):
        """
        Remove a menu item from the Windows context menu
        
        Args:
            context_type (str): Type of context ('directory', 'file', '*', etc.)
            menu_path (str): Path in context menu (e.g., 'ScriptTools\\PDF Tools')
            
        Returns:
            bool: Success or failure
        """
        try:
            # Determine the registry key based on context type
            if context_type == 'directory':
                base_key = r'Directory\shell'
            elif context_type == 'file':
                base_key = r'*\shell'
            else:
                base_key = rf'{context_type}\shell'
            
            # Split the menu path into components
            path_components = menu_path.split('\\')
            
            # Build the full path to the command key
            key_path = base_key
            for i, component in enumerate(path_components[:-1]):
                key_path = f"{key_path}\\{component}"
                if i == 0:
                    key_path = f"{key_path}\\shell"
            
            # Delete the command key
            command_key_path = f"{key_path}\\{path_components[-1]}"
            
            # Delete command subkey first
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"{command_key_path}\\command")
            except Exception:
                pass
            
            # Then delete the command key itself
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, command_key_path)
            except Exception:
                pass
            
            # Cleanup: if this was the last item in a category, remove the category too
            try:
                # Check if the shell key is empty
                shell_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
                try:
                    winreg.EnumKey(shell_key, 0)
                    has_subkeys = True
                except Exception:
                    has_subkeys = False
                winreg.CloseKey(shell_key)
                
                if not has_subkeys:
                    # Remove empty parent keys (working backwards from bottom to top)
                    for i in range(len(path_components) - 1, 0, -1):
                        parent_path = base_key
                        for j in range(i):
                            parent_path = f"{parent_path}\\{path_components[j]}"
                            if j == 0:
                                parent_path = f"{parent_path}\\shell"
                        
                        try:
                            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, parent_path)
                        except Exception:
                            # If we can't delete it (not empty), stop going up
                            break
            except Exception:
                pass
            
            return True
        except Exception as e:
            print(f"Error removing menu item: {e}")
            return False

class ScriptManager:
    """Class for managing script files and configuration"""
    
    def __init__(self):
        # Define the directory where scripts will be stored
        self.app_dir = os.path.join(os.environ['APPDATA'], 'AdvancedScriptLauncher')
        self.scripts_dir = os.path.join(self.app_dir, 'scripts')
        self.config_path = os.path.join(self.app_dir, 'config.json')
        self.launcher_path = os.path.join(self.app_dir, 'launcher.py')
        
        # Create necessary directories if they don't exist
        os.makedirs(self.scripts_dir, exist_ok=True)
        
        # Create launcher script if it doesn't exist
        if not os.path.exists(self.launcher_path):
            self._create_launcher_script()
        
        # Load configuration
        self.config = self.load_config()
    
    def load_config(self):
        """Load the configuration file or create a new one if it doesn't exist."""
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r') as f:
                return json.load(f)
        else:
            default_config = {
                'categories': [],
                'scripts': [],
                'python_path': sys.executable
            }
            with open(self.config_path, 'w') as f:
                json.dump(default_config, f, indent=4)
            return default_config
    
    def save_config(self):
        """Save the current configuration to file."""
        with open(self.config_path, 'w') as f:
            json.dump(self.config, f, indent=4)
    
    def add_category(self, name, parent=None):
        """
        Add a new category.
        
        Args:
            name (str): Category name
            parent (str, optional): Parent category ID
            
        Returns:
            str: Category ID
        """
        category_id = str(uuid.uuid4())
        category = {
            'id': category_id,
            'name': name,
            'parent': parent
        }
        self.config['categories'].append(category)
        self.save_config()
        return category_id
    
    def remove_category(self, category_id):
        """
        Remove a category and all its scripts.
        
        Args:
            category_id (str): Category ID
            
        Returns:
            bool: Success or failure
        """
        # Find all scripts in this category and remove them
        scripts_to_remove = [s for s in self.config['scripts'] if s['category'] == category_id]
        for script in scripts_to_remove:
            self.remove_script(script['id'])
        
        # Remove subcategories recursively
        subcategories = [c for c in self.config['categories'] if c['parent'] == category_id]
        for subcat in subcategories:
            self.remove_category(subcat['id'])
        
        # Remove the category itself
        self.config['categories'] = [c for c in self.config['categories'] if c['id'] != category_id]
        self.save_config()
        return True
    
    def get_category_path(self, category_id):
        """
        Get the full path for a category (for context menu).
        
        Args:
            category_id (str): Category ID
            
        Returns:
            str: Full category path (e.g., 'ScriptTools\\PDF Tools')
        """
        if not category_id:
            return "ScriptTools"
        
        # Find this category
        category = next((c for c in self.config['categories'] if c['id'] == category_id), None)
        if not category:
            return "ScriptTools"
        
        # If has parent, get parent path recursively
        if category['parent']:
            parent_path = self.get_category_path(category['parent'])
            return f"{parent_path}\\{category['name']}"
        else:
            return f"ScriptTools\\{category['name']}"
    
    def add_script(self, script_path, name, category_id, contexts, icon_path=None):
        """
        Add a new script.
        
        Args:
            script_path (str): Path to the script file
            name (str): Script name
            category_id (str): Category ID
            contexts (list): List of contexts ('file', 'directory')
            icon_path (str, optional): Path to an icon file
            
        Returns:
            str: Script ID
        """
        # Generate a UUID for the script
        script_id = str(uuid.uuid4())
        
        # Copy script to scripts directory with UUID filename to avoid collisions
        file_ext = os.path.splitext(script_path)[1]
        dest_filename = f"{script_id}{file_ext}"
        dest_path = os.path.join(self.scripts_dir, dest_filename)
        shutil.copy2(script_path, dest_path)
        
        # Copy icon if provided
        icon_dest = None
        if icon_path:
            icon_ext = os.path.splitext(icon_path)[1]
            icon_filename = f"{script_id}_icon{icon_ext}"
            icon_dest = os.path.join(self.scripts_dir, icon_filename)
            shutil.copy2(icon_path, icon_dest)
        
        # Add to configuration
        script_info = {
            'id': script_id,
            'name': name,
            'category': category_id,
            'filename': dest_filename,
            'contexts': contexts,
            'icon': icon_filename if icon_path else None
        }
        self.config['scripts'].append(script_info)
        self.save_config()
        
        # Create registry entries
        self._create_registry_entries(script_info)
        
        return script_id
    
    def update_script(self, script_id, name=None, category_id=None, contexts=None, icon_path=None):
        """
        Update a script.
        
        Args:
            script_id (str): Script ID
            name (str, optional): New script name
            category_id (str, optional): New category ID
            contexts (list, optional): New list of contexts
            icon_path (str, optional): New icon path
            
        Returns:
            bool: Success or failure
        """
        # Find the script
        script = next((s for s in self.config['scripts'] if s['id'] == script_id), None)
        if not script:
            return False
        
        # Remove old registry entries
        self._remove_registry_entries(script)
        
        # Update script properties
        if name:
            script['name'] = name
        if category_id:
            script['category'] = category_id
        if contexts:
            script['contexts'] = contexts
        
        # Update icon if provided
        if icon_path:
            # Remove old icon if exists
            if script['icon']:
                old_icon_path = os.path.join(self.scripts_dir, script['icon'])
                if os.path.exists(old_icon_path):
                    os.remove(old_icon_path)
            
            # Copy new icon
            icon_ext = os.path.splitext(icon_path)[1]
            icon_filename = f"{script_id}_icon{icon_ext}"
            icon_dest = os.path.join(self.scripts_dir, icon_filename)
            shutil.copy2(icon_path, icon_dest)
            script['icon'] = icon_filename
        
        # Save configuration
        self.save_config()
        
        # Create new registry entries
        self._create_registry_entries(script)
        
        return True
    
    def remove_script(self, script_id):
        """
        Remove a script.
        
        Args:
            script_id (str): Script ID
            
        Returns:
            bool: Success or failure
        """
        # Find the script
        script = next((s for s in self.config['scripts'] if s['id'] == script_id), None)
        if not script:
            return False
        
        # Remove registry entries
        self._remove_registry_entries(script)
        
        # Remove script file
        script_path = os.path.join(self.scripts_dir, script['filename'])
        if os.path.exists(script_path):
            os.remove(script_path)
        
        # Remove icon if exists
        if script['icon']:
            icon_path = os.path.join(self.scripts_dir, script['icon'])
            if os.path.exists(icon_path):
                os.remove(icon_path)
        
        # Remove from configuration
        self.config['scripts'] = [s for s in self.config['scripts'] if s['id'] != script_id]
        self.save_config()
        
        return True
    
    def _create_registry_entries(self, script):
        """Create the registry entries for a script."""
        # Get the script path
        script_path = os.path.join(self.scripts_dir, script['filename'])
        
        # Get category path
        category_path = self.get_category_path(script['category'])
        menu_path = f"{category_path}\\{script['name']}"
        
        # Get icon path
        icon_path = None
        if script['icon']:
            icon_path = os.path.join(self.scripts_dir, script['icon'])
        
        # Create command string
        launcher_cmd = f'"{self.config["python_path"]}" "{self.launcher_path}" "{script_path}" "%1"'
        
        # Add registry entries for each context
        for context in script['contexts']:
            ContextMenuRegistry.add_menu_item(context, menu_path, launcher_cmd, icon_path)
    
    def _remove_registry_entries(self, script):
        """Remove the registry entries for a script."""
        # Get category path
        category_path = self.get_category_path(script['category'])
        menu_path = f"{category_path}\\{script['name']}"
        
        # Remove registry entries for each context
        for context in script['contexts']:
            ContextMenuRegistry.remove_menu_item(context, menu_path)
    
    def _create_launcher_script(self):
        """Create the launcher script that will execute the selected Python script."""
        with open(self.launcher_path, 'w') as f:
            f.write('''
import sys
import os
import importlib.util
import subprocess

def run_script(script_path, target_path):
    """Run a Python script with the target path as argument."""
    # If it's a .py file, import and run it
    if script_path.endswith('.py'):
        # Get the directory of the script
        script_dir = os.path.dirname(script_path)
        
        # Add the script directory to sys.path
        if script_dir not in sys.path:
            sys.path.insert(0, script_dir)
        
        # Import the script
        module_name = os.path.basename(script_path).replace('.py', '')
        spec = importlib.util.spec_from_file_location(module_name, script_path)
        module = importlib.util.module_from_spec(spec)
        
        # Set sys.argv for the script
        original_argv = sys.argv.copy()
        sys.argv = [script_path, target_path]
        
        # Execute the script
        try:
            spec.loader.exec_module(module)
            # If the script has a main function, call it
            if hasattr(module, 'main'):
                module.main()
        except Exception as e:
            print(f"Error executing script: {e}")
            input("Press Enter to continue...")
        finally:
            # Restore sys.argv
            sys.argv = original_argv
    else:
        # For non-Python scripts, run as a subprocess
        subprocess.run([sys.executable, script_path, target_path], check=True)

if __name__ == "__main__":
    # Check command line arguments
    if len(sys.argv) >= 3:
        script_path = sys.argv[1]
        target_path = sys.argv[2]
        
        # Run the script
        run_script(script_path, target_path)
    else:
        print("Usage: launcher.py script_path target_path")
        input("Press Enter to continue...")
''')

class AdvancedContextMenuGUI(tk.Tk):
    """GUI for managing context menu scripts with categories and subcategories"""
    
    def __init__(self):
        super().__init__()
        self.title("Advanced Context Menu Manager")
        self.geometry("900x700")
        self.minsize(800, 600)
        
        # Initialize script manager
        self.script_manager = ScriptManager()
        
        # Set up UI
        self._setup_ui()
        
        # Load data
        self._load_data()
    
    def _setup_ui(self):
        """Set up the user interface"""
        # Create menu bar
        self._setup_menu()
        
        # Create main panes
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left panel - Categories and scripts tree
        self.left_frame = ttk.Frame(self.paned)
        self.paned.add(self.left_frame, weight=1)
        
        # Tree view for categories and scripts
        self.tree_frame = ttk.Frame(self.left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tree with scrollbar
        self.tree_scroll = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree_scroll.config(command=self.tree.yview)
        
        # Configure tree
        self.tree["columns"] = ("type",)
        self.tree.column("#0", width=250, minwidth=150)
        self.tree.column("type", width=100, anchor=tk.CENTER)
        self.tree.heading("#0", text="Name")
        self.tree.heading("type", text="Type")
        
        # Add root node
        self.tree.insert("", "end", "root", text="ScriptTools", values=("Main",))
        
        # Tree right-click menu
        self.tree_menu = tk.Menu(self.tree, tearoff=0)
        self.tree_menu.add_command(label="Add Category", command=self.add_category)
        self.tree_menu.add_command(label="Add Script", command=self.add_script)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Edit", command=self.edit_item)
        self.tree_menu.add_command(label="Delete", command=self.delete_item)
        
        # Bind tree events
        self.tree.bind("<Button-3>", self._show_tree_menu)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        
        # Tree toolbar
        self.tree_toolbar = ttk.Frame(self.left_frame)
        self.tree_toolbar.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(self.tree_toolbar, text="Add Category", command=self.add_category).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.tree_toolbar, text="Add Script", command=self.add_script).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.tree_toolbar, text="Edit", command=self.edit_item).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.tree_toolbar, text="Delete", command=self.delete_item).pack(side=tk.LEFT, padx=2)
        
        # Right panel - Details
        self.right_frame = ttk.Frame(self.paned)
        self.paned.add(self.right_frame, weight=2)
        
        # Details area
        self.details_frame = ttk.LabelFrame(self.right_frame, text="Details")
        self.details_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # No selection label
        self.no_selection_label = ttk.Label(self.details_frame, text="Select an item to view details")
        self.no_selection_label.pack(pady=50)
        
        # Category details frame (initially hidden)
        self.category_frame = ttk.Frame(self.details_frame)
        
        ttk.Label(self.category_frame, text="Category Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_name_var = tk.StringVar()
        self.category_name_entry = ttk.Entry(self.category_frame, textvariable=self.category_name_var, width=40)
        self.category_name_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.category_frame, text="Parent Category:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_parent_var = tk.StringVar()
        self.category_parent_combo = ttk.Combobox(self.category_frame, textvariable=self.category_parent_var, width=38)
        self.category_parent_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Button(self.category_frame, text="Save Changes", command=self.save_category).grid(row=2, column=0, columnspan=2, pady=10)
        
        # Script details frame (initially hidden)
        self.script_frame = ttk.Frame(self.details_frame)
        
        ttk.Label(self.script_frame, text="Script Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.script_name_var = tk.StringVar()
        self.script_name_entry = ttk.Entry(self.script_frame, textvariable=self.script_name_var, width=40)
        self.script_name_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.script_frame, text="Category:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.script_category_var = tk.StringVar()
        self.script_category_combo = ttk.Combobox(self.script_frame, textvariable=self.script_category_var, width=38)
        self.script_category_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.script_frame, text="Context:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        context_frame = ttk.Frame(self.script_frame)
        context_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.file_context_var = tk.BooleanVar(value=True)
        self.dir_context_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(context_frame, text="Files", variable=self.file_context_var).pack(side=tk.LEFT)
        ttk.Checkbutton(context_frame, text="Directories", variable=self.dir_context_var).pack(side=tk.LEFT)
        
        ttk.Label(self.script_frame, text="Icon:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        icon_frame = ttk.Frame(self.script_frame)
        icon_frame.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.script_icon_var = tk.StringVar()
        self.script_icon_entry = ttk.Entry(icon_frame, textvariable=self.script_icon_var, width=30)
        self.script_icon_entry.pack(side=tk.LEFT)
        ttk.Button(icon_frame, text="Browse...", command=self._browse_icon).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.script_frame, text="Script Content:").grid(row=4, column=0, sticky=tk.NW, padx=5, pady=5)
        self.script_content = ScrolledText(self.script_frame, width=40, height=15)
        self.script_content.grid(row=4, column=1, sticky=tk.NSEW, padx=5, pady=5)
        
        ttk.Button(self.script_frame, text="Save Changes", command=self.save_script).grid(row=5, column=0, columnspan=2, pady=10)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Set initial status
        self.status_var.set("Ready")
    
    def _setup_menu(self):
        """Set up the menu bar"""
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        
        # File menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        
        self.file_menu.add_command(label="Set Python Path", command=self.set_python_path)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit)
        
        # Tools menu
        self.tools_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Tools", menu=self.tools_menu)
        
        self.tools_menu.add_command(label="Open Scripts Directory", command=self.open_scripts_dir)
        self.tools_menu.add_command(label="Export Configuration", command=self.export_config)
        self.tools_menu.add_command(label="Import Configuration", command=self.import_config)
        
        # Help menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        
        self.help_menu.add_command(label="About", command=self.show_about)
    
    def _load_data(self):
        """Load categories and scripts from configuration"""
        # Clear tree
        for item in self.tree.get_children("root"):
            self.tree.delete(item)
        
        # Dictionary to map category IDs to tree IDs
        self.category_map = {"": "root"}
        
        # Add categories
        for category in self.script_manager.config['categories']:
            parent_id = self.category_map.get(category['parent'], "root")
            tree_id = self.tree.insert(parent_id, "end", text=category['name'], values=("Category",))
            self.category_map[category['id']] = tree_id
        
        # Add scripts
        for script in self.script_manager.config['scripts']:
            parent_id = self.category_map.get(script['category'], "root")
            script_contexts = ", ".join(script['contexts'])
            self.tree.insert(parent_id, "end", text=script['name'], values=(f"Script ({script_contexts})",), tags=(script['id'],))
    
    def _show_tree_menu(self, event):
        """Show context menu for tree view"""
        # Select item under cursor
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree_menu.post(event.x_root, event.y_root)
    
    def _on_tree_select(self, event):
        """Handle selection in tree view"""
        selection = self.tree.selection()
        if not selection:
            # No selection, show default
            self.no_selection_label.pack(pady=50)
            self.category_frame.pack_forget()
            self.script_frame.pack_forget()
            return
        
        item_id = selection[0]
        item_type = self.tree.item(item_id, "values")[0]
        
        # Hide all detail frames
        self.no_selection_label.pack_forget()
        self.category_frame.pack_forget()
        self.script_frame.pack_forget()
        
        if "Category" in item_type:
            # Show category details
            self._show_category_details(item_id)
        elif "Script" in item_type:
            # Show script details
            self._show_script_details(item_id)
    
    def _show_category_details(self, item_id):
        """Show details for selected category"""
        # Get category name
        category_name = self.tree.item(item_id, "text")
        
        # Find category in configuration
        category = None
        category_id = ""
        for cat in self.script_manager.config['categories']:
            if cat['name'] == category_name and self.category_map.get(cat['id']) == item_id:
                category = cat
                category_id = cat['id']
                break
        
        if not category and item_id == "root":
            # Root category
            category = {"name": "ScriptTools", "parent": ""}
        
        if category:
            # Show category frame
            self.category_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Update fields
            self.category_name_var.set(category['name'])
            
            # Update parent category dropdown
            self.category_parent_combo['values'] = ["None"]
            self.category_parent_map = {"None": ""}
            
            for cat in self.script_manager.config['categories']:
                # Skip itself and its children (to avoid circular references)
                if cat['id'] == category_id:
                    continue
                
                # Check if this category is a child of the selected category
                parent_id = cat['parent']
                is_child = False
                while parent_id:
                    if parent_id == category_id:
                        is_child = True
                        break
                    parent_cat = next((c for c in self.script_manager.config['categories'] if c['id'] == parent_id), None)
                    if not parent_cat:
                        break
                    parent_id = parent_cat['parent']
                
                if not is_child:
                    self.category_parent_combo['values'] = list(self.category_parent_combo['values']) + [cat['name']]
                    self.category_parent_map[cat['name']] = cat['id']
            
            # Set current parent
            if category['parent']:
                parent_cat = next((c for c in self.script_manager.config['categories'] if c['id'] == category['parent']), None)
                if parent_cat:
                    self.category_parent_var.set(parent_cat['name'])
                else:
                    self.category_parent_var.set("None")
            else:
                self.category_parent_var.set("None")
            
            # Store category ID for save function
            self.category_frame.category_id = category_id
        else:
            # Category not found (shouldn't happen)
            self.no_selection_label.pack(pady=50)
    
    def _show_script_details(self, item_id):
        """Show details for selected script"""
        # Get script name and ID
        script_name = self.tree.item(item_id, "text")
        script_id = None
        
        # Find the script tag (which contains the ID)
        tags = self.tree.item(item_id, "tags")
        if tags:
            script_id = tags[0]
        
        # Find script in configuration
        script = next((s for s in self.script_manager.config['scripts'] if s['id'] == script_id), None)
        
        if script:
            # Show script frame
            self.script_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Update fields
            self.script_name_var.set(script['name'])
            
            # Update category dropdown
            self.script_category_combo['values'] = ["Root"]
            self.script_category_map = {"Root": ""}
            
            for cat in self.script_manager.config['categories']:
                self.script_category_combo['values'] = list(self.script_category_combo['values']) + [cat['name']]
                self.script_category_map[cat['name']] = cat['id']
            
            # Set current category
            if script['category']:
                category = next((c for c in self.script_manager.config['categories'] if c['id'] == script['category']), None)
                if category:
                    self.script_category_var.set(category['name'])
                else:
                    self.script_category_var.set("Root")
            else:
                self.script_category_var.set("Root")
            
            # Set contexts
            self.file_context_var.set('file' in script['contexts'])
            self.dir_context_var.set('directory' in script['contexts'])
            
            # Set icon
            if script['icon']:
                icon_path = os.path.join(self.script_manager.scripts_dir, script['icon'])
                self.script_icon_var.set(icon_path)
            else:
                self.script_icon_var.set("")
            
            # Load script content
            script_path = os.path.join(self.script_manager.scripts_dir, script['filename'])
            try:
                with open(script_path, 'r') as f:
                    content = f.read()
                self.script_content.delete('1.0', tk.END)
                self.script_content.insert('1.0', content)
            except Exception as e:
                self.script_content.delete('1.0', tk.END)
                self.script_content.insert('1.0', f"Error loading script: {e}")
            
            # Store script ID for save function
            self.script_frame.script_id = script_id
        else:
            # Script not found
            self.no_selection_label.pack(pady=50)
    
    def _browse_icon(self):
        """Browse for an icon file"""
        icon_path = filedialog.askopenfilename(
            title="Select Icon",
            filetypes=[("Icon Files", "*.ico"), ("All Files", "*.*")]
        )
        if icon_path:
            self.script_icon_var.set(icon_path)
    
    def add_category(self):
        """Add a new category"""
        # Create dialog window
        dialog = tk.Toplevel(self)
        dialog.title("Add Category")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        # Dialog content
        ttk.Label(dialog, text="Category Name:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Parent Category:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        parent_var = tk.StringVar(value="Root")
        parent_combo = ttk.Combobox(dialog, textvariable=parent_var, width=28)
        parent_combo.grid(row=1, column=1, padx=10, pady=10)
        
        # Populate parent dropdown
        parent_combo['values'] = ["Root"]
        parent_map = {"Root": ""}
        
        for cat in self.script_manager.config['categories']:
            parent_combo['values'] = list(parent_combo['values']) + [cat['name']]
            parent_map[cat['name']] = cat['id']
        
        # Add button
        ttk.Button(dialog, text="Add Category", command=lambda: self._add_category_action(
            name_var.get(), parent_map[parent_var.get()], dialog
        )).grid(row=2, column=0, columnspan=2, pady=10)
    
    def _add_category_action(self, name, parent, dialog):
        """Process adding a category"""
        if not name:
            messagebox.showerror("Error", "Category name cannot be empty")
            return
        
        # Add category
        category_id = self.script_manager.add_category(name, parent)
        
        # Refresh tree
        self._load_data()
        
        # Close dialog
        dialog.destroy()
        
        # Update status
        self.status_var.set(f"Added category: {name}")
    
    def add_script(self):
        """Add a new script"""
        # Create dialog window
        dialog = tk.Toplevel(self)
        dialog.title("Add Script")
        dialog.geometry("500x300")
        dialog.transient(self)
        dialog.grab_set()
        
        # Dialog content
        ttk.Label(dialog, text="Script File:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        file_frame = ttk.Frame(dialog)
        file_frame.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        
        file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=file_var, width=30).pack(side=tk.LEFT)
        ttk.Button(file_frame, text="Browse...", command=lambda: file_var.set(
            filedialog.askopenfilename(filetypes=[("Python Files", "*.py"), ("All Files", "*.*")])
        )).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(dialog, text="Script Name:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Label(dialog, text="Category:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        category_var = tk.StringVar(value="Root")
        category_combo = ttk.Combobox(dialog, textvariable=category_var, width=38)
        category_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Populate category dropdown
        category_combo['values'] = ["Root"]
        category_map = {"Root": ""}
        
        for cat in self.script_manager.config['categories']:
            category_combo['values'] = list(category_combo['values']) + [cat['name']]
            category_map[cat['name']] = cat['id']
        
        ttk.Label(dialog, text="Context:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        context_frame = ttk.Frame(dialog)
        context_frame.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        
        file_context_var = tk.BooleanVar(value=True)
        dir_context_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(context_frame, text="Files", variable=file_context_var).pack(side=tk.LEFT)
        ttk.Checkbutton(context_frame, text="Directories", variable=dir_context_var).pack(side=tk.LEFT)
        
        ttk.Label(dialog, text="Icon (optional):").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        icon_frame = ttk.Frame(dialog)
        icon_frame.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        
        icon_var = tk.StringVar()
        ttk.Entry(icon_frame, textvariable=icon_var, width=30).pack(side=tk.LEFT)
        ttk.Button(icon_frame, text="Browse...", command=lambda: icon_var.set(
            filedialog.askopenfilename(filetypes=[("Icon Files", "*.ico"), ("All Files", "*.*")])
        )).pack(side=tk.LEFT, padx=5)
        
        # Add button
        ttk.Button(dialog, text="Add Script", command=lambda: self._add_script_action(
            file_var.get(), name_var.get(), category_map[category_var.get()],
            file_context_var.get(), dir_context_var.get(), icon_var.get(), dialog
        )).grid(row=5, column=0, columnspan=2, pady=10)
    
    def _add_script_action(self, script_path, name, category, file_context, dir_context, icon_path, dialog):
        """Process adding a script"""
        if not script_path or not os.path.exists(script_path):
            messagebox.showerror("Error", "Script file does not exist")
            return
        
        if not name:
            # Use filename as name if not provided
            name = os.path.basename(script_path).replace('.py', '')
        
        # Build contexts list
        contexts = []
        if file_context:
            contexts.append('file')
        if dir_context:
            contexts.append('directory')
        
        if not contexts:
            messagebox.showerror("Error", "At least one context must be selected")
            return
        
        # Add script
        script_id = self.script_manager.add_script(
            script_path, name, category, contexts,
            icon_path if icon_path else None
        )
        
        # Refresh tree
        self._load_data()
        
        # Close dialog
        dialog.destroy()
        
        # Update status
        self.status_var.set(f"Added script: {name}")
    
    def edit_item(self):
        """Edit the selected item"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Info", "No item selected")
            return
        
        # Item is already selected, details are already showing
        # Just focus on the appropriate input field
        item_id = selection[0]
        item_type = self.tree.item(item_id, "values")[0]
        
        if "Category" in item_type:
            self.category_name_entry.focus_set()
        elif "Script" in item_type:
            self.script_name_entry.focus_set()
    
    def delete_item(self):
        """Delete the selected item"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Info", "No item selected")
            return
        
        item_id = selection[0]
        item_type = self.tree.item(item_id, "values")[0]
        item_name = self.tree.item(item_id, "text")
        
        # Confirm deletion
        if "Category" in item_type:
            # Check if has scripts or subcategories
            has_children = len(self.tree.get_children(item_id)) > 0
            
            if has_children:
                confirm = messagebox.askyesno(
                    "Confirm Deletion",
                    f"Category '{item_name}' contains scripts or subcategories. "
                    f"All contained items will also be deleted. Continue?"
                )
            else:
                confirm = messagebox.askyesno("Confirm Deletion", f"Delete category '{item_name}'?")
            
            if confirm:
                # Find category ID
                category_id = None
                for cat_id, tree_id in self.category_map.items():
                    if tree_id == item_id:
                        category_id = cat_id
                        break
                
                if category_id:
                    # Delete category
                    self.script_manager.remove_category(category_id)
                    
                    # Refresh tree
                    self._load_data()
                    
                    # Update status
                    self.status_var.set(f"Deleted category: {item_name}")
                    
                    # Hide details
                    self.no_selection_label.pack(pady=50)
                    self.category_frame.pack_forget()
                    self.script_frame.pack_forget()
        
        elif "Script" in item_type:
            confirm = messagebox.askyesno("Confirm Deletion", f"Delete script '{item_name}'?")
            
            if confirm:
                # Get script ID from tags
                tags = self.tree.item(item_id, "tags")
                if tags:
                    script_id = tags[0]
                    
                    # Delete script
                    self.script_manager.remove_script(script_id)
                    
                    # Refresh tree
                    self._load_data()
                    
                    # Update status
                    self.status_var.set(f"Deleted script: {item_name}")
                    
                    # Hide details
                    self.no_selection_label.pack(pady=50)
                    self.category_frame.pack_forget()
                    self.script_frame.pack_forget()
    
    def save_category(self):
        """Save changes to a category"""
        category_id = getattr(self.category_frame, 'category_id', None)
        if not category_id:
            # Root category can't be edited
            self.status_var.set("Cannot edit root category")
            return
        
        category_name = self.category_name_var.get()
        if not category_name:
            messagebox.showerror("Error", "Category name cannot be empty")
            return
        
        # Get parent ID
        parent_name = self.category_parent_var.get()
        parent_id = self.category_parent_map.get(parent_name, "")
        
        # Find category in config
        category = next((c for c in self.script_manager.config['categories'] if c['id'] == category_id), None)
        if category:
            # Update category
            category['name'] = category_name
            category['parent'] = parent_id
            
            # Save config
            self.script_manager.save_config()
            
            # Update registry entries for all scripts in this category
            for script in self.script_manager.config['scripts']:
                if script['category'] == category_id:
                    self.script_manager._remove_registry_entries(script)
                    self.script_manager._create_registry_entries(script)
            
            # Refresh tree
            self._load_data()
            
            # Update status
            self.status_var.set(f"Updated category: {category_name}")
    
    def save_script(self):
        """Save changes to a script"""
        script_id = getattr(self.script_frame, 'script_id', None)
        if not script_id:
            return
        
        script_name = self.script_name_var.get()
        if not script_name:
            messagebox.showerror("Error", "Script name cannot be empty")
            return
        
        # Get category ID
        category_name = self.script_category_var.get()
        category_id = self.script_category_map.get(category_name, "")
        
        # Build contexts list
        contexts = []
        if self.file_context_var.get():
            contexts.append('file')
        if self.dir_context_var.get():
            contexts.append('directory')
        
        if not contexts:
            messagebox.showerror("Error", "At least one context must be selected")
            return
        
        # Get icon path
        icon_path = self.script_icon_var.get()
        
        # Find script in config
        script = next((s for s in self.script_manager.config['scripts'] if s['id'] == script_id), None)
        if script:
            # Get script path
            script_path = os.path.join(self.script_manager.scripts_dir, script['filename'])
            
            # Save script content
            try:
                content = self.script_content.get('1.0', tk.END)
                with open(script_path, 'w') as f:
                    f.write(content)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save script content: {e}")
                return
            
            # Update script
            self.script_manager.update_script(
                script_id, script_name, category_id, contexts,
                icon_path if icon_path and os.path.exists(icon_path) else None
            )
            
            # Refresh tree
            self._load_data()
            
            # Update status
            self.status_var.set(f"Updated script: {script_name}")
    
    def set_python_path(self):
        """Set the Python executable path"""
        python_path = filedialog.askopenfilename(
            title="Select Python Executable",
            filetypes=[("Python Executable", "python*.exe"), ("All Files", "*.*")]
        )
        
        if python_path:
            self.script_manager.config['python_path'] = python_path
            self.script_manager.save_config()
            
            # Update status
            self.status_var.set(f"Python path set to: {python_path}")
            
            # Ask if user wants to update registry entries
            if messagebox.askyesno(
                "Update Registry",
                "Do you want to update all registry entries with the new Python path?"
            ):
                # Update all scripts
                for script in self.script_manager.config['scripts']:
                    self.script_manager._remove_registry_entries(script)
                    self.script_manager._create_registry_entries(script)
                
                self.status_var.set("Updated all registry entries with new Python path")
    
    def open_scripts_dir(self):
        """Open the scripts directory"""
        os.startfile(self.script_manager.scripts_dir)
    
    def export_config(self):
        """Export configuration to a file"""
        export_path = filedialog.asksaveasfilename(
            title="Export Configuration",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        
        if export_path:
            try:
                with open(export_path, 'w') as f:
                    json.dump(self.script_manager.config, f, indent=4)
                
                self.status_var.set(f"Configuration exported to: {export_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export configuration: {e}")
    
    def import_config(self):
        """Import configuration from a file"""
        import_path = filedialog.askopenfilename(
            title="Import Configuration",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        
        if import_path:
            try:
                with open(import_path, 'r') as f:
                    config = json.load(f)
                
                # Validate config
                if not all(key in config for key in ['categories', 'scripts', 'python_path']):
                    messagebox.showerror("Error", "Invalid configuration file format")
                    return
                
                # Confirm import
                if messagebox.askyesno(
                    "Confirm Import",
                    "Importing will replace your current configuration. Continue?"
                ):
                    # Remove all current registry entries
                    for script in self.script_manager.config['scripts']:
                        self.script_manager._remove_registry_entries(script)
                    
                    # Set new config
                    self.script_manager.config = config
                    self.script_manager.save_config()
                    
                    # Create new registry entries
                    for script in self.script_manager.config['scripts']:
                        self.script_manager._create_registry_entries(script)
                    
                    # Refresh tree
                    self._load_data()
                    
                    self.status_var.set("Configuration imported successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import configuration: {e}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
        Advanced Context Menu Manager
        
        A tool for managing Windows context menu entries for scripts
        with support for categories and subcategories.
        
        Version 1.0
        """
        
        messagebox.showinfo("About", about_text)

def is_admin():
    """Check if the script is running with administrator privileges"""
    try:
        return subprocess.run(["net", "session"], stdout=subprocess.PIPE, stderr=subprocess.PIPE).returncode == 0
    except Exception:
        return False

if __name__ == "__main__":
    # Check for admin rights
    if not is_admin():
        messagebox.showwarning(
            "Administrator Privileges Required",
            "This application needs administrator privileges to modify the Windows registry. "
            "Please run it as administrator."
        )
        sys.exit(1)
    
    app = AdvancedContextMenuGUI()
    app.mainloop()