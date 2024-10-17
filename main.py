'''              ---------   Imported Libraries  ----------                         '''
import logging.config
from openpyxl import load_workbook
import tkinter as tk
import pandas as pd
import logging
import colorlog
from tkinter import ttk
from tkinter import messagebox, simpledialog
from openpyxl.worksheet.worksheet import Worksheet

'''                 ---------   Constants   -----------                            '''
full_path ='' 
num_of_files =''

'''                 ---------   Logging configuration   ---------                   '''
class Logger:
    def __init__(self, logger_name='my_logger', log_file='app.log'):
        # Create a logger
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(logging.DEBUG)  # Ensure this is set to DEBUG

        # Create handlers for file and terminal
        self._create_file_handler(log_file)
        self._create_terminal_handler()

    def _create_file_handler(self, log_file):
        """Creates and adds a file handler to the logger."""
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.DEBUG)  # Log everything to the file
        file_formatter = logging.Formatter(
                '%(asctime)s - [%(levelname)s] - %(message)s',
                 datefmt='%Y-%m-%d %H:%M:%S'
                    )
        file_handler.setFormatter(file_formatter)
        self.logger.addHandler(file_handler)

    def _create_terminal_handler(self):
        """Creates and adds a terminal handler with color support to the logger."""
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.DEBUG)  # Ensure this is set to DEBUG

        colored_formatter = colorlog.ColoredFormatter(
            "%(log_color)s[%(asctime)s] - [%(levelname)s] - %(message)s",
            datefmt='%Y-%m-%d %H:%M:%S',
            log_colors={
                'DEBUG': 'cyan',
                'INFO': 'green',
                'WARNING': 'yellow',
                'ERROR': 'red',
                'CRITICAL': 'bold_red',
            }
        )
        stream_handler.setFormatter(colored_formatter)
        self.logger.addHandler(stream_handler)

    def get_logger(self):
        """Returns the logger instance."""
        return self.logger

custom_logger = Logger().get_logger()        
    
def get_list_of_paths(df:pd.DataFrame) ->list:
    """_summary_

    Args:
        df (pd.DataFrame): Data frame of the workSheet

    Returns:
        list: list with the paths of the files 
    """
    list_of_path = list(set(df['Path'].tolist()))
    list_of_path = list(map(lambda element: element[element.find('Documents'):],list_of_path))
    return sorted(list_of_path)
    
def get_num_file_of_node(path:str, df: pd.DataFrame)->int:
    """_summary_

    Args:
        path (str): Path of the folder
        df (pd.DataFrame): data frame of the worksheett

    Returns:
        int: number of files found for that path
    """
    #longest_path = max(paths,key=len)
    longest_path = path
    df['Path'] = df['Path'].str.extract('(Documents.*)', expand=False)
    df = df[(df['Path'].str.contains(longest_path, case=False, na=False, regex=False) ) & (df['File Size'].notna()) ]
    
    return len(df)


def add_tree_items(tree, parent, tree_structure):
    """_summary_

    Args:
        tree (Dict): Dictionary of nodes
        parent (element)
        tree_structure (Dict): Entire tree Structure
    """
    for key, value in tree_structure.items():
        item_id = tree.insert(parent, 'end', text=key)
        if isinstance(value, dict):
            add_tree_items(tree, item_id, value)

def sort_tree(tree):
    """Recursively sorts the keys in the tree structure."""
    sorted_tree = {}
    
    # Sort keys alphabetically and recursively sort children
    for key in sorted(tree, key=lambda x: (x.lower(),x)):
        sorted_tree[key] = sort_tree(tree[key])
    
    return sorted_tree

def build_tree(paths: list)->dict:
    """_summary_

    Args:
        paths (list): list of paths 

    Returns:
        dict: Tree_structure
    """
    tree_structure = {}
    custom_logger.info('Building the Tree Structure . . .')
    
    for path in paths:
        components = path.split('/')
        current_level = tree_structure
        
        for component in components:
            
            if component not in current_level:
                current_level[component] = {}
            
            current_level = current_level[component]
    
    custom_logger.info("Tree Structure Built Successfully.")
    return sort_tree(tree_structure)

def get_full_path_of_selected(tree,item):
    """_summary_

    Args:
        tree (Node)
        item (Selected element)

    Returns:
        Recursively collects the path of the nodes
    """
    parent = tree.parent(item)
    if parent:
        return get_full_path_of_selected(tree,parent)+ '/'+ tree.item(item)['text']
    else:
        return tree.item(item)['text']
    
    
def on_tree_item_click(event, tree, label, df:pd.DataFrame):
    global full_path
    global num_of_files
    """_summary_

    Args:
        event (When user clicks)
        tree (Node)
        label (INFO)
        df (Data Frame)
    """
    try:
        selected_item = tree.selection()[0]
    except IndexError:
        return
    full_path = get_full_path_of_selected(tree,selected_item)
    num_of_files = get_num_file_of_node(full_path,df)
    
    label.config(text= f'Path: {full_path}\nNumber of files: {num_of_files}')

def on_key_press(event,tree,label,df):
    if event.keysym in ["Up", 'Down']:
        on_tree_item_click(event,tree,label,df)

def flag_path():
    global full_path
    global num_of_files
    
    additional_information = simpledialog.askstring('Input',"You flagged this File. Enter additional info if applicable: ")
    if additional_information:
        logged_message = f'Path: {full_path} - Number of files: {num_of_files} - Info: {additional_information}'
    else:
        logged_message = f'Path: {full_path} - Number of files: {num_of_files}'
    
    custom_logger.debug(logged_message)
    messagebox.showinfo('INFO','You flagged this item to logs.txt')


def on_key_search(event,searchEntry:tk.Entry,full_path:str, tree_structure:dict):
    
    search_term = searchEntry.get().strip()
    current_level = tree_structure
    match_found=False
    
    
    if not search_term:
        searchEntry.config(bg='white')
        return
    
   # Split the full path and check if it has components
    current_path = full_path.split('/')
    if not current_path:  # If empty, return early
        return
    #path exists for sure 
    for component in current_path:
        if component in current_level:
            current_level = current_level[component] 
    
    #Set the current level to the right one and now to traverse the tree until element is found
    for item in current_level.keys():
        if search_term == item:
            match_found=True
            break
    
    searchEntry.config(bg='green' if match_found else 'red')
    

def on_closing(self:tk.Tk)->None:
    custom_logger.info('User Terminated Program.')
    self.destroy()
      
def create_gui(tree_structure, df: pd.DataFrame):
    global full_path
    
    root = tk.Tk()
    root.title("Folder Composition")
    root.geometry('1000x500')
    root.protocol('WM_DELETE_WINDOW', lambda: on_closing(root))

    # Create a frame to hold the treeview and the scrollbar
    frame = tk.Frame(root)
    frame.pack(expand=True, fill='both')

    # Create a vertical scrollbar
    tree_scroll = ttk.Scrollbar(frame)
    tree_scroll.pack(side='right', fill='y')

    # Create the treeview and link the scrollbar
    tree = ttk.Treeview(frame, yscrollcommand=tree_scroll.set)
    tree.pack(expand=True, fill='both')

    # Configure the scrollbar to work with the treeview
    tree_scroll.config(command=tree.yview)

    # Add items to the treeview
    add_tree_items(tree, '', tree_structure)

    # Make treeview columns stretchable
    tree.column('#0', stretch=True)

    # Selected item label
    label = tk.Label(root, text="Selected: ", fg='red', font=('Arial', 12, 'bold'))
    label.pack(side='bottom', fill='x')

    # Flag button
    flagButton = tk.Button(master=root, text='Flag', fg='white', bg='red',
                           font=('Arial', 10, 'bold'), borderwidth=2, command=flag_path)
    flagButton.pack(side='right', fill='none')

    # Search entry
    searchEntry = tk.Entry(master=root)
    searchEntry.config(borderwidth=2, font=('Arial', 10, 'bold'), fg='white')
    searchEntry.pack(side='right', fill='none', padx=(0, 50))

    # Search label
    searchLabel = tk.Label(master=root, text='Search within Folder: ', font=('Arial', 10, "bold"))
    searchLabel.pack(side='right', fill='none')

    # Bind events
    tree.bind("<ButtonRelease-1>", lambda event: on_tree_item_click(event, tree, label, df))
    tree.bind("<KeyRelease>", lambda event: on_key_press(event, tree, label, df))
    searchEntry.bind("<KeyRelease>", lambda event: on_key_search(event, searchEntry, full_path, tree_structure))

    root.mainloop()


def main() -> int:
    """_summary_

    Returns:
        int: return 0 for good and 1 for bad
    """
    custom_logger.info('User ran the Program.')
    file_name = input('Enter the file name: ')
    
    try:
        workBook = load_workbook(file_name)
        workSheet = workBook.active
    except Exception as e:
        custom_logger.critical(e)
        custom_logger.error('Program will terminate . . . ')
        return 1
    
    data = workSheet.values
    
    columns = next(data)
    
    df = pd.DataFrame(data, columns=columns)
    
    paths = get_list_of_paths(df)
    
    tree_structure = build_tree(paths)
    
    create_gui(tree_structure, df)
    
    return 0

if __name__ =="__main__": 
    main()
    
    

    
    
