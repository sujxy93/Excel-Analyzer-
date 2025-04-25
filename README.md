# Excel-Analyzer
Purpose
The application allows users to:
- Input details about Excel files, including column information and range of data.
- Perform operations (e.g., count, average) on specified data ranges in Excel sheets.
- Visualize the analyzed data using a bar chart.
- Manage the entered data using a TreeView interface (add, delete entries).
  
It combines Tkinter (GUI),openpyxl (Excel manipulation), and Matplotlib (data visualization) modules.

Key Features
1. TreeView Management: 
   - Entries are shown in a TreeView widget, which displays ID, filename, and title.
   - The user can add, delete, or update entries with ease.

2. Excel Processing:
   - User specifies the column and range within an Excel file.
   - Data is analyzed based on its frequency (count of unique values in the range).

3. Graphical Representation:
   - Results are plotted as a bar chart (using Matplotlib), showing the frequency of values.

4. Interactive GUI:
   - Built with Tkinter, it offers entry fields, drop-down menus, and buttons for user interaction.
   - User-friendly prompts and error handling (e.g., ID duplication).

Functions Breakdown
1. insert(): Takes input from various GUI fields and stores the details in a dictionary (`ID`) while updating the TreeView.

2. delete(): Removes the selected entry from the TreeView and the dictionary.

3. ex(): Processes the Excel file based on user inputs, extracts data from the specified column and range, analyzes unique values, and plots a bar chart.

4. exit(): Safely exits the application.

Required Libraries
- openpyxl: For interacting with Excel files (.xlsx).
- matplotlib.pyplot: For data visualization.
- tkinter: To build the graphical interface.
- ttk: Enhanced widgets like TreeView, Combobox.

Flow of Operation
1. User inputs Excel file details (filename, column, range, etc.).
2. Presses "Add" to save the entry and display it in the TreeView.
3. Presses "Execute" to analyze the selected Excel file, visualize the data, and show a bar chart.
4. Entries can be removed or the app can be exited using the respective buttons.
