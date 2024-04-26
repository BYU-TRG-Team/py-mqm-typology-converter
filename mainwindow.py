import os
import shutil
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

from PIL import Image, ImageTk

from worksheetwindow import WorksheetWindow


class MainWindow:
    def __init__(self, root):
        """
        Initializes the main window of the MQM Typology Converter application.

        Parameters:
        - root: The root Tkinter object representing the main window.

        Returns:
        None
        """

        # Set the main window to be the 'root' of the application.
        self.root = root
        # Set the title of the main window.
        self.root.title("MQM Typology Converter")
        # Set the size of the main window.
        self.root.geometry("350x370")
        # Set the main window to be non-resizable.
        self.root.resizable(False, False)

        # Find the directory where the application's resources are stored.
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))

        # Specify the path to the logo image file.
        image_path = os.path.join(bundle_dir, "mqm_logo.png")

        # Open the logo image and resize it to fit the application.
        png_image = Image.open(image_path)
        png_image = png_image.resize((32, 32))

        # Convert the image to a format that Tkinter can use.
        icon = ImageTk.PhotoImage(png_image)

        # Set the window's icon to be the application's logo.
        root.iconphoto(False, icon)

        # Create a menu bar for the main window.
        menu = tk.Menu(root)
        # Create a 'File' menu for the menu bar.
        file_menu = tk.Menu(menu, tearoff=0)
        # Add a 'Save log to file' option to the 'File' menu.
        file_menu.add_command(label="Save log to file", command=self.export_to_txt)
        # Add a separator to the 'File' menu.
        file_menu.add_separator()
        # Add an 'Exit' option to the 'File' menu that closes the application.
        file_menu.add_command(label="Exit", command=root.quit)
        # Add the 'File' menu to the menu bar.
        menu.add_cascade(label="File", menu=file_menu)

        # Create an 'Other' menu for the menu bar.
        other_menu = tk.Menu(menu, tearoff=0)
        # Add a 'Save Typology Schema' option to the 'Other' menu.
        other_menu.add_command(label="Save Typology Schema", command=self.save_topology_schema)
        # Add the 'Other' menu to the menu bar.
        menu.add_cascade(label="Other", menu=other_menu)
        # Attach the menu bar to the main window.
        root.config(menu=menu)

        # Define a font style for error messages.
        error_font = ("Arial", 10, "bold")

        # Create and position a label for the spreadsheet input.
        source_label = tk.Label(root, text="MQM Typology Spreadsheet Path:")
        source_label.grid(row=1, column=0, sticky="w", padx=10, pady=(10, 0))
        # Create an input field for the spreadsheet path.
        self.source_path_input = tk.Entry(root, width=40)
        self.source_path_input.grid(row=2, column=0, sticky="w", padx=10, pady=5)
        # Create a 'Browse' button that opens a file dialog to select the spreadsheet.
        browse_source_button = tk.Button(root, text="Browse", command=self.browse_source_path, width=10)
        browse_source_button.grid(row=2, column=2, pady=5)
        # Create a label for displaying errors related to the source input.
        self.source_error_label = tk.Label(root, text="", fg="red", font=error_font)
        self.source_error_label.grid(row=3, column=0, columnspan=3, pady=(0, 5))

        # Create and position a label for the XML output path.
        target_label = tk.Label(root, text="Output Typology XML Path:")
        target_label.grid(row=4, column=0,  sticky="w", padx=10, pady=(10, 0))
        # Create an input field for the XML output path.
        self.target_path_input = tk.Entry(root, width=40)
        self.target_path_input.grid(row=5, column=0, sticky="w", padx=10, pady=5)
        # Create a 'Browse' button that opens a file dialog to select the XML output path.
        browse_target_button = tk.Button(root, text="Browse", command=self.browse_target_path, width=10)
        browse_target_button.grid(row=5, column=2, pady=5)
        # Create a label for displaying errors related to the target input.
        self.target_error_label = tk.Label(root, text="", fg="red", font=error_font)
        self.target_error_label.grid(row=6, column=0, columnspan=3, pady=(0, 5))

        # Create a 'Convert' button that triggers the conversion process.
        convert_button = tk.Button(root, text="Convert", command=self.convert, width=10)
        convert_button.grid(row=7, column=0, columnspan=3, pady=5)

        # Create a text box for displaying messages to the user.
        self.message_textbox = tk.Text(root, wrap=tk.WORD, height=6, width=40, state=tk.DISABLED)
        self.message_textbox.grid(row=8, column=0, columnspan=3)

        # Create a label for displaying the application's version number.
        version_label = tk.Label(root, text="v1.1.0.0")
        version_label.grid(row=9, column=2, sticky="e")
        # Validate the source and target input fields when the user types in them.
        self.validate_source_input(None)
        self.validate_target_input(None)
        self.source_path_input.bind("<KeyRelease>", self.validate_source_input)
        self.target_path_input.bind("<KeyRelease>", self.validate_target_input)
        # Configure the text box to display messages in different colors.
        self.message_textbox.tag_configure("green", background="green")
        self.message_textbox.tag_configure("red", background="red")

    def validate_source_input(self, _):
        """
        Validates the source input file path and displays any error message.

        Parameters:
        - _: Placeholder parameter (not used)

        Returns:
        - None
        """
        # Get the file path from the source input field.
        input_file = self.source_path_input.get()
        # Validate the file path and display any error message.
        source_error = self.validate_file(input_file, ".xlsx", True)
        # Update the error label with the error message.
        self.source_error_label.config(text=source_error)

    def validate_target_input(self, _):
        """
        Validates the target input file path and displays any errors.

        Parameters:
        - _: Placeholder parameter (not used)

        Returns:
        - None
        """
        # Get the file path from the target input field.
        output_file = self.target_path_input.get()
        # Validate the file path and display any error message.
        target_error = self.validate_file(output_file, ".xml", False)
        # Update the error label with the error message.
        self.target_error_label.config(text=target_error)

    @staticmethod
    def validate_file(file_path, file_extension, check_for_existence):
        """
            Validates the given file path.

            Args:
                file_path (str): The path of the file to be validated.
                file_extension (str): The expected file extension.
                check_for_existence (bool): Flag to indicate whether to check for file existence.

            Returns:
                str: An error message if the file path is invalid, otherwise an empty string.
        """
        # Check if the file path is empty
        if not file_path:
            return "path cannot be empty"
        # Check if the file exists (if we need to check existence) and if it ends with the expected file extension.
        # This checks two things: 
        # 1. If 'check_for_existence' is True, then it checks whether the file actually exists at the given path.
        # 2. It checks whether the file has the correct file extension (like '.txt', '.xlsx', etc.).
        # If any of these checks fail, return an error message indicating the requirement.
        if (not (os.path.exists(file_path) or not check_for_existence)
                or not file_path.lower().endswith(file_extension)):
            return f"must be a valid path to {file_extension.upper()} file"
        return ""

    def browse_source_path(self):
        """
            Opens a file dialog to browse and select an Excel file (.xlsx) as the source path.
            Updates the source path input field with the selected file path.
            Calls the validate_source_input method to validate the selected file path.
        """
        source_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if source_path:
            self.source_path_input.delete(0, tk.END)
            self.source_path_input.insert(0, source_path)
            self.validate_source_input(None)

    def browse_target_path(self):
        """
            Opens a file dialog to browse and select an XML file (.xml) as the target path.
            Updates the target path input field with the selected file path.
            Calls the validate_target_input method to validate the selected file path.
        """
        target_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML Files", "*.xml")])
        if target_path:
            self.target_path_input.delete(0, tk.END)
            self.target_path_input.insert(0, target_path)
            self.validate_target_input(None)

    def convert(self):
        """
            Initiates the conversion process and displays the result in the message text box.
            
            If the source or target paths are invalid, displays an error message and returns.
            Otherwise, opens a new window to display the worksheet and waits for the user to close it.
            After the user closes the window, displays the result of the conversion in the message text box.
        """
        # Enable the message text box to display the result.
        self.message_textbox.config(state=tk.NORMAL)

        # Get the error messages from the source and target input fields.
        source_error = self.source_error_label.cget("text")
        target_error = self.target_error_label.cget("text")

        # If there are any errors in the source or target path, display a message and stop the conversion.
        if source_error or target_error:
            self.message_textbox.insert(tk.END, "Please provide paths first.\n", "red")
            return

        # Get the file paths from the source and target input fields.
        input_file = self.source_path_input.get()
        output_file = self.target_path_input.get()

        # Create a new window to display the worksheet sheet names.
        worksheet_root = tk.Toplevel(self.root)
        # Disable the main window while the worksheet window is open.
        worksheet_root.grab_set()
        # Create a new WorksheetWindow object to display the worksheet and wait for the user to choose a sheet.
        ww = WorksheetWindow(worksheet_root, input_file, output_file)
        # Wait for the user to close the worksheet window.
        self.root.wait_window(worksheet_root)
    
        # Retrieve the success status and any exception message from the worksheet operation.
        success = ww.success
        exception = ww.exception_str

        # If user doesn't select a sheet and closes the window, return.
        if success is None:
            return
        elif success:
            # If the conversion was successful, display a success message.
            self.message_textbox.insert(tk.END, "Conversion complete and validation successful!\n", "green")
        else:
            # If the conversion failed, display an error message.
            self.message_textbox.insert(tk.END, f"Conversion failed: {exception}\n", "red")

        # Disable the message text box to prevent further editing.
        self.message_textbox.config(state=tk.DISABLED)
        # Release the main window to allow user interaction.
        worksheet_root.grab_release()

    @staticmethod
    def save_topology_schema():
        # Open a file dialog to select the location to save the schema file.
        output_file = filedialog.asksaveasfilename(defaultextension=".xsd", filetypes=[("XSD Files", "*.xsd")],
                                                   initialfile="mqmTypology.xsd")
        
        # If the user closes the file dialog, return.
        if not output_file:
            return

        # Find the directory where the application's resources are stored.
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))

        # Specify the path to the schema file.
        source_file = os.path.join(bundle_dir, "typologySchema.xsd")

        # Copy the schema file to the selected location.
        shutil.copy(source_file, output_file)

    def export_to_txt(self):
        """
        Opens a file dialog to save the log messages to a text file.
        """
        # Get the current date to use in the default file name.
        today_date = datetime.now().strftime('%Y-%m-%d')
        default_file_name = f"mqmTypologyConverterLog_{today_date}.txt"

        # Open a file dialog to select the location to save the log file.
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=default_file_name,
                                                 filetypes=[("Text Files", "*.txt"),
                                                            ("Log Files", "*.log")])

        # If a file path is selected, proceed to save the log
        if file_path:
            # Open the file and write the log messages to it.
            with open(file_path, 'w') as file:
                # Get the text content from the message text box.
                text_content = self.message_textbox.get("1.0", tk.END)
                # Write the text content to the file.
                file.write(text_content)
