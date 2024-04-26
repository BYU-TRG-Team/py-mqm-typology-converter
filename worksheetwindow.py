import os
import sys
import tkinter as tk

from PIL import Image, ImageTk

from xlsxfile import XlsxFile


class WorksheetWindow:
    def __init__(self, root, input_file, output_file):
        # Initialize variables for storing any exceptions and the success status.
        self.exception_str = None
        self.success = None

        # Set the root window and its properties.
        self.root = root

        # Set up the main window for the WorksheetWindow.
        self.root = root
        self.root.title("Worksheet Selection")  # Set the window title.
        self.root.geometry("450x500")  # Define the size of the window.
        self.root.resizable(False, False)  # Prevent resizing of the window.
        self.output_file = output_file  # Store the output file path.

        # Get the directory where the application's resources are stored.
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
        source_file = os.path.join(bundle_dir, "mqm_logo.png") # Path to the logo image.

        # Set the application icon.
        png_image = Image.open(source_file)
        png_image = png_image.resize((32, 32))

        # Convert the image to a format that Tkinter can use and set it as the window icon.
        icon = ImageTk.PhotoImage(png_image)
        root.iconphoto(False, icon)

        # Load the input file into the XlsxFile class to handle Excel operations.
        self.worksheet = XlsxFile(input_file)

        # Create a canvas to hold the window's widgets.
        self.canvas = tk.Canvas(root, background="white")
        self.canvas.pack(fill=tk.BOTH, expand=True) # Make the canvas fill the window.

        # Add a label to instruct the user to select a worksheet.
        label = tk.Label(self.canvas, text="Select the worksheet containing the typology to convert.",
                         font=("Arial", 13), background="white")
        label.place(x=20, y=20) # Position the label on the canvas

        # Create a listbox to display the available worksheets in the input file.
        self.worksheet_listbox = tk.Listbox(self.canvas, selectmode=tk.SINGLE, height=20, width=71)
        self.worksheet_listbox.place(x=10, y=70)
        # Populate the listbox with the names of the worksheets.
        for worksheet_name in self.worksheet.get_sheet_names():
            self.worksheet_listbox.insert(tk.END, worksheet_name)

        # Add a button for the user to select a worksheet from the list.
        select_button = tk.Button(self.canvas, text="Select Worksheet", height=2, width=20,
                                  command=self.select_worksheet)
        select_button.place(x=150, y=425) # Position the button on the canvas.

    def select_worksheet(self):
        # Get the selected index from the listbox.
        selected_index = self.worksheet_listbox.curselection()
        if selected_index:
            # Retrieve the name of the selected worksheet.
            selected_worksheet = self.worksheet_listbox.get(selected_index)
            # Convert the selected worksheet to XML and save it to the output file.
            # 'success' will be True if the operation succeeded, and 'exception_str' will contain any error message.
            self.success, self.exception_str = self.worksheet.convert_to_xml(selected_worksheet, self.output_file)
            # Close the window once the selection is made and the conversion is complete.
            self.root.destroy()
