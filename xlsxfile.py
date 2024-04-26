import os
import sys

import pandas as pd
import xml.etree.ElementTree as et
from lxml import etree


class XlsxFile:
    def __init__(self, location):
        # Store the location of the Excel file and initialize variables.
        self.location = location
        # 'df' will hold the data from the Excel file.
        self.df = None
        # Open the Excel file using pandas.
        self.xlsx_file = pd.ExcelFile(location)
        # Initialize dictionaries to map issues and their IDs.
        self.issue_element_map = {}
        self.issue_id_map = {}

    def get_sheet_names(self):
        # Return the names of all sheets in the Excel file.
        return self.xlsx_file.sheet_names

    def convert_to_xml(self, sheet_name, xml_file):
        try:
            # Create the root element for the XML file.
            typology_file = et.Element("typology", edition="MQM2021")

            # Check if the selected worksheet exists in the Excel file.
            if sheet_name not in self.get_sheet_names():
                return "Couldn't identify a worksheet to open"

            # Parse the specified sheet and fill any missing values
            mqm_index = self.find_mqm_and_prepare_df(sheet_name)
            self.df.fillna("", inplace=True)

            # Reset the issue maps for a fresh conversion.
            self.issue_element_map = {}
            self.issue_id_map = {}
            # Parse the worksheet and check for any errors.
            success, message = self.parse_worksheet(mqm_index)
            if not success:
                return success, message

            # Recursively nest error type elements in the XML structure.
            ids = self.issue_id_map[""]
            self.nest_error_type_elements_recursively(ids, typology_file)

            # Write the XML structure to a file.
            tree = et.ElementTree(typology_file)
            et.indent(tree, space="\t", level=0)
            tree.write(xml_file, encoding="utf-8", xml_declaration=True)

            # Validate the generated XML file.
            return self.validate_xml(xml_file)
        except Exception as e:
            # Return False and the exception message if an error occurs.
            return False, str(e)

    def parse_worksheet(self, mqm_index):
        # Identify the columns in the worksheet by their headers.
        header_row = self.df.iloc[0]
        name_column = None
        id_column = None
        parent_column = None
        description_column = None
        examples_column = None
        notes_column = None
        pid_column = None

        for idx, cell in enumerate(header_row):
            header = cell.lower()
            # Map column names to their indices based on headers.
            if "name" in header:
                name_column = idx
            elif "type id" in header:
                id_column = idx
            elif "parent" in header:
                parent_column = idx
            elif "description" in header:
                description_column = idx
            elif "examples" in header:
                examples_column = idx
            elif "notes" in header:
                notes_column = idx
            elif "type pid" in header:
                pid_column = idx

        # Check if the necessary columns were found.
        if None in (name_column, description_column, examples_column, notes_column):
            return False, "The necessary columns were not found. Expected Name, Description, Examples, and Notes."

        # Process each row in the worksheet. The first row is the header, so start at index 1.
        for index in range(1, len(self.df)):
            row = self.df.iloc[index]  # Retrieve the current row of data.

            # Extract the 'name' value from the row. This is the name of the error type.
            name = row.iloc[name_column]

            # If the 'name' is empty, return an error message indicating the problem.
            if not name:
                return False, f"An error cannot have a blank id. Row: {index + mqm_index}"

            # Extract the 'id' value from the row. This is the unique identifier for the error type.
            row_id = row.iloc[id_column]

            # If the 'id' is empty, return an error message indicating the problem.
            if not row_id:
                return False, f"An error cannot have a blank id. Error name: {name}, Row: {index + 2}"

            # Extract the 'parent', 'description', 'examples', and 'notes' values from the row.
            parent = row.iloc[parent_column]
            description = row.iloc[description_column].replace('\n', '<br/>').strip()
            examples = row.iloc[examples_column].replace('\n', '<br/>').strip()
            notes = row.iloc[notes_column].replace('\n', '<br/>').strip()
            pid = row.iloc[pid_column].strip()

            # Create an XML element for the error type and set its 'name' and 'id' attributes.
            element = et.Element("errorType")
            element.set("name", name.strip())
            element.set("id", row_id)
            element.set("PID", pid)

            # Create sub-elements for the 'description', 'notes', and 'examples' values.
            description_element = et.Element("description")
            description_element.text = description
            notes_element = et.Element("notes")
            notes_element.text = notes
            examples_element = et.Element("examples")
            examples_element.text = examples

            # Append the sub-elements to the error type element.   
            element.append(description_element)
            element.append(notes_element)
            element.append(examples_element)

            # Store the error type element in the issue_element_map dictionary.
            if row_id not in self.issue_element_map:
                # If the row_id is not in the map, store the element with the row_id as the key.
                self.issue_element_map[row_id] = element
            else:
                # If the row_id is already in the map, append a '2' to the key and store the element.
                self.issue_element_map[row_id + str(mqm_index)] = element

            # Store the row_id in the issue_id_map dictionary with the parent as the key.
            if parent not in self.issue_id_map:
                # If the parent is not in the map, store the row_id in a list with the parent as the key.
                self.issue_id_map[parent] = [row_id]
            else:
                # Store the row_id in the list for the parent key.
                if row_id in self.issue_id_map[parent]:
                    # If the row_id is already in the list, append a '2' to the row_id and store it.
                    self.issue_id_map[parent].append(row_id + mqm_index)
                else:
                    # If the row_id is not in the list, store it in the list for the parent key.
                    self.issue_id_map[parent].append(row_id)

        # After processing all rows, return True to indicate successful parsing with no error message.
        return True, ""

    def find_mqm_and_prepare_df(self, sheet_name):
        # Read the whole sheet to a temporary DataFrame
        temp_df = pd.read_excel(self.location, sheet_name=sheet_name, header=None)
        mqm_row_index = None

        # Find the row index containing "MQM"
        for index, row in temp_df.iterrows():
            if row.str.contains("MQM").any():
                mqm_row_index = index
                break

        # If "MQM" is found, read the file again, setting headers appropriately
        if mqm_row_index is not None:
            self.df = pd.read_excel(self.location, sheet_name=sheet_name, header=mqm_row_index)
        else:
            raise ValueError("MQM not found in any row")
        return mqm_row_index + 2

    def nest_error_type_elements_recursively(self, ids, parent_element, depth=0):
        # Recursively nest error type elements in the XML structure based on their parent-child relationships.
        for row_id in ids:
            # Get the error type element for the current row_id.
            element = self.issue_element_map[row_id]
            # Set the 'level' attribute of the element to the current depth in the XML structure.
            element.set("level", str(depth))
            if row_id in self.issue_id_map:
                # If the current row_id has children, recursively nest the children under the current element.
                self.nest_error_type_elements_recursively(self.issue_id_map[row_id], element, depth + 1)

            # Append the current element to the parent element.
            parent_element.append(element)

    @staticmethod
    def validate_xml(xml_file):
        # Validate the generated XML file against a schema.

        # Get the directory where the application's resources are stored.
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
        source_file = os.path.join(bundle_dir, "typologySchema.xsd")

        # Parse the schema and the XML file.
        xsd_schema = etree.XMLSchema(etree.parse(source_file))
        # Parse the XML file.
        lxml_root_element = etree.parse(xml_file)
        try:
            # Attempt to validate the XML file against the schema.
            xsd_schema.assertValid(lxml_root_element)
            # If the validation is successful, return True and an empty string.
            return True, ""
        except etree.DocumentInvalid:
            # If the validation fails, return False and an error message.
            return False, "File was written, but failed validation"
