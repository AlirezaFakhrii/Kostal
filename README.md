# Excel Processing Application

This application processes an Excel document by splitting requirements based on the occurrence of LK_SWC entries and adds a parent reference ID column to the output.

## Features

- **Read the Excel Document**: Load the Excel file into a Pandas DataFrame.
- **Identify the Requirements**: Filter the DataFrame to identify rows where `LK_ContentType` is 'Req'.
- **Split Requirements**: Split the requirements based on the count of `LK_SWC`.
- **Add Parent Reference Column**: Add a new column to the DataFrame that contains the parent Requirement ID for each requirement.
- **Save the Modified Data**: Save the modified DataFrame back to an Excel file.
- **User-Friendly Interface**: Provide a user-friendly interface using a file dialog for selecting and saving files.

## Usage Instructions

### Selecting the Excel File

1. Click the **"Select Excel File"** button.
2. Use the file dialog to choose the Excel file you want to process.

### Processing the Excel File

1. After selecting a file, click the **"Process and Save as Excel"** button.
2. The application will read the file, process the data, and prompt you to save the processed file.

### Opening the Processed File

- Once the file is saved, you can open it by clicking the **"Open Processed File"** button.

## Dependencies

- **tkinter**: For creating the graphical user interface.
- **pandas**: For data manipulation and processing.
- **openpyxl**: For reading from and writing to Excel files.
