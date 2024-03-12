# Word to PDF Converter

This project is a small Python script that converts a Microsoft Word document (.docx) into a PDF file. It utilizes the `comtypes.client` module to interact with Microsoft Word through COM (Component Object Model) and the `docx` module to handle Word document manipulation.

## Usage

1. **Installation**: Make sure you have Python installed on your system. Install the required packages by running the following command:

    ```bash
    pip install comtypes python-docx
    ```

2. **Input**: Place the Word document (`word.docx`) you want to convert into the project directory.

3. **Output**: Specify the desired output PDF file name (`pdf.pdf`) in the script.

4. **Run the Script**: Execute the script `word_to_pdf_converter.py` using Python:

    ```bash
    python word_to_pdf_converter.py
    ```

5. **Result**: The script will convert the Word document into a PDF file and save it in the project directory.

## Dependencies

- `comtypes`: Used to interact with Microsoft Word through COM (Component Object Model).
- `python-docx`: Used to manipulate Word documents in Python.

## Additional Notes

- Ensure that Microsoft Word is installed on your system for this script to work properly.
- The visibility of the Word application is set to False to run the conversion process in the background.
- The script currently supports converting only one Word document at a time.

Feel free to modify and enhance the script according to your specific requirements or integrate it into other projects.

For any issues or suggestions, please open an issue on the GitHub repository.
