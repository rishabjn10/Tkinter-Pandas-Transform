# DataFrame Transformer GUI

This project provides a simple GUI application for importing, transforming, and exporting Excel files using Pandas. The application is built using `tkinter` for the GUI and `pandas` for data manipulation. The final product can be compiled into an executable (.exe) file that can be run on any Windows computer.

## Features

- Import Excel files
- Apply transformations to the data:
  - Add a 'Range' column (difference between 'High' and 'Low' columns)
  - Convert 'Date' column to datetime objects
  - Add a 'Next Month' column (date one month after the 'Date' column)
- Export transformed data to a new Excel file

## Setup

### Prerequisites

- Python 3.11.x
- `pip` (Python package installer)

### Installing Requirements

1. Clone the repository (if using Git):

    ```bash
    git clone https://github.com/yourusername/your-repo-name.git
    cd repo-name
    ```

2. Set up a virtual environment:

    ```bash
    python -m venv venv
    ```

3. Activate the virtual environment:

    - On Windows:

        ```bash
        venv\Scripts\activate
        ```

    - On macOS and Linux:

        ```bash
        source venv/bin/activate
        ```

4. Install the required Python libraries:

    ```bash
    pip install -r requirements.txt
    ```

### Running the Application

1. Run the Python script directly:

    ```bash
    python transform.py
    ```

### Compiling to .exe

1. Install `pyinstaller`:

    ```bash
    pip install pyinstaller
    ```

2. Compile the script into an executable:

    ```bash
    pyinstaller --onefile --windowed transform.py
    ```

3. The executable file will be created in the `dist` directory. You can run it by double-clicking the `.exe` file.

## Using the Application

1. **Import Button**: Opens a file dialog to select an Excel file to import.
2. **Transform Button**: Applies the specified transformations to the imported data.
3. **Export Button**: Opens a file dialog to save the transformed data to a new Excel file.

## Requirements

- `pandas`
- `openpyxl`
- `python-dateutil`

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [Tkinter Documentation](https://docs.python.org/3/library/tkinter.html)
- [PyInstaller Documentation](https://pyinstaller.readthedocs.io/en/stable/)
