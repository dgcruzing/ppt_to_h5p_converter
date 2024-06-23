# PowerPoint to H5P Converter


## For Developers

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Setting up the Development Environment

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/ppt-to-h5p-converter.git
   cd ppt-to-h5p-converter
   ```

2. Create a virtual environment:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

### Running the Script

```
python ppt_to_h5p_converter.py
```

### Building the Executable

We use PyInstaller to create the executable. To build:

1. Ensure you're in the project directory and your virtual environment is activated.
2. Run:
   ```
   pyinstaller ppt_to_h5p_converter.spec
   ```
3. The executable will be created in the `dist` directory.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check the [issues page](https://github.com/yourusername/ppt-to-h5p-converter/issues) if you want to contribute.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
