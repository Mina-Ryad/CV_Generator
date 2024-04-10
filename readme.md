# CV Generator Documentation

## Introduction
The CV Generator script is designed to generate a professional Curriculum Vitae (CV) document in DOCX format based on the provided data in a JSON file. It utilizes the Python `docx` library to create and format the document.

## Usage
To use the CV Generator script, follow these steps:

1. Ensure you have Python installed on your system.
2. Prepare your CV data in a JSON file with the required fields.
3. Run the `generate_CV.py` script, providing the path to the JSON file as an argument.
4. The script will generate a DOCX file containing the formatted CV.

## Script Overview
The script consists of the following main components:

- `generate_CV.py`: The main script file responsible for reading CV data from a JSON file and generating the CV document.
- `cv_data.json`: A sample JSON file containing placeholder data for generating the CV. Replace this file with your actual CV data.

## Dependencies
The script relies on the following Python libraries:
- `docx`: For creating and formatting the DOCX document.
- `json`: For reading CV data from the JSON file.

## Running the Script
To run the script, execute the following command in your terminal or command prompt:

```python
python generate_CV.py cv_data.json
```


Replace `cv_data.json` with the path to your JSON file containing the CV data.

## Output
The script generates a DOCX file named `Mina_Ryad_CV.docx` containing the formatted CV based on the provided data.

## Author
This script was developed by Mina Ryad.

## Version History
- Version 1.0.0 (2/4/2024): Initial release.
