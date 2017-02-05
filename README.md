# HtmlLocalizer

## What is it?
A simple tool which localizes a HTML document using an Excel as data source. It contains:
- **example**: A folder with 2 Excel files (en-US and pt-BR), also a simple HTML template files
- **src**: The source code of the project

## How to use?
1. Select the folder which contains the Excel files
2. Select the template you want translated
3. Push the "Localize" button

## How does it work?
The tool will:

1. Loop through all Excel files located in the folder you selected
2. Get the keys and values from each the Excel
3. Load the template file
4. Replace all template keys by their corresponding values
5. Save the newly created localized files in the _result_ folder, which will be created automatically in the same place the _template_ folder is located

Feel free to download, use and contribute.
