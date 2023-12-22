# JSON to Excel Converter

This Javascript script converts JSON data to an Excel file. It recursively processes nested JSON structures and creates separate sheets in the Excel file. Each sheet represents a level of nesting or an array of objects within the JSON data.

## Table of Contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [File Structure](#file-structure)
- [FAQ](#faq)
- [Contributing](#contributing)

## Requirements

- Node.js installed on your machine
- Input JSON file with the data to be converted

## Installation

1. Clone the repository or download the code.
2. Open a terminal and navigate to the project directory.
3. Run `npm install` to install the required dependencies.

## Usage

1. Replace the placeholder JSON file (`input.json`) with your actual JSON data.
2. Open a terminal and navigate to the project directory.
3. Run the conversion script:

   ```bash
   node script.js
   ```

4. The script will process the JSON data and create an Excel file (`output_file.xlsx`) in the same directory.

## File Structure

- `convert-json-to-excel.js`: The main script file containing the JSON-to-Excel conversion logic.
- `input.json`: Placeholder JSON file (replace it with your data).
- `output_file.xlsx`: The generated Excel file.
- `node_modules/`: Directory containing Node.js modules and dependencies.
- `package.json` and `package-lock.json`: Node.js project configuration files.

## FAQ

### How does the script handle nested JSON structures?

The `processJSON` function recursively processes nested JSON structures, creating separate sheets for each level of nesting. Arrays of objects are treated as separate sheets, and other variables are added to the current sheet.

### Can the script handle large JSON files efficiently?

The script is designed to handle large JSON files, but for optimal performance, consider implementing streaming JSON parsing and asynchronous file reading.

### Why was the hyperlink feature removed from the original code?

The initial implementation attempted to create hyperlinks to each sheet, but due to issues, it was simplified to just display the sheet name without creating actual hyperlinks.

## Contributing

Feel free to open issues or submit pull requests to improve the script or address any issues.

