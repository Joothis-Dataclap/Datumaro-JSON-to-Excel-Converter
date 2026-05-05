# Datumaro JSON to Excel Converter

A Streamlit application that converts Datumaro format JSON annotation files to structured Excel spreadsheets.

## Features

✨ **Dynamic Processing**
- Upload any Datumaro format JSON file
- Automatically extracts all labels from annotations
- No hardcoding needed - works with any set of labels

📊 **Professional Excel Output**
- Serial number, image name, frame number columns
- Dynamic label columns with Box and Polygon sub-columns
- Summary row with total counts for each label
- Professional formatting with colors and borders

🎯 **Clear Streamlit Display**
- Data preview table
- Statistics dashboard
- Annotation breakdown by label
- Download button for Excel file
- Auto-saves to output folder

## Installation

```bash
# Install dependencies
pip install streamlit pandas openpyxl
```

## Usage

```bash
# Run the Streamlit app
streamlit run app.py
```

Then:
1. Open the Streamlit interface (usually at http://localhost:8501)
2. Upload a Datumaro format JSON file
3. View the preview and statistics
4. Download the Excel file

## Datumaro JSON Format

The application expects JSON in this format:

```json
{
  "items": [
    {
      "id": "1",
      "image": {
        "path": "image_001.jpg"
      },
      "annotations": [
        {
          "type": "bbox",
          "label": "Wall",
          "x": 10,
          "y": 20,
          "width": 100,
          "height": 150
        },
        {
          "type": "polygon",
          "label": "Window",
          "points": [[10, 10], [50, 10], [50, 50], [10, 50]]
        }
      ]
    }
  ]
}
```

## Excel Output Structure

| SI No | Image Name | Frame Number | Label1 - Box | Label1 - Polygon | Label2 - Box | Label2 - Polygon | ... | Total Annotations |
|-------|------------|--------------|------|----------|------|----------|-----|-------------------|
| 1 | image_001.jpg | 1 | 2 | 1 | 1 | 0 | ... | 4 |
| 2 | image_002.jpg | 2 | 1 | 2 | 0 | 1 | ... | 4 |

**Summary Row** (at bottom):
Shows total counts for each label and overall total.

## Test Data

A sample Datumaro JSON file is included: `sample_datumaro.json`

You can use this to test the application and see how it works with multiple images and labels.

## Notes

- All Excel files are automatically saved to the `output/` folder
- The application supports multiple annotation types: bbox, polygon
- Labels are dynamically extracted from the JSON
- Summary statistics are calculated automatically
