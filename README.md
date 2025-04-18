# HYSYS Mass & Energy Balance Formatter
This is a lightweight Python utility that takes raw data from Aspen HYSYS simulations and formats it into a standardized mass and energy balance spreadsheet. The formatted output follows the structure required by my then client, Koura Global, for documentation, internal review, or vendor handover.

## Features
* Parses raw Excel data exported from HYSYS
* Extracts key parameters including Phase, Temperature, Pressure, Enthalpy, and Density
* Formats the data into a clean, readable Excel spreadsheet
* Applies styling such as auto-width columns and centered text
* Saves output dedicated ```Outputs/``` folder

## Usage
1) Open Aspen HYSYS and navigate to the results section.
   
2) Expand all sections and subsections so that all relevant stream data is visible (e.g. not just "Mole Flows", but each component underneath).

3) Export this data to Excel and save the file as:
   ```Data/Raw_Data_Input.xlsx```

4) Run the script:
   ```python3 SRC/ME_Balance_Programme_Code.py```

5) The formatted output will be saved as:
   ```Outputs/Results/xlsx```

## Folder Structure
```
HYSYS-ME-Balance-Formatting/
├── SRC/
│   └── ME_Balance_Programme_Code.py     # Main Python script
├── Data/
│   └── Raw_Data_Input.xlsx              # Raw HYSYS data file
├── Outputs/
│   └── Results.xlsx                     # Formatted spreadsheet output
├── Requirements.txt                     # Python dependencies
└── .gitignore                           # Files/folders to exclude from Git tracking
```

## Requirements
* Python 3.8+
* Required libraries (install via pip):
  - pandas
  - numpy
  - openpyxl
  - jinja2

## Installation
1) clone or download this repository:
   ```
   git clone https://github.com/Mo-Somji/HYSYS-ME-Balance-Formatting.git
   cd HYSYS-ME-Balance-Formatting
   ```
2) Install the required packages:
   ```
   pip install -r Requirements.txt
   ```

## Notes
* Ensure the exported Excel structure matches what the script expects (raw stream data with key headers like 'Phase', 'Temperature', etc)
* Use ```python3``` explicitly on macOS if ```python``` defaults to Python 2
* If styling fails, ensure ```jinja2``` is installed (used by pandas' Styler)














    
