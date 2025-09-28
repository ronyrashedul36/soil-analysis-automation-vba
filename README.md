# Automated Soil Nutrient Analysis Tool for Excel

An advanced VBA-powered tool designed to automate the tedious process of analyzing and summarizing soil nutrient data. This project transforms a raw dataset of soil sample results into professionally formatted, percentage-based summary reports with a single click, saving significant time and reducing manual error.

## The Problem
Agronomists, researchers, and farmers often deal with large spreadsheets containing dozens of soil sample results. Manually categorizing each data point for parameters like pH, Nitrogen, and Potassium is time-consuming, repetitive, and prone to errors. This tool was built to solve that exact problem.

## Key Features
- **Fully Automated Analysis**: Processes raw data for 14 distinct soil parameters across thousands of rows in seconds.
- **Dynamic Summary Reports**: Automatically generates clean summary tables showing the total count and percentage for each classification level (e.g., 'Very Low', 'Low', 'Medium', 'High', 'Very High').
- **User-Friendly Interface**: Managed by simple "Analyze" and "Clear" buttons directly on the worksheet—no coding knowledge required for end-users.
- **Modular and Scalable Code**: Built with a data-driven design, allowing developers to easily add new soil parameters or adjust classification thresholds in a single configuration block without rewriting logic.
- **Error-Resilient**: The script gracefully handles non-numeric or empty cells, ensuring the analysis runs smoothly.

---

## Technology Stack
- **Microsoft Excel**
- **Visual Basic for Applications (VBA)**

---

## Getting Started

Follow these steps to get the tool up and running on your local machine.

### Prerequisites
- Microsoft Excel 2007 or newer.
- Your soil sample data ready in an Excel sheet.

### Installation & Usage
1.  **Download:** Clone this repository or download the `your-file-name.xlsm` file.
2.  **Enable Macros:** Open the `.xlsm` file. You will likely see a security warning at the top. Click **"Enable Content"** or **"Enable Macros"**. This is required for the tool to function.
3.  **Prepare Your Data:**
    -   Your raw data must be on a sheet named **`nabinagar`**.
    -   Data should start from row 2 (assuming row 1 is for headers).
    -   Ensure the data is organized in the correct columns as follows:
        -   **Column H**: Acidity (pH)
        -   **Column J**: Organic Matter
        -   **Column K**: Nitrogen
        -   **Column L**: Phosphorus(o)
        -   ... *(add other columns as needed)*
4.  **Run the Analysis:** Click the **"Analyze Data"** button on the worksheet. The summary tables will appear, starting in cell `Z13`.
5.  **Clear the Output:** Click the **"Clear Output"** button to remove the generated tables and prepare for a new analysis.

---

## Customization

This tool was designed to be easily adaptable. If you need to add a new soil parameter to analyze or change the thresholds for an existing one, you only need to edit the `GetAnalysisConfigurations` function in the VBA module.

**Example:** To add a "Sodium" analysis for data in Column W:

1.  Open the VBA Editor (`Alt + F11`).
2.  Navigate to the module containing the `RoundedRectangle2_Click` function.
3.  Add a new `AllAnalyses(14)` block with the relevant details:

```vba
' 14: Sodium (Column W) - NEW
With AllAnalyses(14)
    .Name = "Sodium Status"
    .ColumnLetter = "W"
    .OutputStartCell = "AH13" ' Choose an empty location
    .Thresholds = Array(10, 25, 50) ' e.g., Low, Medium, High
    .Labels = Array("Low (<=10)", "Medium (11-25)", "High (26-50)", "Very High (>50)")
End With
