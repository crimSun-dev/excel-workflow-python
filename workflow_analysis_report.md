# Workflow Analysis Report: Financial Data Processing in Excel

This report details the workflow and pipeline demonstrated in the YouTube video ["0mv89NcqKhI"](https://youtu.be/0mv89NcqKhI), which focuses on processing, enriching, and summarizing raw financial data using Microsoft Excel. The analysis breaks down each step, identifying the tools used, inputs, processes, and outputs.

## Workflow Steps

### **Step 1: Data Import and Column Separation**

This initial step involves bringing raw data into Excel and structuring it into separate columns.

*   **Tool Used:** Text to Columns (found under the Data Tab in Excel).
*   **Inputs:** Raw financial data, typically pasted from a `.txt` file, residing in a single column (e.g., Column A) within an Excel spreadsheet. The data fields are delimited by a pipe symbol (`|`).
*   **Process:**
    1.  The user selects the column containing the raw, delimited data.
    2.  Navigates to the **Data** tab in the Excel ribbon and clicks on the **Text to Columns** feature.
    3.  In the Text to Columns wizard, **Delimited** is chosen as the data type.
    4.  The delimiter is specified as **Other**, and the pipe symbol (`|`) is entered into the corresponding field.
*   **Output:** The single column of raw data is successfully parsed and distributed into multiple distinct columns, each representing a specific data field.

### **Step 2: Data Enrichment with VLOOKUP**

This step enhances the dataset by incorporating additional information from a reference source using a lookup function.

*   **Tool Used:** VLOOKUP function in Excel.
*   **Inputs:** The `KODE_UKER` column from the currently processed sheet, which contains codes to be looked up. A separate Excel workbook serves as a reference table, containing mapping data for bank branches (e.g., `KODE_UKER` to `MAIN_CODE` and `MAIN_BRANCH`).
*   **Process:**
    1.  Two new columns are inserted into the main dataset, typically named `MAIN_CODE` and `MAIN_BRANCH`.
    2.  The `VLOOKUP` function is applied to these new columns. For `MAIN_CODE`, the function searches for the `KODE_UKER` value in the reference workbook and retrieves the corresponding data from the 3rd column of the reference table.
    3.  Similarly, for `MAIN_BRANCH`, the `VLOOKUP` function retrieves data from the 4th column of the reference table based on the `KODE_UKER` value.
*   **Output:** The original dataset is enriched with standardized bank branch codes (`MAIN_CODE`) and their respective names (`MAIN_BRANCH`), providing more context to the financial data.

### **Step 3: Data Summarization with Pivot Tables**

This step aggregates the enriched data to provide a high-level summary.

*   **Tool Used:** PivotTable (found under the Insert Tab in Excel).
*   **Inputs:** The complete, enriched data table from the previous steps.
*   **Process:**
    1.  The user navigates to the **Insert** tab in Excel and selects **PivotTable**.
    2.  The PivotTable fields are configured as follows:
        *   **Filters:** The `SEGMEN` field is added to allow filtering of specific segments (e.g., wholesale, corporate) if required for analysis.
        *   **Rows:** The `MAIN_CODE` and `MAIN_BRANCH` fields are added to define the rows of the summary table, grouping data by bank branch.
        *   **Values:** The `Sum of VOLUME_IN_IDR` field is added to calculate the total volume in Indonesian Rupiah for each branch.
*   **Output:** A preliminary summary table is generated, showing aggregated financial volumes per bank branch.

### **Step 4: Pivot Table Refinement and Formatting**

This final step involves enhancing the readability and presentation of the PivotTable.

*   **Tools Used:** PivotTable Design settings and Number Formatting options in Excel.
*   **Inputs:** The newly created PivotTable from Step 3.
*   **Process:**
    1.  **Layout Adjustment:** In the **Design** tab of the PivotTable Tools, the **Report Layout** is changed to **Show in Tabular Form** for a clearer, more structured display.
    2.  **Subtotal Removal:** To streamline the report, the user right-clicks on the `MAIN_CODE` field within the PivotTable and unchecks **Subtotal "MAIN_CODE"**, removing redundant subtotals.
    3.  **Number Formatting:** The values column (e.g., `Sum of VOLUME_IN_IDR`) is selected, and its number format is changed from **General** to **Number**. This action ensures that large numerical values are displayed fully, eliminating scientific notation (e.g., `E+12`) and improving readability.
*   **Final Output:** A clean, well-formatted, and easily readable summary report. This report clearly presents the total volume in IDR for each bank branch, organized by their main codes, ready for further analysis or presentation.

---

**Author:** Manus AI
**Date:** July 21, 2026
