# Automated E-Commerce Profit & Expense Tracker (with VBA Automation)
![Dashboard Screenshot](<img width="1229" height="692" alt="dashboard1" src="https://github.com/user-attachments/assets/9053a73a-704a-43e6-9bd4-55cb413c3080" />
)

A powerful, fully automated financial bookkeeping and data visualization tool designed for e-commerce sellers, marketing agencies, and dropshippers. This project bridges the gap between manual data entry and complex accounting software by enabling **1-Click Data Imports** via Excel VBA and offering an interactive executive dashboard.

## 🚀 Key Features

*   **Dynamic Executive Dashboard:** Automatically tracks and visualizes Key Performance Indicators (KPIs) including Total Revenue, Total Expenses, Net Profit, and Profit Margin.
*   **1-Click VBA Data Importer:** Eliminates manual data entry fatigue. Users can click a custom button to instantly upload and append raw transaction logs from `.csv` or `.xlsx` export files directly into the data core.
*   **Automated Expense Breakdown:** Real-time categorical tracking of major operational overheads (Product Cost, Facebook Ads, TikTok Ads, Shipping Fees) mapped dynamically via robust `SUMIFS` engineering.
*   **Visual Analytics:** Clean, client-ready interactive column charts that auto-refresh as new transaction data is imported.

## 📁 Repository Structure

*   `Automated E-Commerce Profit & Expense Tracker_V2.xlsm` - The core application (Macro-Enabled Excel Workbook containing formulas, dashboard UI, and the VBA module).
*   `Ecom_Test_Data_2026.csv` - Dummy transaction payload file provided for testing the 1-click import functionality.
*   `README.md` - Documentation and project guide.

## 🛠️ How to Use

1.  Clone or download this repository and open the `.xlsm` file in Microsoft Excel.
2.  **Important:** Make sure to click **"Enable Macros"** or **"Enable Content"** at the top bar when Excel opens, otherwise the automated import button will not function.
3.  Navigate to the `Dashboard` sheet.
4.  Click the custom **"Click to Import Data"** button.
5.  Select the provided `Ecom_Test_Data_2026.csv` file from your local machine.
6.  Watch the `Data_Log` update instantly and the entire dashboard/charts refresh automatically!

## 💻 Technical Stack Behind the Project
*   **Advanced Spreadsheet Architecture:** Structured Reference Tables, Dynamic Named Ranges, Data Validation.
*   **Formulas Utilized:** `SUMIFS`, `IFERROR`, nested logic handling data type variations.
*   **Automation:** Visual Basic for Applications (VBA) utilizing `FileDialog` wrappers for smooth system OS file picking.
