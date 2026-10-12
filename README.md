# Order planning report

This local Streamlit application converts the daily ERP Order Status workbook into the validated first-36-eligible-orders Excel report. It does not call ChatGPT or any AI API.

## Start on this computer

Double-click `start_app.bat`. Streamlit opens the application in the default browser. Keep the command window open while using the app; close it to stop the app.

## First-time setup on another Windows computer

1. Install Python 3.12 or newer.
2. Double-click `setup_app.bat` while connected to the internet.
3. After setup finishes, double-click `start_app.bat`.

## Daily use

1. Upload the daily `.xlsx` Order Status workbook.
2. Choose **Plan to Ship** (source sheet order) or **Buyer Requested Shipment Date** (earliest first).
3. Choose the complete workbook or one condition-specific output.
4. Select **Generate report**.
5. Review the validation summary.
6. Download the generated Excel workbook.

The source workbook is read in memory and is never edited.

## Eligibility rules

For Plan to Ship, the app selects the first 36 eligible orders in source sheet order. For Buyer Requested Shipment Date, it sorts whole orders by that date ascending (keeping source order for ties), then selects the first 36 eligible orders. It excludes a finished-good line when its `Prod Name` contains `STD`, begins with `FG`, or exactly matches an `Item_Name` in `data/exemptions/finished_good_exemptions.xlsx`. All comparisons are case-insensitive. If every FG line in an order is excluded, the order is skipped and the next eligible order takes its sequence number. For a mixed order, only the excluded FG lines are removed. An eligible order needed for the first 36 in Buyer Requested Shipment Date mode must have a valid date.

To update exemptions, replace `data/exemptions/finished_good_exemptions.xlsx` with the latest list. Keep an `Item_Name` column; no app upload is needed.
