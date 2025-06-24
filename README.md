# ScanParcel Google Sheets Add‑on

This repository contains a Google Apps Script used to scan parcel numbers and update
a Google Sheets document. The script adds a custom menu and a sidebar that allows
you to scan or manually input parcel codes. It also integrates with Shopify to
cancel orders when parcels are returned.

## Deploying to Google Sheets

1. Open the Google Sheet where you want to use the scanner.
2. Choose **Extensions → Apps Script** to open the Apps Script editor.
3. Copy the contents of `code.gs` and `ScannerSidebar.html` from this repository
   into new files of the same names in the Apps Script project.
4. Save the project. When you reload the spreadsheet you will see a **Scanner**
   menu with items **Open Scanner Sidebar** and **Reconcile COD Payments**.

## Configuring Shopify credentials

To enable automatic order cancellation on Shopify you must set two script
properties:

- `SHOP_TOKEN` – a private app access token.
- `SHOP_DOMAIN` – your shop domain, e.g. `example.myshopify.com`.

To create these properties:

1. In the Apps Script editor, choose **Project Settings** (gear icon).
2. Under **Script Properties**, click **Add script property**.
3. Enter `SHOP_TOKEN` as the key and paste your Shopify token as the value.
4. Add another property named `SHOP_DOMAIN` with your shop domain.
5. Save the changes.

The script reads these properties when making API calls. If they are missing, the
Shopify features will be skipped.

## Using the Sidebar

Select **Scanner → Open Scanner Sidebar** from the spreadsheet menu. The sidebar
contains a single input box and four buttons:

- **Scan** – records the parcel as dispatched or prompts to mark it returned.
- The scan will also warn if another undelivered order has the same customer
  name or phone number so you can choose whether to dispatch it.
- **Undo Last Scan** – reverts the most recent scan and adjusts the summary
  sheets.
- **Cancel Order** – marks an order as "Cancelled by Customer" using the parcel
  number and attempts to cancel the corresponding order on Shopify.
- **Cancel by Order #** – enter just the numeric portion of an order number to
  cancel it in the sheet and on Shopify.
- Once cancelled by the customer, an order cannot be dispatched.
- **Set Custom Status** – choose a status like "Dispatched" or "Returned" and optionally select a date to apply it.
  - **Reconcile COD Payments** – cross-checks the `TCS Invoice` sheet and marks dispatched orders as "Paid ✅" or "Dispatched – No COD ❌".
  After each action a short message appears at the bottom of the sidebar to confirm
  what happened.

## Importing TCS Invoice Data

Copy or import your invoice rows into the `TCS Invoice` sheet. The first row
should contain the headers `ParcelNo`, `CODAmount` and `Status`. Then choose
**Scanner → Reconcile COD Payments** to cross‑check dispatched orders.
The script marks each parcel in column **N** as either "Paid ✅" or
"Dispatched – No COD ❌". Delivered parcels listed in `TCS Invoice` that are
missing from `Sheet1` are highlighted in yellow so you can easily spot them.
Invoice rows that correspond to orders marked **Paid ✅** are highlighted in
light green.

## Dispatch Summary

The **Scanner** menu also provides a **Dispatch Summary** submenu. Choose one of the preset ranges (Last 5 Days, Last Week or Last Month) or pick **Custom Range…** to enter your own start and end dates. A sheet named **Dispatch Summary** will be generated with total quantities grouped by product for the selected period.

## Local Rider Orders

When the shipping status is set to **Dispatch through Local Rider**, the script now records each order ID and amount in a sheet named **Local Rider Orders**. This sheet lists every local rider dispatch on a daily basis so you can easily reconcile cash payments collected by riders.

