# Daily Cashbook & Denomination Tracker 🚀

A premium, state-of-the-art web application for managing daily cash transactions, bank entries, and physical cash verification.

## ✨ Features
- **Unified Dashboard**: Single interface for Cashiers and Admins.
- **Role-Based Access**: Secure login with Admin and Cashier roles.
- **Daily Ledger**: Track Receipts (Deposits) and Payments (Expenses).
- **Bank Ledger**: Manage Bank Credits and Debits.
- **Cash Verification**: Denomination counter with automated variance calculation.
- **Global Search**: Deep-scan all historic records for any keyword.
- **Advanced Reporting**: Interactive charts, export to Excel (XLSX), PDF, or Copy to Clipboard.
- **Verification System**: Lock daily records after verification to prevent tampering.
- **Theme Support**: Dynamic Light and Dark modes.
- **Cloud/Offline Sync**: Works with Google Sheets as a backend or in standalone offline mode.

## 🛠️ Quick Setup (Google Sheets)

1.  **Create a Google Sheet**: Open [sheets.new](https://sheets.new).
2.  **Open Script Editor**: Go to `Extensions` > `Apps Script`.
3.  **Paste Code**: Copy the contents of `Code.gs` into the script editor.
4.  **Create HTML File**: In the script editor, click `+` > `HTML`, name it `Daily_Cashbook`, and paste the contents of `Daily Cashbook.html`.
5.  **Deploy**: 
    - Click `Deploy` > `New Deployment`.
    - Select type `Web App`.
    - Set `Execute as` to `Me`.
    - Set `Who has access` to `Anyone`.
    - Click `Deploy` and copy the **Web App URL**.
6.  **Connect App**:
    - Open the Web App URL.
    - Go to `Admin Portal` (Login with `admin` / `123`).
    - Paste the Web App URL into the "System Configuration" box and click `Update & Connect`.
7.  **Initialize**: Run the `setupSheets` function once from the Apps Script editor to create all necessary sheets.

## 🔒 Default Credentials
- **Username**: `admin`
- **Password**: `123`
*(Change these immediately in the Admin Portal)*

## 📂 Project Structure
- `Code.gs`: Server-side logic for Google Apps Script.
- `Daily Cashbook.html`: Frontend UI, CSS, and Client-side logic.
- `README.md`: This documentation.

---
Built with ❤️ for professional financial management.
