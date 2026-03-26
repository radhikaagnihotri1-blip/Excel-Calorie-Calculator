# Health Tracker – Setup Instructions

Step-by-step guide to paste and run the Google Apps Script inside your Google Sheet.

---

## Step 1 – Open your Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and open (or create) a spreadsheet named **Health Tracker**.

---

## Step 2 – Open the Apps Script editor

1. In the menu bar click **Extensions → Apps Script**.
2. A new browser tab opens showing the Apps Script editor.
3. You will see a default file called `Code.gs` with an empty function inside it.

---

## Step 3 – Paste the script

1. **Select all** the existing placeholder code in `Code.gs` (Ctrl+A / Cmd+A) and **delete** it.
2. Copy the entire contents of `HealthTracker.gs` from this repository.
3. Paste it into the editor.
4. Click the **floppy-disk icon** (or press Ctrl+S / Cmd+S) to save.
   Give the project a name like **Health Tracker Script** when prompted.

---

## Step 4 – Set up the sheet headers (first-time only)

1. In the Apps Script editor, find the function dropdown at the top (it probably says `onOpen` or `myFunction`).
2. Click the dropdown and select **`setupHeaders`**.
3. Click the **▶ Run** button.
4. A permissions dialog will appear – click **Review permissions → Allow**.
   *(This lets the script read/write your spreadsheet and call external URLs.)*
5. Switch back to your Google Sheet. You should now see a green header row on the **Ingredients** tab with columns:
   `Dish Name | Ingredient Name | Quantity (g) | Calories | Protein (g) | Carbs (g) | Fat (g)`

---

## Step 5 – Start entering data

Fill in the **Ingredients** sheet like this:

| A (Dish Name) | B (Ingredient Name) | C (Quantity g) | D | E | F | G |
|---|---|---|---|---|---|---|
| Dal Tadka | Onion | 150 | ← auto | ← auto | ← auto | ← auto |
| Dal Tadka | Tomato | 100 | ← auto | ← auto | ← auto | ← auto |
| Dal Tadka | Toor Dal | 80 | ← auto | ← auto | ← auto | ← auto |

As soon as you fill in all three of columns A, B, and C for a row, the script
automatically fetches nutrition data and fills D–G within a few seconds.

---

## Step 6 – Use the custom menu

After the sheet reloads (or after you run `onOpen` once from the editor), you will
see a new menu in the Google Sheets menu bar called **"Update Nutrition"** with three options:

| Menu item | What it does |
|---|---|
| **Fill missing rows only** | Fetches data only for rows where D–G are empty |
| **Refresh ALL rows (overwrite existing)** | Re-fetches everything, even if D–G already have values |
| **Rebuild Summary sheet** | Recalculates the per-dish totals on the Summary tab |

---

## Step 7 – Check the Summary sheet

A second sheet tab called **Summary** is created automatically.
It shows one row per dish with the total calories, protein, carbs, and fat for that dish.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Row is highlighted **yellow** | The ingredient name was not found in Open Food Facts. Try a simpler English name (e.g. "lentil" instead of "Toor Dal"). You can edit the name and run "Fill missing rows only" again. |
| Script doesn't run automatically on edit | Make sure you granted permissions in Step 4. You can also go to Apps Script → Triggers and confirm the `onEdit` trigger exists. |
| "You do not have permission" error | Re-run `setupHeaders` and go through the permissions dialog again. |
| API returns wrong nutrition values | Open Food Facts is crowd-sourced and may have inaccurate entries. Cross-check with a trusted source and edit values manually if needed. |

---

## How the API works (for reference)

The script calls the **Open Food Facts** API — a free, open-source food database.

- **No API key required.**
- It searches by ingredient name and picks the first result that has all four nutrition fields.
- Values are returned **per 100 g** and the script scales them by `(your quantity / 100)`.
- API endpoint used: `https://world.openfoodfacts.org/cgi/search.pl`
