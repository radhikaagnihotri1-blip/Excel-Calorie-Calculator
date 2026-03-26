/**
 * ============================================================
 * Health Tracker - Google Apps Script
 * ============================================================
 * This script auto-fills nutritional data (calories, protein,
 * carbs, fat) for each ingredient using the Open Food Facts API,
 * and maintains a per-dish summary sheet.
 *
 * Sheet columns expected in "Ingredients" sheet:
 *   A: Dish Name
 *   B: Ingredient Name
 *   C: Quantity (grams)
 *   D: Calories       ← auto-filled
 *   E: Protein (g)    ← auto-filled
 *   F: Carbs (g)      ← auto-filled
 *   G: Fat (g)        ← auto-filled
 * ============================================================
 */

// ── Configuration constants ────────────────────────────────
var INGREDIENTS_SHEET_NAME = "Ingredients";
var SUMMARY_SHEET_NAME     = "Summary";
var DATA_START_ROW         = 2;   // Row 1 is the header row

var COL_DISH       = 1;  // A
var COL_INGREDIENT = 2;  // B
var COL_QUANTITY   = 3;  // C
var COL_CALORIES   = 4;  // D
var COL_PROTEIN    = 5;  // E
var COL_CARBS      = 6;  // F
var COL_FAT        = 7;  // G

var NOT_FOUND_COLOR = "#FFFF00"; // Yellow highlight for missing data
var FOUND_COLOR     = "#FFFFFF"; // White (clear highlight) for found data


// ── 1. Add a custom menu when the spreadsheet opens ───────
/**
 * onOpen() runs automatically every time the Google Sheet is opened.
 * It creates a custom menu called "Update Nutrition" in the menu bar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Update Nutrition")
    .addItem("Fill missing rows only", "fillMissingNutrition")
    .addItem("Refresh ALL rows (overwrite existing)", "refreshAllNutrition")
    .addItem("Rebuild Summary sheet", "buildSummarySheet")
    .addToUi();
}


// ── 2. onEdit trigger – auto-fill when a row is completed ─
/**
 * onEdit(e) runs automatically every time a cell is edited.
 * It checks if the edited row has Dish Name, Ingredient Name, AND
 * Quantity filled in, but nutrition columns are still empty.
 * If so, it fetches and fills nutrition for just that row.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();

  // Only react to edits on the Ingredients sheet
  if (sheet.getName() !== INGREDIENTS_SHEET_NAME) return;

  var row = e.range.getRow();

  // Ignore the header row
  if (row < DATA_START_ROW) return;

  var dishName   = sheet.getRange(row, COL_DISH).getValue();
  var ingredient = sheet.getRange(row, COL_INGREDIENT).getValue();
  var quantity   = sheet.getRange(row, COL_QUANTITY).getValue();
  var calories   = sheet.getRange(row, COL_CALORIES).getValue();

  // Only proceed if A, B, C are filled AND D (calories) is still empty
  if (dishName && ingredient && quantity && !calories) {
    fillRowNutrition(sheet, row, false);
    buildSummarySheet(); // Rebuild summary after each new entry
  }
}


// ── 3. Fill only rows that are missing nutrition data ─────
/**
 * fillMissingNutrition() is triggered by the custom menu.
 * It loops through every row that has A, B, C filled in but D is empty,
 * and fetches the nutrition data for each such row.
 */
function fillMissingNutrition() {
  var sheet = getOrCreateSheet(INGREDIENTS_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No data found. Please add ingredients first.");
    return;
  }

  var filledCount = 0;

  for (var row = DATA_START_ROW; row <= lastRow; row++) {
    var dish       = sheet.getRange(row, COL_DISH).getValue();
    var ingredient = sheet.getRange(row, COL_INGREDIENT).getValue();
    var quantity   = sheet.getRange(row, COL_QUANTITY).getValue();
    var calories   = sheet.getRange(row, COL_CALORIES).getValue();

    // Skip rows that are incomplete or already have data
    if (!dish || !ingredient || !quantity) continue;
    if (calories !== "" && calories !== null) continue;

    var filled = fillRowNutrition(sheet, row, false);
    if (filled) filledCount++;

    // Pause briefly to avoid hitting API rate limits
    Utilities.sleep(300);
  }

  buildSummarySheet();
  SpreadsheetApp.getUi().alert("Done! Filled nutrition for " + filledCount + " row(s).");
}


// ── 4. Refresh ALL rows (overwrite existing data) ─────────
/**
 * refreshAllNutrition() is triggered by the custom menu.
 * It re-fetches nutrition data for EVERY row that has A, B, C filled,
 * even if D-G already have values. Use this to refresh stale data.
 */
function refreshAllNutrition() {
  var sheet = getOrCreateSheet(INGREDIENTS_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No data found. Please add ingredients first.");
    return;
  }

  var filledCount = 0;

  for (var row = DATA_START_ROW; row <= lastRow; row++) {
    var dish       = sheet.getRange(row, COL_DISH).getValue();
    var ingredient = sheet.getRange(row, COL_INGREDIENT).getValue();
    var quantity   = sheet.getRange(row, COL_QUANTITY).getValue();

    if (!dish || !ingredient || !quantity) continue;

    var filled = fillRowNutrition(sheet, row, true); // true = overwrite
    if (filled) filledCount++;

    Utilities.sleep(300);
  }

  buildSummarySheet();
  SpreadsheetApp.getUi().alert("Done! Refreshed nutrition for " + filledCount + " row(s).");
}


// ── 5. Core function: fetch & fill one row ────────────────
/**
 * fillRowNutrition(sheet, row, overwrite)
 *
 * Fetches nutritional data for the ingredient in the given row
 * using the Open Food Facts API, then writes the calculated values
 * into columns D-G.
 *
 * @param {Sheet}   sheet     - The Ingredients sheet object
 * @param {number}  row       - The row number to process
 * @param {boolean} overwrite - If true, overwrite existing D-G values
 * @returns {boolean} true if data was successfully written
 */
function fillRowNutrition(sheet, row, overwrite) {
  var ingredient = sheet.getRange(row, COL_INGREDIENT).getValue().toString().trim();
  var quantity   = parseFloat(sheet.getRange(row, COL_QUANTITY).getValue());

  if (!ingredient || isNaN(quantity) || quantity <= 0) return false;

  // Skip if data already exists and we are not overwriting
  if (!overwrite) {
    var existing = sheet.getRange(row, COL_CALORIES).getValue();
    if (existing !== "" && existing !== null) return false;
  }

  // Call the Open Food Facts API
  var nutrition = fetchNutrition(ingredient);

  if (!nutrition) {
    // Ingredient not found – highlight the row yellow and leave cells blank
    sheet.getRange(row, COL_CALORIES, 1, 4).clearContent();
    sheet.getRange(row, COL_DISH, 1, 7)
         .setBackground(NOT_FOUND_COLOR);
    Logger.log("Not found: " + ingredient);
    return false;
  }

  // Calculate actual values based on quantity entered
  // API returns values per 100g, so we scale by (quantity / 100)
  var factor   = quantity / 100;
  var calories = roundTo2(nutrition.calories * factor);
  var protein  = roundTo2(nutrition.protein  * factor);
  var carbs    = roundTo2(nutrition.carbs    * factor);
  var fat      = roundTo2(nutrition.fat      * factor);

  // Write values into columns D, E, F, G
  sheet.getRange(row, COL_CALORIES).setValue(calories);
  sheet.getRange(row, COL_PROTEIN).setValue(protein);
  sheet.getRange(row, COL_CARBS).setValue(carbs);
  sheet.getRange(row, COL_FAT).setValue(fat);

  // Clear any previous yellow highlight (ingredient was found this time)
  sheet.getRange(row, COL_DISH, 1, 7).setBackground(FOUND_COLOR);

  return true;
}


// ── 6. Open Food Facts API call ───────────────────────────
/**
 * fetchNutrition(ingredientName)
 *
 * Searches the Open Food Facts API for the given ingredient name.
 * Returns an object { calories, protein, carbs, fat } per 100g,
 * or null if nothing useful is found.
 *
 * API docs: https://world.openfoodfacts.org/data
 * No API key is required – it is completely free and open.
 *
 * @param  {string} ingredientName - e.g. "onion"
 * @returns {object|null}
 */
function fetchNutrition(ingredientName) {
  try {
    // Build the search URL. We ask for 5 results and pick the best one.
    var encodedName = encodeURIComponent(ingredientName);
    var url = "https://world.openfoodfacts.org/cgi/search.pl"
            + "?search_terms=" + encodedName
            + "&search_simple=1"
            + "&action=process"
            + "&json=1"
            + "&page_size=5"
            + "&fields=product_name,nutriments";

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    // If the API call itself failed, return null
    if (response.getResponseCode() !== 200) {
      Logger.log("API error for: " + ingredientName + " | HTTP " + response.getResponseCode());
      return null;
    }

    var data = JSON.parse(response.getContentText());

    // Walk through the returned products and find the first one
    // that has all four nutritional fields populated
    if (data && data.products && data.products.length > 0) {
      for (var i = 0; i < data.products.length; i++) {
        var product    = data.products[i];
        var nutriments = product.nutriments;

        if (!nutriments) continue;

        // Open Food Facts uses these standard field names (per 100g)
        var cal     = nutriments["energy-kcal_100g"];
        var protein = nutriments["proteins_100g"];
        var carbs   = nutriments["carbohydrates_100g"];
        var fat     = nutriments["fat_100g"];

        // Accept this product only if at least calories and one macro are present
        if (cal != null && (protein != null || carbs != null || fat != null)) {
          return {
            calories: parseFloat(cal)     || 0,
            protein:  parseFloat(protein) || 0,
            carbs:    parseFloat(carbs)   || 0,
            fat:      parseFloat(fat)     || 0
          };
        }
      }
    }

    // No suitable product found
    return null;

  } catch (err) {
    Logger.log("Exception fetching nutrition for '" + ingredientName + "': " + err.message);
    return null;
  }
}


// ── 7. Build / rebuild the Summary sheet ─────────────────
/**
 * buildSummarySheet()
 *
 * Reads all ingredient rows from the Ingredients sheet, groups them
 * by Dish Name, sums up the nutrition values, and writes the results
 * to a separate "Summary" sheet tab.
 *
 * Creates the Summary sheet if it doesn't already exist.
 */
function buildSummarySheet() {
  var ingredSheet = getOrCreateSheet(INGREDIENTS_SHEET_NAME);
  var summarySheet = getOrCreateSheet(SUMMARY_SHEET_NAME);

  var lastRow = ingredSheet.getLastRow();

  // Clear old summary content (keep the sheet, just wipe its data)
  summarySheet.clearContents();
  summarySheet.clearFormats();

  // ── Write summary header ──
  var headers = ["Dish Name", "Total Calories", "Total Protein (g)", "Total Carbs (g)", "Total Fat (g)"];
  var headerRange = summarySheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4A90D9");
  headerRange.setFontColor("#FFFFFF");

  if (lastRow < DATA_START_ROW) return; // Nothing to summarise

  // ── Aggregate nutrition data per dish ──
  // We use a plain object as a map: { dishName: { cal, protein, carbs, fat } }
  var totals = {};

  for (var row = DATA_START_ROW; row <= lastRow; row++) {
    var dish     = ingredSheet.getRange(row, COL_DISH).getValue().toString().trim();
    var calories = parseFloat(ingredSheet.getRange(row, COL_CALORIES).getValue()) || 0;
    var protein  = parseFloat(ingredSheet.getRange(row, COL_PROTEIN).getValue())  || 0;
    var carbs    = parseFloat(ingredSheet.getRange(row, COL_CARBS).getValue())    || 0;
    var fat      = parseFloat(ingredSheet.getRange(row, COL_FAT).getValue())      || 0;

    if (!dish) continue; // Skip blank rows

    if (!totals[dish]) {
      totals[dish] = { calories: 0, protein: 0, carbs: 0, fat: 0 };
    }

    totals[dish].calories += calories;
    totals[dish].protein  += protein;
    totals[dish].carbs    += carbs;
    totals[dish].fat      += fat;
  }

  // ── Write aggregated rows to Summary sheet ──
  var summaryRow = 2;
  for (var dishName in totals) {
    var t = totals[dishName];
    summarySheet.getRange(summaryRow, 1).setValue(dishName);
    summarySheet.getRange(summaryRow, 2).setValue(roundTo2(t.calories));
    summarySheet.getRange(summaryRow, 3).setValue(roundTo2(t.protein));
    summarySheet.getRange(summaryRow, 4).setValue(roundTo2(t.carbs));
    summarySheet.getRange(summaryRow, 5).setValue(roundTo2(t.fat));

    // Alternate row colours for readability
    var rowColor = (summaryRow % 2 === 0) ? "#EAF4FF" : "#FFFFFF";
    summarySheet.getRange(summaryRow, 1, 1, 5).setBackground(rowColor);

    summaryRow++;
  }

  // Auto-resize columns for neatness
  summarySheet.autoResizeColumns(1, 5);
}


// ── 8. Ensure the header row exists on first run ──────────
/**
 * setupHeaders()
 *
 * Writes the column headers for the Ingredients sheet if they are missing.
 * You can run this once manually from the Apps Script editor to initialise
 * a brand-new sheet.
 */
function setupHeaders() {
  var sheet = getOrCreateSheet(INGREDIENTS_SHEET_NAME);

  var headers = [
    "Dish Name",
    "Ingredient Name",
    "Quantity (g)",
    "Calories",
    "Protein (g)",
    "Carbs (g)",
    "Fat (g)"
  ];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#34A853");
  headerRange.setFontColor("#FFFFFF");

  sheet.autoResizeColumns(1, 7);
  SpreadsheetApp.getUi().alert("Headers set up successfully on the Ingredients sheet!");
}


// ── Helper: get or create a sheet by name ────────────────
/**
 * getOrCreateSheet(name)
 * Returns the sheet with the given name, or creates it if it doesn't exist.
 *
 * @param  {string} name - Sheet tab name
 * @returns {Sheet}
 */
function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}


// ── Helper: round a number to 2 decimal places ───────────
/**
 * roundTo2(value)
 * @param  {number} value
 * @returns {number}
 */
function roundTo2(value) {
  return Math.round(value * 100) / 100;
}
