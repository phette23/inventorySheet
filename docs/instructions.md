# FOLIO Inventory Sheet Usage Instructions

Pulls physical item lists directly from FOLIO in call-number shelving order so library staff can run shelf inventories. This app is **read-only**—it does not modify records. It does not save your password or login token.

## 1. First-time setup

1. Make your own copy of the Google Sheet (**File → Make a copy**).
2. Reload the spreadsheet and wait a few seconds for the **FOLIO Inventory** menu to appear in the top menu bar.
3. Click **FOLIO Inventory → Setup config sheet**.
4. The first time you run a menu action, Google will ask you to authorize the Apps Script. Grant permission, then click **FOLIO Inventory → Setup config sheet** again if the `config` tab was not created yet.

You can safely re-run **Setup config sheet** if needed.

## 2. Configure the `config` tab

Fill in values in **column B** next to each setting label in column A (labels are case-insensitive).

### Required settings

Find your API URL and Tenant ID under the Settings app > System information (all the way at the bottom) > Look in the **API gateway services** section where bullet points say "For tenant" and "On URL".

| Setting | Example |
| --- | --- |
| **API URL** | `https://api-scelc.folio.ebsco.com` |
| **Tenant ID** | `cs00000003m001` |
| **Username** | `jdoe` or `staff@school.edu` |

### Optional filter & export settings

Leave optional settings blank to skip them. Multiple filters combine with `AND`.

| Setting | Example | What it does |
| --- | --- | --- |
| **Call number start** | `QA` | Lower bound on `effectiveShelvingOrder` (case-insensitive). |
| **Call number end** | `QC` | Inclusive upper bound on `effectiveShelvingOrder` (includes all items through the end of that prefix, e.g., all `QC` call numbers up to `QD`). |
| **Status filter** | `Available` or `Available, Missing` | Exact item status match. Separate multiple statuses with commas. |
| **Note contains** | `2026-11` | Matches items whose item notes or administrative notes contain the given text. |
| **Max rows** | `2,500` | Maximum items to fetch in one run (defaults to `2000`). Helps avoid Google Apps Script's 6-minute execution limit. |
| **CQL query** | `effectiveLocation.name=="Main Stacks"` | Optional advanced FOLIO CQL filter combined with the settings above. Results sort by `effectiveShelvingOrder` unless you include an explicit `sortby` clause. |

## 3. Choose output columns (`Return fields`)

Below the settings on the `config` tab is a collapsed **Return fields** section. Click the **`+`** button in the row margin on the left to expand it and toggle columns beyond the default inventory set (such as `Barcode` and `Call number`; `Barcode` is always included) to add optional fields (such as `Material type` or `Contributor names`).

### Shared print retention fields (`Holdings action note` & `OCN`)

- **Holdings action note** pulls any holdings note whose note type is named `Action note` (such as shared print retention commitments).
- **OCN** pulls the parent instance's `OCLC` identifier(s).

Leave both unchecked for routine inventories to keep runs fast. Checking either option triggers extra batched lookups against parent holdings and/or instance records.

## 4. Run an inventory export

1. Click **FOLIO Inventory → Create inventory**.
2. Enter your FOLIO password in the dialog and click **Create inventory**.
3. Wait for the status message in the dialog:
   - **Created `<tab name>` with N records:** A new tab has been added and opened, with row 1 frozen and rows sorted in call-number order. Tab names summarize active filters and the run date (e.g., `QA-QC Available 2026-11-25`, or `QA-QC Available 2026-11-25 (2)` if repeated on the same day).
   - **Stopped at max rows limit:** More items matched than `Max rows` allowed. Narrow your call number range (e.g., split `QA–QZ` into `QA–QH` and `QJ–QZ`) or raise `Max rows` carefully.
   - **No records matched your query:** Nothing happens.

## 5. Common recipes

- **Expected shelf list for a call number range:** Set `Call number start` = `QA`, `Call number end` = `QC`, and `Status filter` = `Available`.
- **Single location shelf list:** Add `effectiveLocation.name=="Main Stacks"` to `CQL query` alongside your call number range and `Status filter`.
- **Check items with a recent inventory note:** Set `Note contains` = `2026-11` (or your library's date/note marker) to pull matching items, or pull all `Available` items in the range and filter the `Item notes` / `Administrative notes` column in Google Sheets.
- **Shared print retention check:** Expand **Return fields**, check **Holdings action note** and/or **OCN**, and run your range/status query.

## 6. Required FOLIO permissions

The FOLIO account used in the sheet only needs **read-only** inventory permissions:

- In the FOLIO UI, **Inventory: View instances, holdings, and items** is sufficient for all features.
- For minimal API permission sets:
  - Standard exports: `inventory.items.collection.get`
  - `Holdings action note`: also `inventory-storage.holdings.collection.get` and `inventory-storage.holdings-note-types.collection.get`
  - `OCN`: also `inventory-storage.holdings.collection.get`, `inventory-storage.instances.collection.get`, and `inventory-storage.identifier-types.collection.get`

**We strongly recommend creating an account solely for the inventory sheet with only the necessary permissions above.**

## 7. Troubleshooting

- **HTTP 401 / 403 error:** Check your password, `Tenant ID`, and whether the FOLIO account has sufficient permissions for the endpoints above.
- **Entry `"Action note"` or `"OCLC"` was not found:** The enrichment feature matches those exact type names (case-insensitive) in FOLIO Settings → Inventory → Holdings note types / Resource identifier types. Making this more flexible is a potential feature request.
- **Script times out or hits `Max rows`:** Google Apps Script has a 6-minute execution limit. Break large collections into smaller call number ranges, and keep `Holdings action note` and `OCN` unchecked unless needed.
- **Call number sub-ranges look wider or narrower than expected:** Prefix ranges like `QA`–`QC` match `effectiveShelvingOrder` directly; numeric sub-ranges depend on how your FOLIO tenant normalizes shelving order. Include `Effective shelving order` in **Return fields** to see the exact values FOLIO sorts on.
