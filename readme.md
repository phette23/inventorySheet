# FOLIO inventory sheet

Exports FOLIO physical items (`/inventory/items`) in shelving order to a new Google Sheets tab. Read-only; passwords and access tokens are never stored.

Read the [full librarian usage instructions](./docs/instructions.md).

## Setup

In **Extensions → Apps Script**, copy `src/app.js` into a script file, create an HTML file named `password` from `src/password.html`, and replace the manifest with `src/appsscript.json` (enable manifest visibility in Project Settings). Save and reload the spreadsheet.

Choose **FOLIO Inventory → Setup config sheet** to create (or normalize) the `config` tab with all settings and a collapsed **Return fields** checkbox list. Fill in the required values in column B (keys in column A are case-insensitive):

| Key | Value |
| --- | --- |
| API URL | `https://your-folio-gateway.example.org` |
| Tenant ID | `your-tenant` |
| Username | `your-username` |

Optional settings (`Call number start`, `Call number end`, `Status filter`, `Note contains`, `Max rows`, `CQL query`) and output column checkboxes can be adjusted on the same tab.

Choose **FOLIO Inventory → Create inventory** and enter your FOLIO password. Authentication uses `/authn/login-with-expiry`. Usernames ending in `@scelc.org` automatically authenticate against the SCELC central tenant (`cs00000003`) while querying the configured `Tenant ID`; member accounts use `Tenant ID` for both.

## FOLIO permissions

The script is strictly read-only. In the FOLIO UI, standard **Inventory: View instances, holdings, and items** permissions cover everything. For minimal API permission sets:

- **Standard runs:** `inventory.items.collection.get` (`/inventory/items`)
- **With `Holdings action note` checked:** also `inventory-storage.holdings.collection.get` and `inventory-storage.holdings-note-types.collection.get`
- **With `OCN` checked:** also `inventory-storage.holdings.collection.get` (when items omit `instanceId`), `inventory-storage.instances.collection.get`, and `inventory-storage.identifier-types.collection.get`

**We strongly recommend creating an account solely for the inventory sheet with only the necessary permissions above.**

## Output & limits

Each run creates a descriptively named tab (for example, `QA-QC Available 2026-11-25`) with one item per row, sorted by `effectiveShelvingOrder` unless your `CQL query` specifies another `sortby`. If zero items match, no blank tab is created. Values starting with `=`, `+`, `-`, or `@` are escaped so Sheets does not evaluate them as formulas.

By default, 8 inventory columns are included (`Item UUID`, `Barcode`, `Title`, `Call number`, `Location`, `Item status`, `Item notes`, `Administrative notes`), and 17 optional fields—including `Holdings action note` and `OCN`—can be toggled via checkboxes on `config`. Enrichment requests only run when `Holdings action note` or `OCN` is checked.

All records are fetched before writing, capped at `Max rows` (default `2000`) to stay within Apps Script execution and memory limits. Individual cells over 50,000 characters fail. Exports do not resume, refresh tokens, or snapshot records during pagination.

## Development

Node.js ≥22.13.0 and pnpm.

```sh
pnpm install
pnpm lint # ESLint: Apps Script, embedded HTML JavaScript, config, and tests
pnpm test # Node unit tests (offline)
pnpm test:live # Live FOLIO integration tests (uses .env or FOLIO_* env vars)
pnpm validate # Lint and unit tests
```

ESLint globals cover Apps Script in `src/app.js` and `.gs` files, browser/`google` in HTML, and Node in config/tests. Optional semicolons warn.
