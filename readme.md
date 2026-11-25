# FOLIO inventory sheet

Exports FOLIO instances to a new Google Sheets tab.

## Setup

In **Extensions → Apps Script**, copy `app.js` into a script file, create an HTML file named `password` from `password.html`, and replace the manifest with `appsscript.json` (enable manifest visibility in Project Settings). Save and reload the spreadsheet.

Create a `config` sheet with keys in column A and values in column B. Keys are case insensitive.

| Key | Value |
| --- | --- |
| api url | https://your-folio-gateway.example.org |
| tenant id | your-tenant |
| username | your-username |

Choose **FOLIO Inventory → Create inventory** and enter your password. The account needs inventory instance read permission. Authentication uses `/authn/login-with-expiry`; passwords and tokens are not persisted. Instances are fetched from `/inventory/instances` in pages of 100.

Read the [full usage instructions here](./docs/instructions.md).

## Output

Each run creates a sheet with one instance per row and discovered JSON leaf paths as headers, including fields first encountered in later records. Objects use dotted paths (`metadata.createdDate`); arrays use indexes (`contributors.0.name`). Null and missing values are blank; empty objects and arrays add no columns.

Leave `INSTANCE_COLUMNS` in `app.js` empty for all fields, or select, rename, and order fields with entries such as `{ header: 'Author', path: 'contributors.0.name' }`.

All records are fetched before writing. Large exports can exceed Apps Script execution/memory or Sheets cell limits; values over 50,000 characters fail. Indexed arrays can create many columns. Exports do not resume, refresh tokens, or snapshot records during pagination.

## Development

Node.js ≥22.13.0 and pnpm.

```sh
pnpm install
pnpm lint # ESLint: Apps Script, embedded HTML JavaScript, config, and tests
pnpm test # Node test runner
pnpm test:live # Live FOLIO integration tests (uses .env or FOLIO_* env vars)
pnpm validate # Lint and tests
```

ESLint globals cover Apps Script in `app.js` and `.gs` files, browser/`google` in HTML, and Node in config/tests. Optional semicolons warn.
