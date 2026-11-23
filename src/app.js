const ITEM_PAGE_SIZE = 100
const DEFAULT_MAX_ROWS = 2000
const SCELC_CENTRAL_TENANT = 'cs00000003'
// Config tab settings keyed by config property name.
// REM: cast sheet strings to lowercase before comparing with setting labels
const CONFIG_SETTINGS = {
  url: { label: 'API URL', required: true, defaultValue: '' },
  tenant: { label: 'Tenant ID', required: true, defaultValue: '' },
  username: { label: 'Username', required: true, defaultValue: '' },
  maxRows: { label: 'Max rows', required: false, defaultValue: DEFAULT_MAX_ROWS },
  cqlQuery: { label: 'CQL query', required: false, defaultValue: '' }
}
const FIELDS_SECTION_HEADER = 'Return fields'

// Field descriptor modeled on Caltech Library's Boffo project.
function Field(name, enabled, required, getValue) {
  this.name = name
  this.enabled = enabled
  this.required = required
  this.getValue = getValue
}

// Concatenate a string property across an array of objects.
function collect(list, subfield, separator = '; ') {
  if (!Array.isArray(list)) return ''
  return list
    .map(entry => (entry && entry[subfield] ? String(entry[subfield]).trim() : ''))
    .filter(Boolean)
    .join(separator)
}

// Concatenate an array of plain strings (e.g. administrativeNotes).
function collectStrings(list, separator = '\n') {
  if (!Array.isArray(list)) return ''
  return list
    .map(entry => (entry == null ? '' : String(entry).trim()))
    .filter(Boolean)
    .join(separator)
}

// Concatenate FOLIO note objects ({ note }) across separate lines.
function collectNotes(list) {
  return collect(list, 'note', '\n')
}

// Assemble a full display call number from FOLIO's effective call number components and item parts.
function formatCallNumber(item) {
  const parts = item.effectiveCallNumberComponents || {}
  const base = parts.callNumber || item.callNumber || ''
  return [
    parts.prefix,
    base,
    parts.suffix,
    item.volume,
    item.enumeration,
    item.chronology,
    item.copyNumber
  ]
    .map(part => (part == null ? '' : String(part).trim()))
    .filter(Boolean)
    .join(' ')
}

// FOLIO /inventory/items fields available as checkboxes on the config sheet.
const FIELDS = [
  //         Name                       Enabled Required getValue(item)
  new Field('Item UUID',                 true,   false,  item => item.id),
  new Field('Barcode',                   true,   true,   item => item.barcode),
  new Field('Title',                     true,   false,  item => item.title),
  new Field('Call number',               true,   false,  item => formatCallNumber(item)),
  new Field('Location',                  true,   false,  item => item.effectiveLocation && item.effectiveLocation.name),
  new Field('Item status',               true,   false,  item => item.status && item.status.name),
  new Field('Item notes',                true,   false,  item => collectNotes(item.notes)),
  new Field('Administrative notes',      true,   false,  item => collectStrings(item.administrativeNotes)),
  new Field('Contributor names',         false,  false,  item => collect(item.contributorNames, 'name')),
  new Field('Circulation notes',         false,  false,  item => collectNotes(item.circulationNotes)),
  new Field('Material type',             false,  false,  item => item.materialType && item.materialType.name),
  new Field('Permanent loan type',       false,  false,  item => item.permanentLoanType && item.permanentLoanType.name),
  new Field('Temporary loan type',       false,  false,  item => item.temporaryLoanType && item.temporaryLoanType.name),
  new Field('Permanent location',        false,  false,  item => item.permanentLocation && item.permanentLocation.name),
  new Field('Temporary location',        false,  false,  item => item.temporaryLocation && item.temporaryLocation.name),
  new Field('Copy number',               false,  false,  item => item.copyNumber),
  new Field('Volume',                    false,  false,  item => item.volume),
  new Field('Enumeration',               false,  false,  item => item.enumeration),
  new Field('Chronology',                false,  false,  item => item.chronology),
  new Field('Item HRID',                 false,  false,  item => item.hrid),
  new Field('Holdings UUID',             false,  false,  item => item.holdingsRecordId),
  new Field('Status date',               false,  false,  item => item.status && item.status.date),
  new Field('Effective shelving order',  false,  false,  item => item.effectiveShelvingOrder)
]

/* exported onOpen, showPasswordDialog, createInventory, setupConfigSheet */

// On opening the sheet, add a "FOLIO Inventory" menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('FOLIO Inventory')
    .addItem('Create inventory', 'showPasswordDialog')
    .addItem('Setup config sheet', 'setupConfigSheet')
    .addToUi()
}

// Show password.html, triggered by "Create inventory" menu button
function showPasswordDialog() {
  getConfig()
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('password').setWidth(380).setHeight(240),
    'Create FOLIO inventory'
  )
}

// Create or update the "config" tab with core settings and a collapsed field checkbox group.
function setupConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spreadsheet.getSheetByName('config')
  if (!sheet) sheet = spreadsheet.insertSheet('config')

  const existingKeys = new Map()
  const lastRow = sheet.getLastRow()

  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, 2).getDisplayValues().forEach((row, index) => {
      const key = String(row[0]).trim().toLocaleLowerCase()
      if (key && !existingKeys.has(key)) existingKeys.set(key, index + 1)
    })
  }

  let nextRow = lastRow + 1

  Object.values(CONFIG_SETTINGS).forEach(({ label, defaultValue }) => {
    const key = label.toLocaleLowerCase()
    if (!existingKeys.has(key)) {
      sheet.getRange(nextRow, 1, 1, 2).setValues([[label, defaultValue]])
      existingKeys.set(key, nextRow)
      nextRow += 1
    }
  })

  const headerKey = FIELDS_SECTION_HEADER.toLocaleLowerCase()
  if (!existingKeys.has(headerKey)) {
    const headerRange = sheet.getRange(nextRow, 1, 1, 2)
    headerRange.setValues([[FIELDS_SECTION_HEADER, 'Include']])
    if (typeof headerRange.setFontWeight === 'function') headerRange.setFontWeight('bold')
    existingKeys.set(headerKey, nextRow)
    nextRow += 1
  }

  const addedFieldStart = nextRow

  FIELDS.forEach(field => {
    const key = field.name.toLocaleLowerCase()
    if (!existingKeys.has(key)) {
      const rowRange = sheet.getRange(nextRow, 1, 1, 2)
      rowRange.setValues([[field.name, field.enabled]])
      const checkboxCell = sheet.getRange(nextRow, 2, 1, 1)
      if (typeof checkboxCell.insertCheckboxes === 'function') checkboxCell.insertCheckboxes()
      existingKeys.set(key, nextRow)
      nextRow += 1
    }
  })

  const addedFieldCount = nextRow - addedFieldStart
  if (addedFieldCount > 0) {
    const groupRange = sheet.getRange(addedFieldStart, 1, addedFieldCount, 1)
    if (typeof groupRange.shiftRowGroupDepth === 'function') {
      groupRange.shiftRowGroupDepth(1)
    }
    if (typeof sheet.collapseAllRowGroups === 'function') {
      sheet.collapseAllRowGroups()
    }
  }

  return spreadsheet.setActiveSheet(sheet)
}

// Main function, referenced in password.html
function createInventory(password) {
  if (typeof password !== 'string' || !password) throw new Error('Enter your FOLIO password.')

  const config = getConfig()
  const token = authenticate(config, password)
  const { items, totalRecords, truncated } = getItems(config, token)

  if (!items.length) {
    return { sheetName: null, count: 0, totalRecords, truncated: false }
  }

  const sheet = writeItems(items, config.fields)

  return {
    sheetName: sheet.getName(),
    count: items.length,
    totalRecords,
    truncated
  }
}

// Parse Google Sheets checkbox display values ("TRUE"/"FALSE") or raw booleans.
function parseCheckbox(value) {
  if (typeof value === 'boolean') return value
  const normalized = String(value || '').trim().toLocaleLowerCase()
  if (normalized === 'true') return true
  if (normalized === 'false') return false
  return null
}

// Pull configuration values and optional field checkboxes from a sheet named "config"
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config')
  const entries = Object.entries(CONFIG_SETTINGS)
  const requiredSettings = entries.filter(([, setting]) => setting.required)
  if (!sheet) {
    const names = requiredSettings.map(([, setting]) => setting.label).join(', ')
    throw new Error(`Create a sheet named "config" with ${names} in column A and their values in column B.`)
  }

  const settingKeyByLabel = new Map(
    entries.map(([key, setting]) => [setting.label.toLocaleLowerCase(), key])
  )
  const fieldsByKey = new Map(FIELDS.map(field => [field.name.toLocaleLowerCase(), field]))
  const values = {}
  const fieldOverrides = {}

  sheet.getDataRange().getDisplayValues().forEach(row => {
    const label = String(row[0]).trim().toLocaleLowerCase()
    if (settingKeyByLabel.has(label)) {
      values[settingKeyByLabel.get(label)] = String(row[1] || '').trim()
    } else if (fieldsByKey.has(label)) {
      const checked = parseCheckbox(row[1])
      if (checked !== null) fieldOverrides[fieldsByKey.get(label).name] = checked
    }
  })

  const missing = requiredSettings
    .filter(([key]) => !values[key])
    .map(([, setting]) => setting.label)
  if (missing.length) throw new Error(`Missing config values: ${missing.join(', ')}`)

  let maxRows = CONFIG_SETTINGS.maxRows.defaultValue
  if (values.maxRows) {
    // Strip thousands separators that Google Sheets display formatting may add.
    const parsed = Number(values.maxRows.replace(/,/g, ''))
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new Error(`Config value "${CONFIG_SETTINGS.maxRows.label}" must be a positive whole number.`)
    }
    maxRows = parsed
  }

  const fields = FIELDS.filter(field => {
    if (field.required) return true
    if (Object.prototype.hasOwnProperty.call(fieldOverrides, field.name)) {
      return fieldOverrides[field.name]
    }
    return field.enabled
  })

  return {
    url: values.url.replace(/\/+$/, ''),
    tenant: values.tenant,
    username: values.username,
    maxRows,
    cqlQuery: values.cqlQuery || CONFIG_SETTINGS.cqlQuery.defaultValue,
    fields
  }
}

// HTTP request with FOLIO authentication
function http(config, path, token, body, tenant = config.tenant) {
  const headers = { 'x-okapi-tenant': tenant, Accept: 'application/json' }
  if (token) headers['x-okapi-token'] = token
  const options = {
    headers,
    followRedirects: false,
    method: body ? 'post' : 'get',
    muteHttpExceptions: true
  }

  if (body) {
    options.contentType = 'application/json'
    options.payload = JSON.stringify(body)
  }

  let response
  try {
    response = UrlFetchApp.fetch(config.url + path, options)
  } catch (err) {
    throw new Error('FOLIO request failed. Check the gateway URL and network access.', { cause: err })
  }

  const status = response.getResponseCode()
  if (status < 200 || status >= 300) {
    Logger.log(`HTTP error. Response headers: ${response.getHeaders()}\nResponse body: ${response.getContentText()}`)
    throw new Error(`FOLIO returned HTTP ${status}. Check credentials, permissions, and endpoint availability.`)
  }

  return response
}

// Trade username & password for an access token.
// SCELC Office accounts (@scelc.org) authenticate against the central tenant before querying member data.
function authenticate(config, password) {
  const authTenant = /@scelc\.org$/i.test(config.username) ? SCELC_CENTRAL_TENANT : config.tenant
  const response = http(config, '/authn/login-with-expiry', null, {
    username: config.username, password
  }, authTenant)
  let body = {}

  try {
    body = JSON.parse(response.getContentText())
  } catch (_) { /* Cookie-only login. */}

  if (body && typeof body.accessToken === 'string' && body.accessToken) return body.accessToken

  const headers = response.getAllHeaders()
  for (const name of Object.keys(headers)) {
    if (name.toLowerCase() !== 'set-cookie') continue
    const cookies = Array.isArray(headers[name]) ? headers[name] : [headers[name]]
    for (const cookie of cookies) {
      const match = String(cookie).match(/(?:^|,\s*)folioAccessToken=([^;,\s]+)/)
      if (match) return match[1]
    }
  }

  throw new Error('FOLIO login returned no access token. Check the authentication endpoint for your deployment.')
}

// Compose one or more CQL filter clauses with an optional sortby clause.
function buildCqlQuery(filter, sortBy = 'id') {
  const clauses = (Array.isArray(filter) ? filter : [filter])
    .map(clause => String(clause || '').trim())
    .filter(Boolean)

  let query = clauses.length === 0
    ? 'cql.allRecords=1'
    : clauses.length === 1
      ? clauses[0]
      : clauses.map(clause => `(${clause})`).join(' and ')

  // Ignore quoted text when checking if the caller already specified a sortby clause.
  const unquoted = query.replace(/"(?:\\.|[^"])*"/g, '')
  if (sortBy && !/\bsortby\b/i.test(unquoted)) {
    query += ` sortby ${sortBy}`
  }

  return query
}

// FOLIO API request to /inventory/items
function getItems(config, token) {
  const maxRows = config.maxRows || DEFAULT_MAX_ROWS
  const query = buildCqlQuery(config.cqlQuery, 'id')
  const items = []
  let offset = 0
  let totalRecords = null

  while (items.length < maxRows) {
    const remaining = maxRows - items.length
    const pageSize = Math.min(ITEM_PAGE_SIZE, remaining)
    const path = `/inventory/items?limit=${pageSize}&offset=${offset}&query=${encodeURIComponent(query)}`
    const response = http(config, path, token)

    let body
    try {
      body = JSON.parse(response.getContentText())
    } catch (err) {
      throw new Error('FOLIO items response was not valid JSON.', { cause: err })
    }

    if (!body || !Array.isArray(body.items)) throw new Error('Expected an items array from FOLIO.')

    if (Number.isInteger(body.totalRecords)) totalRecords = body.totalRecords

    const page = body.items.slice(0, remaining)
    items.push(...page)
    offset += page.length

    if (body.items.length < pageSize ||
      (Number.isInteger(totalRecords) && offset >= totalRecords)) {
      return { items, totalRecords, truncated: false }
    }
  }

  // Flag truncation when more records exist (or totalRecords is absent after hitting the cap).
  const truncated = Number.isInteger(totalRecords) ? totalRecords > items.length : true
  return { items, totalRecords, truncated }
}

// Ensure values function safely in Google Sheets before writing them
function cellValue(value) {
  if (value === undefined || value === null) return ''
  const cell = value
  if (typeof cell !== 'string') return cell
  if (cell.length > 50000) throw new Error('An output value exceeds the Google Sheets 50,000-character cell limit.')
  // Prevent API strings from being evaluated as spreadsheet formulas.
  return /^[=+\-@]/.test(cell) ? "'" + cell : cell
}

function getDefaultFields() {
  return FIELDS.filter(field => field.required || field.enabled)
}

// Write FOLIO item records onto a new sheet using the selected field descriptors.
function writeItems(items, fields = getDefaultFields()) {
  const columns = fields.length ? fields : getDefaultFields()
  const rows = [columns.map(field => cellValue(field.name))]

  items.forEach(item => {
    rows.push(columns.map(field => cellValue(field.getValue(item))))
  })

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.insertSheet() // Automatically assigns a unique name.

  // Ensure the sheet's dimensions are large enough for the data
  if (rows.length > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows.length - sheet.getMaxRows())
  }
  if (rows[0].length > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), rows[0].length - sheet.getMaxColumns())
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows)
  sheet.setFrozenRows(1)
  return spreadsheet.setActiveSheet(sheet)
}
