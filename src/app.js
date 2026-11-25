const ITEM_PAGE_SIZE = 100
const BATCH_ID_CHUNK_SIZE = 40
const DEFAULT_MAX_ROWS = 2000
const SCELC_CENTRAL_TENANT = 'cs00000003'
const HOLDINGS_ACTION_NOTE_FIELD = 'Holdings action note'
const OCN_FIELD = 'OCN'
// shortcut: matches "Action note" and "OCLC" by exact name; make configurable if a tenant uses custom type names
const HOLDINGS_ACTION_NOTE_TYPE = 'Action note'
const OCN_IDENTIFIER_TYPE = 'OCLC'
// Config tab settings keyed by config property name.
// REM: cast sheet strings to lowercase before comparing with setting labels
const CONFIG_SETTINGS = {
  url: { label: 'API URL', required: true, defaultValue: '' },
  tenant: { label: 'Tenant ID', required: true, defaultValue: '' },
  username: { label: 'Username', required: true, defaultValue: '' },
  callNumberStart: { label: 'Call number start', required: false, defaultValue: '' },
  callNumberEnd: { label: 'Call number end', required: false, defaultValue: '' },
  statusFilter: { label: 'Status filter', required: false, defaultValue: '' },
  noteContains: { label: 'Note contains', required: false, defaultValue: '' },
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
  new Field('Effective shelving order',  false,  false,  item => item.effectiveShelvingOrder),
  new Field(HOLDINGS_ACTION_NOTE_FIELD,  false,  false,  item => item._holdingsActionNote),
  new Field(OCN_FIELD,                   false,  false,  item => item._ocn)
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

  const enrichedItems = enrichItems(config, token, items)
  const sheet = writeItems(enrichedItems, config.fields, buildSheetName(config))

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
    callNumberStart: values.callNumberStart || CONFIG_SETTINGS.callNumberStart.defaultValue,
    callNumberEnd: values.callNumberEnd || CONFIG_SETTINGS.callNumberEnd.defaultValue,
    statusFilter: values.statusFilter || CONFIG_SETTINGS.statusFilter.defaultValue,
    noteContains: values.noteContains || CONFIG_SETTINGS.noteContains.defaultValue,
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

// Perform a GET request and parse the JSON response body.
function httpJson(config, path, token, label) {
  const response = http(config, path, token)
  try {
    return JSON.parse(response.getContentText())
  } catch (err) {
    throw new Error(`FOLIO ${label} response was not valid JSON.`, { cause: err })
  }
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

// Escape characters with special meaning inside CQL double-quoted strings.
function escapeCql(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"')
}

// Split a trailing unquoted `sortby ...` clause from a CQL expression so it stays outside parentheses.
function splitSortBy(cql) {
  let inQuotes = false
  for (let i = 0; i < cql.length; i++) {
    if (cql[i] === '"' && cql[i - 1] !== '\\') {
      inQuotes = !inQuotes
    } else if (!inQuotes && (i === 0 || /\s/.test(cql[i - 1]))) {
      const match = cql.slice(i).match(/^sortby\s+(.+)$/i)
      if (match) {
        return { filter: cql.slice(0, i).trim(), sortBy: match[1].trim() }
      }
    }
  }
  return { filter: cql.trim(), sortBy: '' }
}

// Translate plain-language config filters into CQL clauses for /inventory/items.
function buildFilterClauses(config = {}) {
  const clauses = []

  if (config.callNumberStart) {
    const start = escapeCql(String(config.callNumberStart).trim().toUpperCase())
    if (start) clauses.push(`effectiveShelvingOrder>="${start}"`)
  }

  if (config.callNumberEnd) {
    // Append "~" (ASCII 126) so an upper bound like "QC" includes all "QC ..." items up to "QD".
    const end = escapeCql(String(config.callNumberEnd).trim().toUpperCase())
    if (end) clauses.push(`effectiveShelvingOrder<="${end}~"`)
  }

  if (config.statusFilter) {
    const statuses = String(config.statusFilter)
      .split(',')
      .map(status => status.trim())
      .filter(Boolean)
    if (statuses.length === 1) {
      clauses.push(`status.name=="${escapeCql(statuses[0])}"`)
    } else if (statuses.length > 1) {
      clauses.push(statuses.map(status => `status.name=="${escapeCql(status)}"`).join(' or '))
    }
  }

  if (config.noteContains) {
    const term = escapeCql(String(config.noteContains).trim())
    if (term) clauses.push(`notes=="*${term}*" or administrativeNotes=="*${term}*"`)
  }

  if (config.cqlQuery) {
    clauses.push(config.cqlQuery)
  }

  return clauses
}

// Compose one or more CQL filter clauses with an optional sortby clause.
function buildCqlQuery(filter, sortBy = 'effectiveShelvingOrder') {
  const rawClauses = (Array.isArray(filter) ? filter : [filter])
    .map(clause => String(clause || '').trim())
    .filter(Boolean)

  const filters = []
  let explicitSort = ''

  rawClauses.forEach(clause => {
    const parts = splitSortBy(clause)
    if (parts.filter) filters.push(parts.filter)
    if (parts.sortBy) explicitSort = parts.sortBy
  })

  let query = filters.length === 0
    ? 'cql.allRecords=1'
    : filters.length === 1
      ? filters[0]
      : filters.map(clause => `(${clause})`).join(' and ')

  const effectiveSort = explicitSort || sortBy
  if (effectiveSort) {
    query += ` sortby ${effectiveSort}`
  }

  return query
}

// FOLIO API request to /inventory/items
function getItems(config, token) {
  const maxRows = config.maxRows || DEFAULT_MAX_ROWS
  const query = buildCqlQuery(buildFilterClauses(config))
  const items = []
  let offset = 0
  let totalRecords = null

  while (items.length < maxRows) {
    const remaining = maxRows - items.length
    const pageSize = Math.min(ITEM_PAGE_SIZE, remaining)
    const path = `/inventory/items?limit=${pageSize}&offset=${offset}&query=${encodeURIComponent(query)}`
    const body = httpJson(config, path, token, 'items')

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

// Resolve a FOLIO reference-data UUID by case-insensitive name (e.g. "Action note", "OCLC").
function lookupTypeId(config, token, endpoint, arrayKey, typeName) {
  const body = httpJson(config, `${endpoint}?limit=1000`, token, arrayKey)
  if (!body || !Array.isArray(body[arrayKey])) {
    throw new Error(`Expected a ${arrayKey} array from FOLIO.`)
  }
  const target = typeName.toLocaleLowerCase()
  const match = body[arrayKey].find(entry => entry && String(entry.name || '').trim().toLocaleLowerCase() === target)
  if (!match || !match.id) {
    throw new Error(`FOLIO ${arrayKey} entry "${typeName}" was not found.`)
  }
  return match.id
}

// Batch-fetch FOLIO storage records by UUID in URL-safe CQL chunks.
function fetchRecordsByIds(config, token, endpoint, arrayKey, ids) {
  const uniqueIds = [...new Set(ids.filter(Boolean))]
  const byId = new Map()

  for (let i = 0; i < uniqueIds.length; i += BATCH_ID_CHUNK_SIZE) {
    const chunk = uniqueIds.slice(i, i + BATCH_ID_CHUNK_SIZE)
    const query = `id==(${chunk.map(id => `"${escapeCql(id)}"`).join(' or ')})`
    const path = `${endpoint}?limit=${chunk.length}&query=${encodeURIComponent(query)}`
    const body = httpJson(config, path, token, arrayKey)
    if (!body || !Array.isArray(body[arrayKey])) {
      throw new Error(`Expected a ${arrayKey} array from FOLIO.`)
    }
    body[arrayKey].forEach(record => {
      if (record && record.id) byId.set(record.id, record)
    })
  }

  return byId
}

// Enrich items with parent holdings action notes and/or instance OCNs only when selected.
function enrichItems(config, token, items) {
  const selected = new Set((config.fields || []).map(field => field.name))
  const wantActionNote = selected.has(HOLDINGS_ACTION_NOTE_FIELD)
  const wantOcn = selected.has(OCN_FIELD)
  if (!items.length || (!wantActionNote && !wantOcn)) return items

  const actionNoteTypeId = wantActionNote
    ? lookupTypeId(config, token, '/holdings-note-types', 'holdingsNoteTypes', HOLDINGS_ACTION_NOTE_TYPE)
    : null
  const ocnTypeId = wantOcn
    ? lookupTypeId(config, token, '/identifier-types', 'identifierTypes', OCN_IDENTIFIER_TYPE)
    : null

  // /inventory/items does not reliably include instanceId, so fetch parent holdings whenever
  // action notes are requested or an item lacks instanceId for OCN lookup.
  const holdingsIds = wantActionNote
    ? items.map(item => item.holdingsRecordId)
    : items.filter(item => !item.instanceId).map(item => item.holdingsRecordId)
  const holdingsById = fetchRecordsByIds(
    config, token, '/holdings-storage/holdings', 'holdingsRecords', holdingsIds
  )

  let instancesById = new Map()
  if (wantOcn) {
    const instanceIds = items.map(item => {
      const holdings = holdingsById.get(item.holdingsRecordId)
      return item.instanceId || (holdings && holdings.instanceId)
    })
    instancesById = fetchRecordsByIds(
      config, token, '/instance-storage/instances', 'instances', instanceIds
    )
  }

  return items.map(item => {
    const enriched = Object.assign({}, item)
    const holdings = holdingsById.get(item.holdingsRecordId)

    if (wantActionNote) {
      const notes = holdings && Array.isArray(holdings.notes)
        ? holdings.notes.filter(note => note && note.holdingsNoteTypeId === actionNoteTypeId)
        : []
      enriched._holdingsActionNote = collectNotes(notes)
    }

    if (wantOcn) {
      const instanceId = item.instanceId || (holdings && holdings.instanceId)
      const instance = instancesById.get(instanceId)
      const ocns = instance && Array.isArray(instance.identifiers)
        ? instance.identifiers.filter(id => id && id.identifierTypeId === ocnTypeId)
        : []
      enriched._ocn = collect(ocns, 'value')
    }

    return enriched
  })
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

function formatIsoDate(date) {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

// Build a descriptive tab name from active filters and the run date (e.g. "QA-QC Available 2026-11-23").
function buildSheetName(config = {}, date = new Date()) {
  const parts = []
  const start = String(config.callNumberStart || '').trim().toUpperCase()
  const end = String(config.callNumberEnd || '').trim().toUpperCase()

  if (start && end) {
    parts.push(start === end ? start : `${start}-${end}`)
  } else if (start) {
    parts.push(`${start}+`)
  } else if (end) {
    parts.push(`to ${end}`)
  }

  if (config.statusFilter) {
    parts.push(String(config.statusFilter).trim())
  }

  if (!parts.length) parts.push('Inventory')
  parts.push(formatIsoDate(date))

  // Google Sheets forbids : \ / ? * [ ] in tab names and caps names at 100 chars.
  return parts.join(' ').replace(/[:\\/?*[\]]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90)
}

// Append " (2)", " (3)", ... when a sheet with the same name already exists.
function uniqueSheetName(spreadsheet, baseName) {
  if (!baseName || typeof spreadsheet.getSheetByName !== 'function') return baseName
  if (!spreadsheet.getSheetByName(baseName)) return baseName
  let suffix = 2
  while (spreadsheet.getSheetByName(`${baseName} (${suffix})`)) suffix += 1
  return `${baseName} (${suffix})`
}

// Write FOLIO item records onto a new sheet using the selected field descriptors.
function writeItems(items, fields = getDefaultFields(), sheetName = '') {
  const columns = fields.length ? fields : getDefaultFields()
  const rows = [columns.map(field => cellValue(field.name))]

  items.forEach(item => {
    rows.push(columns.map(field => cellValue(field.getValue(item))))
  })

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const name = uniqueSheetName(spreadsheet, sheetName)
  const sheet = name ? spreadsheet.insertSheet(name) : spreadsheet.insertSheet()

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
