// Add { header: 'Column label', path: 'nested.json.field' } entries here later.
// Numeric path segments select array entries. Leave empty to discover all JSON leaf fields.
const INSTANCE_COLUMNS = []
const INSTANCE_PAGE_SIZE = 100
const DEFAULT_MAX_ROWS = 2000
// Backend-only override for SCELC central tenant; leave empty for members
let AUTH_TENANT_OVERRIDE = ''
// let AUTH_TENANT_OVERRIDE = 'cs00000003'
// REM: cast sheet strings to lowercase before comparing with these
const REQUIRED_CONFIG_KEYS = ['api url', 'tenant id', 'username']
const OPTIONAL_CONFIG_KEYS = ['max rows', 'cql query']

/* exported onOpen, showPasswordDialog, createInventory */

// On opening the sheet, add a "FOLIO Inventory" menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('FOLIO Inventory')
    .addItem('Create inventory', 'showPasswordDialog')
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

// Main function, referenced in password.html
function createInventory(password) {
  if (typeof password !== 'string' || !password) throw new Error('Enter your FOLIO password.')

  const config = getConfig()
  const token = authenticate(config, password)
  const { instances, totalRecords, truncated } = getInstances(config, token)

  if (!instances.length) {
    return { sheetName: null, count: 0, totalRecords, truncated: false }
  }

  const sheet = writeInstances(instances)

  return {
    sheetName: sheet.getName(),
    count: instances.length,
    totalRecords,
    truncated
  }
}

// Pull configuration values from a sheet named "config"
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config')
  if (!sheet) throw new Error(`Create a sheet named "config" with ${REQUIRED_CONFIG_KEYS.join(', ')} in column A and their values in column B.`)

  const knownKeys = [...REQUIRED_CONFIG_KEYS, ...OPTIONAL_CONFIG_KEYS]
  const values = {}

  sheet.getDataRange().getDisplayValues().forEach(row => {
    const key = String(row[0]).trim().toLocaleLowerCase()
    if (knownKeys.includes(key)) values[key] = String(row[1] || '').trim()
  })

  const missing = REQUIRED_CONFIG_KEYS.filter(key => !values[key])
  if (missing.length) throw new Error(`Missing config values: ${missing.join(', ')}`)

  let maxRows = DEFAULT_MAX_ROWS
  if (values['max rows']) {
    // Strip thousands separators that Google Sheets display formatting may add.
    const parsed = Number(values['max rows'].replace(/,/g, ''))
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new Error('Config value "max rows" must be a positive whole number.')
    }
    maxRows = parsed
  }

  return {
    url: values['api url'].replace(/\/+$/, ''),
    tenant: values['tenant id'],
    username: values.username,
    maxRows,
    cqlQuery: values['cql query'] || ''
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

// Trade username & password for an access token
function authenticate(config, password, authTenant = AUTH_TENANT_OVERRIDE || config.tenant) {
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

// FOLIO API request to /inventory/instances
function getInstances(config, token) {
  const maxRows = config.maxRows || DEFAULT_MAX_ROWS
  const query = buildCqlQuery(config.cqlQuery, 'id')
  const instances = []
  let offset = 0
  let totalRecords = null

  while (instances.length < maxRows) {
    const remaining = maxRows - instances.length
    const pageSize = Math.min(INSTANCE_PAGE_SIZE, remaining)
    const path = `/inventory/instances?limit=${pageSize}&offset=${offset}&query=${encodeURIComponent(query)}`
    const response = http(config, path, token)

    let body
    try {
      body = JSON.parse(response.getContentText())
    } catch (err) {
      throw new Error('FOLIO instances response was not valid JSON.', { cause: err })
    }

    if (!body || !Array.isArray(body.instances)) throw new Error('Expected an instances array from FOLIO.')

    if (Number.isInteger(body.totalRecords)) totalRecords = body.totalRecords

    const page = body.instances.slice(0, remaining)
    instances.push(...page)
    offset += page.length

    if (body.instances.length < pageSize ||
      (Number.isInteger(totalRecords) && offset >= totalRecords)) {
      return { instances, totalRecords, truncated: false }
    }
  }

  // Flag truncation when more records exist (or totalRecords is absent after hitting the cap).
  const truncated = Number.isInteger(totalRecords) ? totalRecords > instances.length : true
  return { instances, totalRecords, truncated }
}

// Ensure values function in Google Sheets before writing them
function cellValue(value) {
  if (value === undefined || value === null) return ''
  const cell = value
  if (typeof cell !== 'string') return cell
  if (cell.length > 50000) throw new Error('An output value exceeds the Google Sheets 50,000-character cell limit. Configure specific JSON fields in INSTANCE_COLUMNS.')
  // Prevent API strings from being evaluated as spreadsheet formulas.
  return /^[=+\-@]/.test(cell) ? "'" + cell : cell
}

// Flatten arrays by index so related values (for example identifier type and value)
// stay together. Empty objects/arrays have no leaf values and produce no columns.
// TODO We need a better way to handle multi-value object arrays which create so many columns
// TODO e.g. subjects are {authorityId, value, sourceId, typeId} so there's 4 cols each
function flattenInstance(value, path = '', fields = new Map()) {
  if (value !== null && typeof value === 'object') {
    Object.keys(value).forEach(key => {
      flattenInstance(value[key], path ? `${path}.${key}` : key, fields)
    })
  } else if (path) {
    fields.set(path, value)
  }
  return fields
}

// Write instances JSON onto a sheet
function writeInstances(instances) {
  const flattened = instances.map(instance => flattenInstance(instance))

  // Use every record to discover fields
  let paths = new Set()
  flattened.forEach(fields => fields.forEach((value, path) => paths.add(path)))

  const columns = INSTANCE_COLUMNS.length
    ? INSTANCE_COLUMNS
    : Array.from(paths, path => ({header: path, path}))

  let rows = [columns.map(column => cellValue(column.header))]
  flattened.forEach(fields => {
    rows.push(columns.map(column => cellValue(fields.get(column.path))))
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
