const { execFileSync } = require('node:child_process')
const fs = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')

const APP_SOURCE = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8')
const ENV_PATH = path.join(__dirname, '../../.env')

// Inline Node script that runs fetch() in a child process so Apps Script's
// synchronous UrlFetchApp.fetch() contract works verbatim against live FOLIO.
const SYNC_FETCH_SCRIPT = `
const chunks = []
process.stdin.on('data', chunk => chunks.push(chunk))
process.stdin.on('end', async () => {
  try {
    const { url, options = {} } = JSON.parse(Buffer.concat(chunks).toString('utf8'))
    const headers = Object.assign({}, options.headers || {})
    if (options.contentType && !Object.keys(headers).some(k => k.toLowerCase() === 'content-type')) {
      headers['Content-Type'] = options.contentType
    }
    const res = await fetch(url, {
      method: (options.method || 'get').toUpperCase(),
      headers,
      body: options.payload,
      redirect: options.followRedirects === false ? 'manual' : 'follow'
    })
    const text = await res.text()
    const allHeaders = {}
    res.headers.forEach((value, key) => { allHeaders[key] = value })
    if (typeof res.headers.getSetCookie === 'function') {
      const cookies = res.headers.getSetCookie()
      if (cookies.length) allHeaders['set-cookie'] = cookies
    }
    process.stdout.write(JSON.stringify({ status: res.status, text, allHeaders }))
  } catch (err) {
    process.stdout.write(JSON.stringify({ error: err.message }))
    process.exitCode = 1
  }
})
`

function getLiveEnv() {
  if (!process.env.FOLIO_URL && fs.existsSync(ENV_PATH) && typeof process.loadEnvFile === 'function') {
    process.loadEnvFile(ENV_PATH)
  }

  const env = {
    url: (process.env.FOLIO_URL || '').trim(),
    tenant: (process.env.FOLIO_TENANT || '').trim(),
    username: (process.env.FOLIO_USERNAME || '').trim(),
    password: process.env.FOLIO_PASSWORD || ''
  }

  const missing = Object.entries(env).filter(([, value]) => !value).map(([key]) => `FOLIO_${key.toUpperCase()}`)
  if (missing.length) {
    throw new Error(`Missing live FOLIO credentials in .env or environment: ${missing.join(', ')}`)
  }

  return env
}

function liveFetch(url, options = {}) {
  const started = Date.now()
  const raw = execFileSync(process.execPath, ['-e', SYNC_FETCH_SCRIPT], {
    // Pass URL, headers, and payload via stdin so credentials never appear in argv.
    input: JSON.stringify({ url, options }),
    encoding: 'utf8',
    maxBuffer: 20 * 1024 * 1024
  })

  const result = JSON.parse(raw)
  if (result.error) throw new Error(result.error)

  const headers = {}
  for (const [key, value] of Object.entries(result.allHeaders)) {
    headers[key] = Array.isArray(value) ? value[0] : value
  }

  if (!options.muteHttpExceptions && (result.status < 200 || result.status >= 300)) {
    throw new Error(`HTTP ${result.status} from ${url}`)
  }

  return {
    ms: Date.now() - started,
    response: {
      getResponseCode: () => result.status,
      getContentText: () => result.text,
      getHeaders: () => headers,
      getAllHeaders: () => result.allHeaders
    }
  }
}

function loadLiveApp({ configOverrides = {}, fieldOverrides = {}, env = getLiveEnv() } = {}) {
  const requests = []
  const createdSheets = []
  const logs = []

  const configRows = [
    ['API URL', env.url],
    ['Tenant ID', env.tenant],
    ['Username', env.username],
    ...Object.entries(configOverrides),
    ...Object.entries(fieldOverrides)
  ]

  const configSheet = {
    name: 'config',
    getName: () => 'config',
    getDataRange: () => ({
      getDisplayValues: () => configRows.map(row => row.map(v => (typeof v === 'boolean' ? String(v).toUpperCase() : String(v))))
    })
  }

  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => {
          if (name === 'config') return configSheet
          return createdSheets.find(sheet => sheet.name === name) || null
        },
        insertSheet: name => {
          const sheet = {
            name: name || `Sheet${createdSheets.length + 1}`,
            rows: null,
            getName: () => sheet.name,
            getMaxRows: () => 1000,
            getMaxColumns: () => 26,
            insertRowsAfter: () => {},
            insertColumnsAfter: () => {},
            getRange: () => ({ setValues: value => { sheet.rows = value } }),
            setFrozenRows: () => {}
          }
          createdSheets.push(sheet)
          return sheet
        },
        setActiveSheet: sheet => sheet
      })
    },
    UrlFetchApp: {
      fetch: (url, options = {}) => {
        const { response, ms } = liveFetch(url, options)
        // Log non-sensitive request metadata only (never payload or auth token).
        requests.push({
          method: (options.method || 'get').toUpperCase(),
          url,
          tenant: options.headers && options.headers['x-okapi-tenant'],
          status: response.getResponseCode(),
          ms
        })
        return response
      }
    },
    Logger: { log: msg => logs.push(String(msg)) }
  })

  vm.runInContext(APP_SOURCE, context)
  return { context, env, requests, createdSheets, logs }
}

module.exports = { getLiveEnv, liveFetch, loadLiveApp }
