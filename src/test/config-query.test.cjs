const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const vm = require('node:vm')
const path = require('node:path')

const APP_SOURCE = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8')

function loadApp({ configRows = [], fetchImpl, authTenantOverride } = {}) {
  const requests = []
  const createdSheets = []

  const configSheet = {
    getDataRange: () => ({ getDisplayValues: () => configRows })
  }

  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => (name === 'config' ? configSheet : null),
        insertSheet: () => {
          const sheet = {
            name: `Sheet${createdSheets.length + 1}`,
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
      fetch: (url, options) => {
        requests.push({ url, options })
        if (!fetchImpl) throw new Error('Unexpected UrlFetchApp.fetch call')
        return fetchImpl(url, options)
      }
    },
    Logger: { log: () => {} }
  })

  const source = authTenantOverride !== undefined
    ? APP_SOURCE.replace("let AUTH_TENANT_OVERRIDE = ''", `let AUTH_TENANT_OVERRIDE = ${JSON.stringify(authTenantOverride)}`)
    : APP_SOURCE

  vm.runInContext(source, context)
  return { context, requests, createdSheets }
}

function jsonResponse(body, status = 200, headers = {}) {
  return {
    getResponseCode: () => status,
    getContentText: () => JSON.stringify(body),
    getHeaders: () => headers,
    getAllHeaders: () => headers
  }
}

test('getConfig uses defaults when optional settings are omitted', () => {
  const { context } = loadApp({
    configRows: [
      ['API URL', 'https://okapi.example.org/'],
      ['Tenant ID', 'member-a'],
      ['Username', 'librarian']
    ]
  })

  const config = context.getConfig()
  assert.deepEqual({ ...config }, {
    url: 'https://okapi.example.org',
    tenant: 'member-a',
    username: 'librarian',
    maxRows: 2000,
    cqlQuery: ''
  })
})

test('getConfig parses formatted max rows and cql query and rejects invalid max rows', () => {
  const { context } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian'],
      ['Max Rows', '2,500'],
      ['CQL Query', 'source=="MARC"']
    ]
  })

  const config = context.getConfig()
  assert.equal(config.maxRows, 2500)
  assert.equal(config.cqlQuery, 'source=="MARC"')

  for (const badValue of ['0', '-10', '12.5', 'abc']) {
    const bad = loadApp({
      configRows: [
        ['api url', 'https://okapi.example.org'],
        ['tenant id', 'member-a'],
        ['username', 'librarian'],
        ['max rows', badValue]
      ]
    })
    assert.throws(() => bad.context.getConfig(), /positive whole number/)
  }
})

test('buildCqlQuery composes filters and respects existing sortby outside quotes', () => {
  const { context } = loadApp()

  assert.equal(context.buildCqlQuery(''), 'cql.allRecords=1 sortby id')
  assert.equal(context.buildCqlQuery('source=="MARC"'), 'source=="MARC" sortby id')
  assert.equal(context.buildCqlQuery('source=="MARC" sortby title'), 'source=="MARC" sortby title')
  assert.equal(context.buildCqlQuery('title="sortby demo"'), 'title="sortby demo" sortby id')
  assert.equal(
    context.buildCqlQuery(['status.name=="Available"', 'barcode=="123"'], 'effectiveShelvingOrder'),
    '(status.name=="Available") and (barcode=="123") sortby effectiveShelvingOrder'
  )
})

test('AUTH_TENANT_OVERRIDE only changes the auth request tenant header', () => {
  const { context, requests } = loadApp({
    authTenantOverride: 'central',
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'scelc-admin']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      return jsonResponse({ instances: [{ id: '1', title: 'Book' }], totalRecords: 1 })
    }
  })

  const result = context.createInventory('secret')
  assert.equal(result.count, 1)
  assert.equal(requests.length, 2)
  assert.equal(requests[0].options.headers['x-okapi-tenant'], 'central')
  assert.equal(requests[1].options.headers['x-okapi-tenant'], 'member-a')
})

test('getInstances clamps final page size to maxRows and flags truncation accurately', () => {
  const { context, requests } = loadApp({
    fetchImpl: url => {
      const params = new URL(url).searchParams
      const limit = Number(params.get('limit'))
      const offset = Number(params.get('offset'))
      const instances = Array.from({ length: limit }, (_, i) => ({ id: `rec-${offset + i}` }))
      return jsonResponse({ instances, totalRecords: 300 })
    }
  })

  const result = context.getInstances(
    { url: 'https://okapi.example.org', tenant: 'member-a', maxRows: 150, cqlQuery: 'source=="MARC"' },
    'tok-123'
  )

  assert.equal(result.instances.length, 150)
  assert.equal(result.totalRecords, 300)
  assert.equal(result.truncated, true)
  assert.equal(requests.length, 2)
  assert.equal(new URL(requests[0].url).searchParams.get('limit'), '100')
  assert.equal(new URL(requests[1].url).searchParams.get('limit'), '50')
  assert.equal(new URL(requests[1].url).searchParams.get('offset'), '100')
  assert.equal(new URL(requests[0].url).searchParams.get('query'), 'source=="MARC" sortby id')
})

test('getInstances does not flag truncation when totalRecords equals maxRows', () => {
  const { context } = loadApp({
    fetchImpl: url => {
      const limit = Number(new URL(url).searchParams.get('limit'))
      const instances = Array.from({ length: limit }, (_, i) => ({ id: `rec-${i}` }))
      return jsonResponse({ instances, totalRecords: 100 })
    }
  })

  const result = context.getInstances(
    { url: 'https://okapi.example.org', tenant: 'member-a', maxRows: 100, cqlQuery: '' },
    'tok-123'
  )

  assert.equal(result.instances.length, 100)
  assert.equal(result.totalRecords, 100)
  assert.equal(result.truncated, false)
})

test('createInventory skips sheet creation when zero records match', () => {
  const { context, createdSheets } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian'],
      ['cql query', 'title=="nonexistent"']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      return jsonResponse({ instances: [], totalRecords: 0 })
    }
  })

  const result = context.createInventory('secret')
  assert.deepEqual({ ...result }, { sheetName: null, count: 0, totalRecords: 0, truncated: false })
  assert.equal(createdSheets.length, 0)
})
