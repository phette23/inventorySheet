const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const vm = require('node:vm')
const path = require('node:path')

const APP_SOURCE = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8')

function loadApp({ configRows = [], hasConfigSheet = true, fetchImpl } = {}) {
  const requests = []
  const createdSheets = []
  const storedConfigRows = configRows.map(row => [...row])
  let checkboxesInserted = 0
  let rowGroupDepthShifts = 0
  let collapsedGroups = 0

  function makeConfigSheet() {
    return {
      name: 'config',
      getName: () => 'config',
      getLastRow: () => storedConfigRows.length,
      getDataRange: () => ({
        getDisplayValues: () => storedConfigRows.map(row => row.map(v => (typeof v === 'boolean' ? String(v).toUpperCase() : String(v))))
      }),
      getRange: (startRow, startCol, numRows, numCols) => ({
        getDisplayValues: () => storedConfigRows
          .slice(startRow - 1, startRow - 1 + numRows)
          .map(row => row.slice(startCol - 1, startCol - 1 + numCols).map(v => (typeof v === 'boolean' ? String(v).toUpperCase() : String(v)))),
        setValues: values => {
          values.forEach((rowVals, offset) => {
            const idx = startRow - 1 + offset
            while (storedConfigRows.length <= idx) storedConfigRows.push(['', ''])
            rowVals.forEach((val, colOffset) => {
              storedConfigRows[idx][startCol - 1 + colOffset] = val
            })
          })
        },
        setFontWeight: () => {},
        insertCheckboxes: () => { checkboxesInserted += numRows },
        shiftRowGroupDepth: delta => { rowGroupDepthShifts += delta }
      }),
      collapseAllRowGroups: () => { collapsedGroups += 1 }
    }
  }

  let configSheet = hasConfigSheet ? makeConfigSheet() : null

  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => {
          if (name === 'config') return configSheet
          return createdSheets.find(sheet => sheet.name === name) || null
        },
        insertSheet: name => {
          if (name === 'config') {
            configSheet = makeConfigSheet()
            return configSheet
          }
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
      fetch: (url, options) => {
        requests.push({ url, options })
        if (!fetchImpl) throw new Error('Unexpected UrlFetchApp.fetch call')
        return fetchImpl(url, options)
      }
    },
    Logger: { log: () => {} }
  })

  vm.runInContext(APP_SOURCE, context)
  return {
    context,
    requests,
    createdSheets,
    storedConfigRows,
    stats: () => ({ checkboxesInserted, rowGroupDepthShifts, collapsedGroups })
  }
}

function jsonResponse(body, status = 200, headers = {}) {
  return {
    getResponseCode: () => status,
    getContentText: () => JSON.stringify(body),
    getHeaders: () => headers,
    getAllHeaders: () => headers
  }
}

test('getConfig uses defaults when optional settings and field rows are omitted', () => {
  const { context } = loadApp({
    configRows: [
      ['API URL', 'https://okapi.example.org/'],
      ['Tenant ID', 'member-a'],
      ['Username', 'librarian']
    ]
  })

  const config = context.getConfig()
  assert.equal(config.url, 'https://okapi.example.org')
  assert.equal(config.tenant, 'member-a')
  assert.equal(config.username, 'librarian')
  assert.equal(config.maxRows, 2000)
  assert.equal(config.cqlQuery, '')
  assert.deepEqual(Array.from(config.fields, field => field.name), [
    'Item UUID',
    'Barcode',
    'Title',
    'Call number',
    'Location',
    'Item status',
    'Item notes',
    'Administrative notes'
  ])
})

test('getConfig parses field checkboxes and always retains required fields', () => {
  const { context } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian'],
      ['Barcode', false],
      ['Item UUID', false],
      ['Item notes', false],
      ['Administrative notes', false],
      ['Contributor names', true],
      ['Material type', 'TRUE']
    ]
  })

  const config = context.getConfig()
  assert.deepEqual(Array.from(config.fields, field => field.name), [
    'Barcode',
    'Title',
    'Call number',
    'Location',
    'Item status',
    'Contributor names',
    'Material type'
  ])
})

test('setupConfigSheet creates collapsed field checkboxes and preserves existing values on re-run', () => {
  const { context, storedConfigRows, stats } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian']
    ]
  })

  const fieldCount = vm.runInContext('FIELDS.length', context)
  context.setupConfigSheet()
  const firstRun = stats()
  assert.equal(firstRun.checkboxesInserted, fieldCount)
  assert.equal(firstRun.rowGroupDepthShifts, 1)
  assert.equal(firstRun.collapsedGroups, 1)
  assert.equal(storedConfigRows[0][1], 'https://okapi.example.org')
  assert.deepEqual(storedConfigRows.slice(3, 9), [
    ['Call number start', ''],
    ['Call number end', ''],
    ['Status filter', ''],
    ['Note contains', ''],
    ['Max rows', 2000],
    ['CQL query', '']
  ])

  // Re-running should be idempotent and not duplicate rows or re-shift group depth.
  const countAfterFirst = storedConfigRows.length
  context.setupConfigSheet()
  assert.equal(storedConfigRows.length, countAfterFirst)
  assert.equal(stats().rowGroupDepthShifts, 1)
})

test('setupConfigSheet scaffolds properly cased labels without colliding with field names', () => {
  const { context, storedConfigRows } = loadApp({ hasConfigSheet: false })

  context.setupConfigSheet()
  assert.deepEqual(storedConfigRows.slice(0, 9), [
    ['API URL', ''],
    ['Tenant ID', ''],
    ['Username', ''],
    ['Call number start', ''],
    ['Call number end', ''],
    ['Status filter', ''],
    ['Note contains', ''],
    ['Max rows', 2000],
    ['CQL query', '']
  ])

  const settingLabels = Object.values(vm.runInContext('CONFIG_SETTINGS', context))
    .map(setting => setting.label.toLocaleLowerCase())
  const fieldNames = vm.runInContext('FIELDS', context)
    .map(field => field.name.toLocaleLowerCase())
  const overlap = settingLabels.filter(label => fieldNames.includes(label))
  assert.deepEqual(overlap, [])
})

test('getConfig parses formatted max rows and cql query and rejects invalid max rows', () => {
  const { context } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian'],
      ['Max Rows', '2,500'],
      ['CQL Query', 'status.name=="Available"']
    ]
  })

  const config = context.getConfig()
  assert.equal(config.maxRows, 2500)
  assert.equal(config.cqlQuery, 'status.name=="Available"')

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

test('buildCqlQuery composes filters and keeps explicit sortby outside parentheses', () => {
  const { context } = loadApp()

  assert.equal(context.buildCqlQuery(''), 'cql.allRecords=1 sortby effectiveShelvingOrder')
  assert.equal(context.buildCqlQuery('status.name=="Available"'), 'status.name=="Available" sortby effectiveShelvingOrder')
  assert.equal(context.buildCqlQuery('status.name=="Available" sortby barcode'), 'status.name=="Available" sortby barcode')
  assert.equal(context.buildCqlQuery('title="sortby demo"'), 'title="sortby demo" sortby effectiveShelvingOrder')
  assert.equal(
    context.buildCqlQuery(['status.name=="Available"', 'barcode=="123" sortby barcode']),
    '(status.name=="Available") and (barcode=="123") sortby barcode'
  )
  assert.equal(
    context.buildCqlQuery(['status.name=="Available"', 'sortby barcode']),
    'status.name=="Available" sortby barcode'
  )
})

test('buildFilterClauses translates plain-language filters and composes with raw CQL', () => {
  const { context } = loadApp()

  const clauses = context.buildFilterClauses({
    callNumberStart: 'qa',
    callNumberEnd: 'qc',
    statusFilter: 'Available, Missing',
    noteContains: '2026-11',
    cqlQuery: 'materialType.name=="book"'
  })

  assert.equal(
    context.buildCqlQuery(clauses),
    '(effectiveShelvingOrder>="QA") and (effectiveShelvingOrder<="QC~") and (status.name=="Available" or status.name=="Missing") and (notes=="*2026-11*" or administrativeNotes=="*2026-11*") and (materialType.name=="book") sortby effectiveShelvingOrder'
  )
})

test('buildSheetName formats filter ranges and deduplicates repeated runs', () => {
  const { context, createdSheets } = loadApp({
    configRows: [
      ['API URL', 'https://okapi.example.org'],
      ['Tenant ID', 'member-a'],
      ['Username', 'librarian'],
      ['Call number start', 'qa'],
      ['Call number end', 'qc'],
      ['Status filter', 'Available']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      return jsonResponse({ items: [{ id: '1', barcode: '123', title: 'Book' }], totalRecords: 1 })
    }
  })

  const fixedDate = new Date(2026, 10, 23)
  assert.equal(context.buildSheetName({}, fixedDate), 'Inventory 2026-11-23')
  assert.equal(
    context.buildSheetName({ callNumberStart: 'qa', callNumberEnd: 'qc', statusFilter: 'Available' }, fixedDate),
    'QA-QC Available 2026-11-23'
  )
  assert.equal(
    context.buildSheetName({ callNumberStart: 'qa', callNumberEnd: 'qa' }, fixedDate),
    'QA 2026-11-23'
  )
  assert.equal(
    context.buildSheetName({ callNumberStart: 'qa' }, fixedDate),
    'QA+ 2026-11-23'
  )

  const first = context.createInventory('secret')
  const second = context.createInventory('secret')
  assert.ok(first.sheetName.startsWith('QA-QC Available '))
  assert.equal(second.sheetName, `${first.sheetName} (2)`)
  assert.equal(createdSheets.length, 2)
})

test('@scelc.org usernames authenticate against the SCELC central tenant while querying the member tenant', () => {
  const scelc = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'Staff@SCELC.ORG']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      return jsonResponse({ items: [{ id: '1', barcode: '123', title: 'Book' }], totalRecords: 1 })
    }
  })

  const result = scelc.context.createInventory('secret')
  assert.equal(result.count, 1)
  assert.equal(scelc.requests.length, 2)
  assert.equal(scelc.requests[0].options.headers['x-okapi-tenant'], 'cs00000003')
  assert.equal(scelc.requests[1].options.headers['x-okapi-tenant'], 'member-a')
  assert.ok(scelc.requests[1].url.includes('/inventory/items?'))

  const member = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian@puc.edu']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-456' })
      }
      return jsonResponse({ items: [{ id: '2', barcode: '456', title: 'Book 2' }], totalRecords: 1 })
    }
  })

  member.context.createInventory('secret')
  assert.equal(member.requests[0].options.headers['x-okapi-tenant'], 'member-a')
  assert.equal(member.requests[1].options.headers['x-okapi-tenant'], 'member-a')
})

test('getItems clamps final page size to maxRows and flags truncation accurately', () => {
  const { context, requests } = loadApp({
    fetchImpl: url => {
      const params = new URL(url).searchParams
      const limit = Number(params.get('limit'))
      const offset = Number(params.get('offset'))
      const items = Array.from({ length: limit }, (_, i) => ({ id: `rec-${offset + i}` }))
      return jsonResponse({ items, totalRecords: 300 })
    }
  })

  const result = context.getItems(
    { url: 'https://okapi.example.org', tenant: 'member-a', maxRows: 150, cqlQuery: 'status.name=="Available"' },
    'tok-123'
  )

  assert.equal(result.items.length, 150)
  assert.equal(result.totalRecords, 300)
  assert.equal(result.truncated, true)
  assert.equal(requests.length, 2)
  assert.ok(requests[0].url.includes('/inventory/items?'))
  assert.equal(new URL(requests[0].url).searchParams.get('limit'), '100')
  assert.equal(new URL(requests[1].url).searchParams.get('limit'), '50')
  assert.equal(new URL(requests[1].url).searchParams.get('offset'), '100')
  assert.equal(new URL(requests[0].url).searchParams.get('query'), 'status.name=="Available" sortby effectiveShelvingOrder')
})

test('getItems does not flag truncation when totalRecords equals maxRows', () => {
  const { context } = loadApp({
    fetchImpl: url => {
      const limit = Number(new URL(url).searchParams.get('limit'))
      const items = Array.from({ length: limit }, (_, i) => ({ id: `rec-${i}` }))
      return jsonResponse({ items, totalRecords: 100 })
    }
  })

  const result = context.getItems(
    { url: 'https://okapi.example.org', tenant: 'member-a', maxRows: 100, cqlQuery: '' },
    'tok-123'
  )

  assert.equal(result.items.length, 100)
  assert.equal(result.totalRecords, 100)
  assert.equal(result.truncated, false)
})

test('createInventory skips sheet creation when zero records match', () => {
  const { context, createdSheets } = loadApp({
    configRows: [
      ['api url', 'https://okapi.example.org'],
      ['tenant id', 'member-a'],
      ['username', 'librarian'],
      ['cql query', 'barcode=="nonexistent"']
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      return jsonResponse({ items: [], totalRecords: 0 })
    }
  })

  const result = context.createInventory('secret')
  assert.deepEqual({ ...result }, { sheetName: null, count: 0, totalRecords: 0, truncated: false })
  assert.equal(createdSheets.length, 0)
})

test('createInventory enriches items with holdings action note and OCN when checked', () => {
  const { context, requests, createdSheets } = loadApp({
    configRows: [
      ['API URL', 'https://okapi.example.org'],
      ['Tenant ID', 'member-a'],
      ['Username', 'librarian'],
      ['Holdings action note', true],
      ['OCN', true]
    ],
    fetchImpl: url => {
      if (url.includes('/authn/login-with-expiry')) {
        return jsonResponse({ accessToken: 'tok-123' })
      }
      if (url.includes('/inventory/items?')) {
        return jsonResponse({
          items: [
            { id: 'i1', barcode: '111', title: 'Book A', holdingsRecordId: 'h1' },
            { id: 'i2', barcode: '222', title: 'Book A copy 2', holdingsRecordId: 'h1' },
            { id: 'i3', barcode: '333', title: 'Book B', holdingsRecordId: 'h2' }
          ],
          totalRecords: 3
        })
      }
      if (url.includes('/holdings-note-types?')) {
        return jsonResponse({
          holdingsNoteTypes: [
            { id: 'nt-other', name: 'General note' },
            { id: 'nt-action', name: 'Action note' }
          ]
        })
      }
      if (url.includes('/identifier-types?')) {
        return jsonResponse({
          identifierTypes: [
            { id: 'id-isbn', name: 'ISBN' },
            { id: 'id-oclc', name: 'oclc' }
          ]
        })
      }
      if (url.includes('/holdings-storage/holdings?')) {
        return jsonResponse({
          holdingsRecords: [
            {
              id: 'h1',
              instanceId: 'inst-1',
              notes: [
                { holdingsNoteTypeId: 'nt-other', note: 'Ignore me' },
                { holdingsNoteTypeId: 'nt-action', note: 'SCELC shared print retention' },
                { holdingsNoteTypeId: 'nt-action', note: 'Retain until 2040' }
              ]
            },
            {
              id: 'h2',
              instanceId: 'inst-2',
              notes: []
            }
          ]
        })
      }
      if (url.includes('/instance-storage/instances?')) {
        return jsonResponse({
          instances: [
            {
              id: 'inst-1',
              identifiers: [
                { identifierTypeId: 'id-isbn', value: '9780123456789' },
                { identifierTypeId: 'id-oclc', value: '(OCoLC)123456' },
                { identifierTypeId: 'id-oclc', value: '(OCoLC)789012' }
              ]
            },
            {
              id: 'inst-2',
              identifiers: [{ identifierTypeId: 'id-isbn', value: '9789999999999' }]
            }
          ]
        })
      }
      throw new Error(`Unexpected URL: ${url}`)
    }
  })

  const result = context.createInventory('secret')
  assert.equal(result.count, 3)

  const rows = createdSheets[0].rows
  const col = (row, header) => rows[row][rows[0].indexOf(header)]
  assert.equal(col(1, 'Holdings action note'), 'SCELC shared print retention\nRetain until 2040')
  assert.equal(col(1, 'OCN'), '(OCoLC)123456; (OCoLC)789012')
  assert.equal(col(2, 'Holdings action note'), 'SCELC shared print retention\nRetain until 2040')
  assert.equal(col(2, 'OCN'), '(OCoLC)123456; (OCoLC)789012')
  assert.equal(col(3, 'Holdings action note'), '')
  assert.equal(col(3, 'OCN'), '')

  // Deduplicated holdings query should only request h1 once alongside h2.
  const holdingsReq = requests.find(r => r.url.includes('/holdings-storage/holdings?'))
  assert.equal(new URL(holdingsReq.url).searchParams.get('limit'), '2')
  assert.equal(new URL(holdingsReq.url).searchParams.get('query'), 'id==("h1" or "h2")')
})

test('enrichItems skips unnecessary lookups when only one enrichment field is selected', () => {
  const { context, requests } = loadApp({
    fetchImpl: url => {
      if (url.includes('/identifier-types?')) {
        return jsonResponse({ identifierTypes: [{ id: 'id-oclc', name: 'OCLC' }] })
      }
      if (url.includes('/instance-storage/instances?')) {
        return jsonResponse({
          instances: [{ id: 'inst-1', identifiers: [{ identifierTypeId: 'id-oclc', value: '(OCoLC)555' }] }]
        })
      }
      throw new Error(`Unexpected URL: ${url}`)
    }
  })

  const allFields = vm.runInContext('FIELDS', context)
  const ocnOnly = allFields.filter(f => f.name === 'Barcode' || f.name === 'OCN')
  const enriched = context.enrichItems(
    { url: 'https://okapi.example.org', tenant: 'member-a', fields: ocnOnly },
    'tok-123',
    [{ id: 'i1', barcode: '111', holdingsRecordId: 'h1', instanceId: 'inst-1' }]
  )

  assert.equal(enriched[0]._ocn, '(OCoLC)555')
  assert.equal(requests.some(r => r.url.includes('/holdings-note-types')), false)
  assert.equal(requests.some(r => r.url.includes('/holdings-storage/holdings')), false)
})

test('fetchRecordsByIds chunks large UUID lists into URL-safe batches', () => {
  const { context, requests } = loadApp({
    fetchImpl: url => {
      const query = new URL(url).searchParams.get('query')
      const ids = [...query.matchAll(/"([^"]+)"/g)].map(match => match[1])
      return jsonResponse({ holdingsRecords: ids.map(id => ({ id, instanceId: `inst-${id}` })) })
    }
  })

  const ids = Array.from({ length: 85 }, (_, i) => `h-${i}`)
  const map = context.fetchRecordsByIds(
    { url: 'https://okapi.example.org', tenant: 'member-a' },
    'tok-123',
    '/holdings-storage/holdings',
    'holdingsRecords',
    [...ids, 'h-0', '', null]
  )

  assert.equal(map.size, 85)
  assert.equal(requests.length, 3)
  assert.deepEqual(
    requests.map(r => Number(new URL(r.url).searchParams.get('limit'))),
    [40, 40, 5]
  )
})

test('lookupTypeId throws a clear error when a required FOLIO type name is missing', () => {
  const { context } = loadApp({
    fetchImpl: () => jsonResponse({ holdingsNoteTypes: [{ id: 'nt-1', name: 'General note' }] })
  })

  assert.throws(
    () => context.lookupTypeId(
      { url: 'https://okapi.example.org', tenant: 'member-a' },
      'tok-123',
      '/holdings-note-types',
      'holdingsNoteTypes',
      'Action note'
    ),
    /Action note/
  )
})
