const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const vm = require('node:vm')
const path = require('node:path')

function setup() {
  let rows
  const sheet = {
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
    insertRowsAfter: () => {},
    insertColumnsAfter: () => {},
    getRange: () => ({ setValues: value => { rows = value } }),
    setFrozenRows: () => {}
  }

  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ insertSheet: () => sheet, setActiveSheet: () => sheet })
    }
  })

  vm.runInContext(fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8'), context)
  return { context, output: () => rows }
}

test('writeItems extracts default item fields cleanly without nested column explosion', () => {
  const { context, output } = setup()

  context.writeItems([
    {
      id: 'item-1',
      barcode: '3123456789',
      title: 'Intro to Cataloging',
      effectiveCallNumberComponents: { prefix: 'REF', callNumber: 'Z693 .I58', suffix: '2024' },
      volume: 'v.1',
      copyNumber: 'c.2',
      effectiveLocation: { id: 'loc-1', name: 'Main Stacks' },
      status: { name: 'Available', date: '2026-10-01T00:00:00Z' },
      notes: [
        { itemNoteTypeId: 'n1', note: 'Inventoried 2026-10-15', staffOnly: true },
        { itemNoteTypeId: 'n2', note: 'Needs spine label', staffOnly: true }
      ],
      administrativeNotes: ['Shelf check passed', '2026-10-15']
    },
    {
      id: 'item-2',
      barcode: '3999999999',
      title: '=HYPERLINK("bad")',
      callNumber: 'QA76.73',
      effectiveLocation: { name: 'Oversize' },
      status: { name: 'Missing' },
      notes: [],
      administrativeNotes: []
    }
  ])

  const rows = output()
  assert.deepEqual([...rows[0]], [
    'Item UUID',
    'Barcode',
    'Title',
    'Call number',
    'Location',
    'Item status',
    'Item notes',
    'Administrative notes'
  ])

  const value = (row, header) => rows[row][rows[0].indexOf(header)]
  assert.equal(value(1, 'Item UUID'), 'item-1')
  assert.equal(value(1, 'Barcode'), '3123456789')
  assert.equal(value(1, 'Title'), 'Intro to Cataloging')
  assert.equal(value(1, 'Call number'), 'REF Z693 .I58 2024 v.1 c.2')
  assert.equal(value(1, 'Location'), 'Main Stacks')
  assert.equal(value(1, 'Item status'), 'Available')
  assert.equal(value(1, 'Item notes'), 'Inventoried 2026-10-15\nNeeds spine label')
  assert.equal(value(1, 'Administrative notes'), 'Shelf check passed\n2026-10-15')

  assert.equal(value(2, 'Title'), "'=HYPERLINK(\"bad\")")
  assert.equal(value(2, 'Call number'), 'QA76.73')
  assert.equal(value(2, 'Item notes'), '')
  assert.equal(value(2, 'Administrative notes'), '')
})

test('writeItems supports custom field lists including nested arrays like contributors and circulation notes', () => {
  const { context, output } = setup()
  const allFields = vm.runInContext('FIELDS', context)
  const chosen = allFields.filter(field => [
    'Barcode',
    'Contributor names',
    'Circulation notes',
    'Material type'
  ].includes(field.name))

  context.writeItems([
    {
      barcode: '111',
      contributorNames: [{ name: 'Kern, Leslie' }, { name: 'Doe, Jane' }],
      circulationNotes: [{ note: 'Check CD pocket' }, { note: 'Fragile binding' }],
      materialType: { name: 'book' }
    }
  ], chosen)

  const rows = output()
  assert.deepEqual([...rows[0]], ['Barcode', 'Contributor names', 'Circulation notes', 'Material type'])
  assert.deepEqual([...rows[1]], ['111', 'Kern, Leslie; Doe, Jane', 'Check CD pocket\nFragile binding', 'book'])
})
