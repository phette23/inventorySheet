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
    insertColumnsAfter: () => {},
    getRange: () => ({ setValues: value => { rows = value } }),
    setFrozenRows: () => {}
  }
  const context = vm.createContext({ SpreadsheetApp: {
    getActiveSpreadsheet: () => ({ insertSheet: () => sheet, setActiveSheet: () => {} })
  } })
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8'), context)
  return { context, output: () => rows }
}

test('nested fields and indexed arrays preserve values and relationships across records', () => {
  const { context, output } = setup()
  context.writeInstances([
    { id: 'a', contributors: [{ name: 'First', primary: false }], metadata: { createdDate: 'today' }, empty: [] },
    { id: 'b', contributors: [{ name: 'Second' }, { name: '=formula' }], later: 0, nullable: null }
  ])
  const rows = output()
  const value = (row, header) => rows[row][rows[0].indexOf(header)]
  assert.equal(value(1, 'contributors.0.name'), 'First')
  assert.equal(value(1, 'contributors.0.primary'), false)
  assert.equal(value(1, 'metadata.createdDate'), 'today')
  assert.equal(value(1, 'contributors.1.name'), '')
  assert.equal(value(2, 'contributors.1.name'), "'=formula")
  assert.equal(value(2, 'later'), 0)
  assert.equal(value(2, 'nullable'), '')
  assert.equal(rows[0].includes('empty'), false)
})
