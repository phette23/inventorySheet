const { test } = require('node:test')
const assert = require('node:assert/strict')
const { getLiveEnv, loadLiveApp } = require('./folio-live.cjs')

const env = getLiveEnv()

test('live FOLIO auth and default /inventory/items export work end-to-end', t => {
  const { context, requests, createdSheets } = loadLiveApp({
    env,
    configOverrides: { 'Max rows': 5 }
  })

  const result = context.createInventory(env.password)
  assert.ok(result.count >= 0 && result.count <= 5)
  assert.equal(requests.length, 2, 'Default run should only make auth + items requests')

  if (result.count > 0) {
    assert.equal(createdSheets.length, 1)
    assert.equal(createdSheets[0].rows.length, result.count + 1)
  } else {
    t.diagnostic(`Tenant "${env.tenant}" returned 0 items (central/empty tenant); verified auth + /inventory/items query.`)
    assert.equal(result.sheetName, null)
    assert.equal(createdSheets.length, 0)
  }
})

test('live FOLIO Stage 4 status and call-number range filters execute cleanly', () => {
  const { context } = loadLiveApp({
    env,
    configOverrides: {
      'Max rows': 5,
      'Status filter': 'Available',
      'Call number start': 'A',
      'Call number end': 'Z'
    }
  })

  const config = context.getConfig()
  const token = context.authenticate(config, env.password)
  const { items } = context.getItems(config, token)

  assert.ok(Array.isArray(items))
  items.forEach(item => {
    assert.equal(item.status && item.status.name, 'Available')
  })
})

test('live FOLIO Stage 5 reference lookups, storage batching, and enrichment work', t => {
  const { context, requests, createdSheets } = loadLiveApp({
    env,
    configOverrides: { 'Max rows': 5 },
    fieldOverrides: {
      'Holdings action note': true,
      OCN: true
    }
  })

  const config = context.getConfig()
  const token = context.authenticate(config, env.password)

  // Always verify reference-data type resolution and storage batching even on a 0-item tenant.
  const actionNoteTypeId = context.lookupTypeId(
    config, token, '/holdings-note-types', 'holdingsNoteTypes', 'Action note'
  )
  const ocnTypeId = context.lookupTypeId(
    config, token, '/identifier-types', 'identifierTypes', 'OCLC'
  )
  assert.ok(actionNoteTypeId, 'Expected Action note holdings note type UUID')
  assert.ok(ocnTypeId, 'Expected OCLC identifier type UUID')

  const sampleInstances = context.httpJson(config, '/instance-storage/instances?limit=5', token, 'instances')
  assert.ok(Array.isArray(sampleInstances.instances))
  if (sampleInstances.instances.length > 0) {
    const ids = sampleInstances.instances.map(instance => instance.id)
    const fetched = context.fetchRecordsByIds(config, token, '/instance-storage/instances', 'instances', ids)
    assert.equal(fetched.size, ids.length)
  }

  const result = context.createInventory(env.password)
  if (result.count > 0) {
    const headers = createdSheets[0].rows[0]
    assert.ok(headers.includes('Holdings action note'))
    assert.ok(headers.includes('OCN'))
    assert.ok(requests.some(r => r.url.includes('/holdings-storage/holdings?')))
    assert.ok(requests.some(r => r.url.includes('/instance-storage/instances?')))
  } else {
    t.diagnostic(`Tenant "${env.tenant}" has 0 items; verified Action note/OCLC lookups and instance batching directly.`)
  }
})
