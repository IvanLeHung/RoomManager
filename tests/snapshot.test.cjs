const { test } = require('node:test');
const assert = require('node:assert/strict');
const { validateSnapshot } = require('../api/lib/snapshot');
const empty = () => Object.fromEntries(['rooms', 'tenants', 'memberships', 'contracts', 'receipts', 'moveOutReports', 'contractRenewals', 'roomTransfers', 'suppliers', 'expenseCategories', 'expensePayments'].map(key => [key, []]));

test('complete snapshot is valid; missing collections and duplicate IDs are rejected', () => {
  assert.equal(validateSnapshot(empty()), null);
  assert.ok(validateSnapshot({ receipts: [] }));
  assert.ok(validateSnapshot({ ...empty(), rooms: [{ id: '301' }, { id: '301' }] }));
  assert.ok(validateSnapshot({ ...empty(), rooms: [null] }));
});

test('API loads and rejects invalid requests before opening database', async () => {
  const handler = require('../api/data');
  const response = { setHeader() {}, status(code) { this.code = code; return this; }, json(body) { this.body = body; return this; } };
  await handler({ method: 'POST', body: { type: 'full_sync', payload: {} } }, response);
  assert.equal(response.code, 400);
  await handler({ method: 'DELETE' }, response);
  assert.equal(response.code, 405);
});
