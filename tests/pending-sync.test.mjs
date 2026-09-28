import { test } from 'node:test';
import assert from 'node:assert/strict';
import { trackChanges, restoreChanges, acknowledgeChanges } from '../src/pending-sync.mjs';

test('October receipt survives reload before server acknowledgement', () => {
  const old = { receipts: [] };
  const receipt = { id: 'october', month: '10/2026', isFinalized: true };
  const pending = JSON.parse(JSON.stringify(trackChanges(old, { receipts: [receipt] })));
  assert.deepEqual(restoreChanges(old, pending).receipts, [receipt]);
});

test('pending edits do not resurrect unrelated server deletions', () => {
  const pending = trackChanges({ receipts: [{ id: 'old' }] }, { receipts: [{ id: 'old' }, { id: 'new' }] });
  assert.deepEqual(restoreChanges({ receipts: [{ id: 'remote' }] }, pending).receipts, [{ id: 'remote' }, { id: 'new' }]);
});

test('edits made during upload remain pending after its response', () => {
  const sent = { receipts: { a: { id: 'a', total: 1 } } };
  const current = { receipts: { a: { id: 'a', total: 2 } } };
  assert.deepEqual(acknowledgeChanges(current, sent), current);
  assert.deepEqual(acknowledgeChanges(sent, sent), {});
});

test('pending deletion survives reload', () => {
  const old = { receipts: [{ id: 'a' }] };
  const pending = trackChanges(old, { receipts: [] });
  assert.deepEqual(restoreChanges(old, pending).receipts, []);
});
