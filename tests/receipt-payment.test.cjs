const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../src/App.jsx'), 'utf8');
const start = source.indexOf('function getReceiptPaymentState(');
const end = source.indexOf('const OTHER_RECEIPT_TYPES', start);
const context = vm.createContext({});
vm.runInContext(source.slice(start, end), context);
const { recordReceiptPayment, getReceiptPaymentState } = context;
const date = '2026-09-03';

test('collecting the remaining balance preserves previous payments and clears debt', () => {
  const receipt = { total: 1000000, paidAmount: 774194, debt: 225806 };
  const updated = recordReceiptPayment(receipt, 225806, date);
  assert.equal(updated.paidAmount, 1000000);
  assert.equal(updated.debt, 0);
  assert.equal(updated.status, 'Đã thanh toán');
  assert.equal(receipt.paidAmount, 774194);
  assert.throws(() => recordReceiptPayment(updated, 225806, date));
});

test('successive partial payments accumulate', () => {
  const first = recordReceiptPayment({ total: 1000000, paidAmount: 0 }, 200000, date);
  const second = recordReceiptPayment(first, 300000, date);
  assert.equal(second.paidAmount, 500000);
  assert.equal(second.debt, 500000);
  assert.equal(second.status, 'Nợ một phần');
});

test('explicit adjustments preserve the base collection', () => {
  const updated = recordReceiptPayment({ total: 1000000, paidAmount: 1000000, adjustmentDueAmount: 300000, adjustmentPaidAmount: 100000 }, 200000, date);
  assert.equal(updated.paidAmount, 1000000);
  assert.equal(updated.adjustmentPaidAmount, 300000);
  assert.equal(updated.debt, 0);
  assert.equal(updated.status, 'Đã thanh toán');
});

test('a paid utility-only receipt does not become unpaid again', () => {
  const state = getReceiptPaymentState({ type: 'monthly', isFinalized: true, total: 225806, paidAmount: 225806, rent: 0, fixedServices: 0, electricAmount: 225806 });
  assert.equal(state.isPaid, true);
  assert.equal(state.debt, 0);
});

test('invalid payments and cancelled receipts are rejected', () => {
  const receipt = { total: 100, paidAmount: 20 };
  for (const amount of [-1, 0, 81, 1.5, NaN, Infinity, 'abc']) {
    assert.throws(() => recordReceiptPayment(receipt, amount, date));
  }
  assert.throws(() => recordReceiptPayment(receipt, 10, ''));
  assert.throws(() => recordReceiptPayment({ ...receipt, status: 'Đã hủy' }, 10, date));
});
