const collections = ['rooms', 'tenants', 'memberships', 'contracts', 'receipts', 'moveOutReports', 'contractRenewals', 'roomTransfers', 'suppliers', 'expenseCategories', 'expensePayments'];

function validateSnapshot(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return 'Invalid snapshot';
  // Missing collections must not be interpreted as requests to delete records.
  for (const key of collections) {
    if (!Array.isArray(payload[key])) return `Missing or invalid collection: ${key}`;
    const ids = new Set();
    for (const row of payload[key]) {
      if (!row || typeof row.id !== 'string' || !row.id.trim() || ids.has(row.id)) return `Invalid or duplicate id in ${key}`;
      ids.add(row.id);
    }
  }
  return null;
}

module.exports = { validateSnapshot };
