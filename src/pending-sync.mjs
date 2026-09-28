export function trackChanges(previous, next, pending = {}) {
  const result = { ...pending };
  for (const key of Object.keys(next)) {
    if (!Array.isArray(next[key]) || !Array.isArray(previous[key])) continue;
    const before = new Map(previous[key].map(row => [row.id, row]));
    const after = new Map(next[key].map(row => [row.id, row]));
    const changes = { ...result[key] };
    for (const [id, row] of after) {
      if (id && JSON.stringify(before.get(id)) !== JSON.stringify(row)) changes[id] = row;
    }
    for (const id of before.keys()) {
      if (id && !after.has(id)) changes[id] = null;
    }
    if (Object.keys(changes).length) result[key] = changes;
  }
  return result;
}

export function restoreChanges(cloud, pending) {
  const result = { ...cloud };
  for (const [key, changes] of Object.entries(pending)) {
    const rows = new Map((cloud[key] || []).map(row => [row.id, row]));
    for (const [id, row] of Object.entries(changes)) {
      if (row === null) rows.delete(id);
      else rows.set(id, row);
    }
    result[key] = [...rows.values()];
  }
  return result;
}

export function acknowledgeChanges(current, sent) {
  const result = {};
  for (const [key, changes] of Object.entries(current)) {
    const remaining = Object.fromEntries(Object.entries(changes).filter(([id, row]) =>
      JSON.stringify(row) !== JSON.stringify(sent[key]?.[id])));
    if (Object.keys(remaining).length) result[key] = remaining;
  }
  return result;
}
