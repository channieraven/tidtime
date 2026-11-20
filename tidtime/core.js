const Decimal = require("decimal.js");

// keyOf returns a stable string or null. Each unkeyed row gets its own group.
function matchRows(before, after, keyOf) {
  const groups = new Map();
  for (const [side, rows] of Object.entries({ before, after })) {
    for (const row of rows) {
      const key = keyOf(row);
      const missing = key == null || String(key).trim() === "";
      const bucket = missing ? Symbol() : key;
      if (!groups.has(bucket))
        groups.set(bucket, {
          key: missing ? null : key,
          before: [],
          after: [],
        });
      groups.get(bucket)[side].push(row);
    }
  }
  return Array.from(groups.values(), (group) => {
    let status = "paired";
    if (group.key === null) status = "missing-key";
    else if (group.before.length > 1 || group.after.length > 1)
      status = "duplicate";
    else if (!group.before.length) status = "after-only";
    else if (!group.after.length) status = "before-only";
    return { ...group, status };
  });
}

function decimal(value) {
  const text = String(value ?? "").trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return null;
  const number = new Decimal(text);
  return number.isFinite() ? number : null;
}

function compareNumbers(group, fields) {
  if (group.status !== "paired") return [];
  return fields.map((field) => {
    const before = group.before[0][field],
      after = group.after[0][field];
    const a = decimal(before),
      b = decimal(after);
    let delta = null;
    if (a && b) {
      // Include both the integer and fractional digits, plus a possible carry.
      const precision =
        Math.max(1, a.e + 1, b.e + 1) + Math.max(a.dp(), b.dp()) + 1;
      const Exact = Decimal.clone({ precision });
      delta = new Exact(b).minus(a).toFixed();
    }
    return { field, before, after, delta };
  });
}

// Callers normalize their calendar to ISO dates. Every row must identify the same day.
function observationDate(rows, dateOf) {
  const dates = new Set(rows.map(dateOf));
  if (dates.size !== 1) return null;
  const [day] = dates;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || !Number.isFinite(Date.parse(day)))
    return null;
  return new Date(day).toISOString().startsWith(day) ? day : null;
}

module.exports = { matchRows, compareNumbers, observationDate };
