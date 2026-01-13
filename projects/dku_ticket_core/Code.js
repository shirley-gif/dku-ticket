/**
 * DKU Ticket Core – Shared Logic
 * SAFE TO SHARE
 */

function calcPriority(impact, urgency) {
  if (impact === 'High' && urgency === 'High') return 'P1';
  if (impact === 'High' || urgency === 'High') return 'P2';
  if (impact === 'Medium' && urgency === 'Medium') return 'P3';
  return 'P4';
}

function calcDueAt(start, priority) {
  const days = { P1: 3, P2: 5, P3: 7, P4: 10 }[priority] || 10;
  const d = new Date(start);
  d.setDate(d.getDate() + days);
  return d;
}

const STATUS_RULES = {
  'New': ['In Progress', 'Closed'],
  'In Progress': ['Waiting on User', 'Waiting on IT/Vendor', 'Resolved', 'Closed'],
  'Waiting on User': ['In Progress', 'Resolved', 'Closed'],
  'Waiting on IT/Vendor': ['In Progress', 'Resolved', 'Closed'],
  'Resolved': ['Closed', 'In Progress'],
  'Closed': []
};

const STATUS_EXPLANATIONS = {
  'New': 'Your request has been received and is queued for review.',
  'In Progress': 'We are currently working on your request.',
  'Waiting on User': 'We are waiting for additional information from you.',
  'Waiting on IT/Vendor': 'We are coordinating with IT or the vendor.',
  'Resolved': 'The issue has been resolved. Please let us know if it persists.',
  'Closed': 'This ticket is now closed. Thank you.'
};

function isValidStatus(status) {
  return Object.prototype.hasOwnProperty.call(STATUS_RULES, String(status || '').trim());
}

function isValidTransition(fromStatus, toStatus) {
  const fromKey = String(fromStatus || '').trim();
  const toKey = String(toStatus || '').trim();
  if (!isValidStatus(fromKey) || !isValidStatus(toKey)) return false;
  return STATUS_RULES[fromKey].includes(toKey);
}

function statusExplanation(status) {
  const key = String(status || '').trim();
  return STATUS_EXPLANATIONS[key] || '';
}

function testStatusTransitionRules() {
  const allowed = [
    ['New', 'In Progress'],
    ['New', 'Closed'],
    ['In Progress', 'Waiting on User'],
    ['In Progress', 'Waiting on IT/Vendor'],
    ['In Progress', 'Resolved'],
    ['In Progress', 'Closed'],
    ['Waiting on User', 'In Progress'],
    ['Waiting on User', 'Resolved'],
    ['Waiting on User', 'Closed'],
    ['Waiting on IT/Vendor', 'In Progress'],
    ['Waiting on IT/Vendor', 'Resolved'],
    ['Waiting on IT/Vendor', 'Closed'],
    ['Resolved', 'Closed'],
    ['Resolved', 'In Progress']
  ];
  for (const [fromStatus, toStatus] of allowed) {
    if (!isValidTransition(fromStatus, toStatus)) {
      throw new Error(`Expected valid transition: ${fromStatus} -> ${toStatus}`);
    }
  }

  const invalid = [
    ['Closed', 'New'],
    ['Closed', 'In Progress'],
    ['New', 'Resolved'],
    ['Waiting on User', 'Waiting on IT/Vendor']
  ];
  for (const [fromStatus, toStatus] of invalid) {
    if (isValidTransition(fromStatus, toStatus)) {
      throw new Error(`Expected invalid transition: ${fromStatus} -> ${toStatus}`);
    }
  }

  return { ok: true };
}
