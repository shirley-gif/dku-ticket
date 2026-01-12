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
