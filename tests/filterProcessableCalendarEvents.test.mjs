import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const helpersSource = fs.readFileSync(new URL('../Helpers.gs', import.meta.url), 'utf8');
const match = helpersSource.match(/function filterProcessableCalendarEvents\(events\)\{[\s\S]*?\n\}/);
if (!match) {
  throw new Error('Could not find filterProcessableCalendarEvents in Helpers.gs');
}

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(`${match[0]}; this.filterProcessableCalendarEvents = filterProcessableCalendarEvents;`, sandbox);

const filterProcessableCalendarEvents = sandbox.filterProcessableCalendarEvents;

assert.equal(filterProcessableCalendarEvents().length, 0, 'undefined input should return []');
assert.equal(filterProcessableCalendarEvents([]).length, 0, 'empty input should return []');

assert.deepEqual(
  filterProcessableCalendarEvents([
    null,
    { id: 'a', status: 'cancelled' },
    { id: 'b', status: 'CANCELLED' },
    { id: 'c', status: 'confirmed' },
    { id: 'd', status: 'tentative' },
    { id: 'e' }
  ]).map((e) => e.id),
  ['c', 'd', 'e'],
  'should remove null and cancelled events only'
);

console.log('ok - filterProcessableCalendarEvents');
