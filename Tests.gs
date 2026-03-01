/**
 * Lightweight regression test for issue #315.
 *
 * Run manually in Apps Script editor.
 */
function test_filterCancelledEvents_regression315() {
  var calendarEvents = [
    { id: '1', status: 'confirmed' },
    { id: '2', status: 'cancelled' },
    { id: '3' }, // missing status should be kept
    { id: '4', status: 'tentative' }
  ];

  var filtered = calendarEvents.filter(function(e) {
    return e.status !== 'cancelled';
  });

  if (filtered.length !== 3) {
    throw new Error('Expected 3 events after filtering cancelled events, got ' + filtered.length);
  }

  var ids = filtered.map(function(e) { return e.id; }).join(',');
  if (ids !== '1,3,4') {
    throw new Error('Unexpected filtered IDs: ' + ids);
  }

  Logger.log('PASS test_filterCancelledEvents_regression315');
}
