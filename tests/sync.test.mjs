import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import vm from 'node:vm';

function load(extras = {}) {
  const calls = [];
  const context = vm.createContext({ console, ...extras });
  vm.runInContext(readFileSync(new URL('../Code.gs', import.meta.url), 'utf8'), context);
  return { context, calls };
}

test('Sheets traverses empty pages with encoded cursors and deduplicates jobs', () => {
  const { context, calls } = load();
  const job = { id: 1, url: 'https://jobicy.com/jobs/1', jobTitle: '=IMPORTXML("bad")' };
  const pages = [
    { jobs: [job], nextCursor: 'opaque+/= token', hasMore: true },
    { jobs: [], nextCursor: 'last', hasMore: true },
    { jobs: [job, { id: 2, url: 'https://example.com/untrusted' }], nextCursor: null, hasMore: false }
  ];
  context.jobicyFetchJson_ = (url) => { calls.push(new URL(url)); return pages.shift(); };
  const jobs = context.jobicyFetchAll_({ geo: 'usa' }, Date.now());
  assert.equal(jobs.length, 1);
  assert.equal(calls.length, 3);
  assert.equal(calls[1].searchParams.get('cursor'), 'opaque+/= token');
  assert.equal(calls[2].searchParams.get('geo'), 'usa');
  assert.equal(context.jobicySafeCell_(job.jobTitle), '\'=IMPORTXML("bad")');
});

test('Sheets retains older rows and keeps unknown distinct from closed', () => {
  const { context } = load();
  const previous = [['1', 'Old role', '', '', '', '', '', '', '', '', 'https://jobicy.com/jobs/1', 'active', 'old', 'old']];
  const rows = context.jobicyMergeRows_(previous, [{ id: 2, url: 'https://jobicy.com/jobs/2' }], 'now');
  assert.equal(rows.length, 2);
  assert.equal(rows[0][11], 'active');
  assert.equal(rows[1][11], 'unknown');
  const statuses = context.jobicyValidateStatuses_({ success: true, count: 2, jobs: [{ id: 1, status: 'unknown' }, { id: 2, status: 'closed' }] }, ['1', '2']);
  assert.equal(statuses.get('1'), 'unknown');
  assert.equal(statuses.get('2'), 'closed');
  assert.throws(() => context.jobicyValidateStatuses_({ success: true, count: 2, jobs: [{ id: 1, status: 'active' }] }, ['1', '2']), /Incomplete/);
});

test('Sheets fails a cyclic traversal before any rows can be returned', () => {
  const { context } = load();
  context.jobicyFetchJson_ = () => ({ jobs: [], nextCursor: 'loop', hasMore: true });
  assert.throws(() => context.jobicyFetchAll_({}, Date.now()), /Repeated/);
});

test('Sheets releases the lock and performs no writes on a failed fetch', () => {
  let released = false;
  let wrote = false;
  const { context } = load({
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => { released = true; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => null, insertSheet: () => { wrote = true; } }) }
  });
  context.jobicyFetchJson_ = () => { throw new Error('HTTP 429'); };
  assert.throws(() => context.syncJobicyJobs(), /429/);
  assert.equal(released, true);
  assert.equal(wrote, false);
});
