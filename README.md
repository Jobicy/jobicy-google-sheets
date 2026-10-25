# Jobicy → Google Sheets

Import remote jobs into Google Sheets, refresh them hourly, and check whether stored listings are still open. Built with Google Apps Script; no Jobicy API key, external server, or paid automation service is required.

## Install in your spreadsheet

1. Create a Google spreadsheet. Choose **Extensions → Apps Script**.
2. Replace the default script with [`Code.gs`](Code.gs).
3. In Apps Script **Project Settings**, enable **Show appsscript.json manifest file in editor**. Replace that file with [`appsscript.json`](appsscript.json).
4. Save, select `syncJobicyJobs`, and click **Run**. Review Google's requested permissions and authorize your own script. It needs spreadsheet access, public HTTP requests, a custom menu, and permission to create its hourly trigger.
5. Open the **Jobicy Jobs** tab generated in your spreadsheet.
6. Reload the spreadsheet. Choose **Jobicy → Install hourly synchronization**. This stores the bound spreadsheet ID and installs one hourly trigger owned by your Google account.

To stop the schedule, choose **Jobicy → Remove synchronization trigger**. This removes only triggers for this integration's `syncJobicyJobs` function.

The generated tab contains ID, title, company, location, employment type, publication time, optional salary, original Jobicy URL, status, last-seen time, and status-check time. Put personal notes and formulas in a separate tab: this integration manages every cell in **Jobicy Jobs**. Leave its generated headers intact.

## Filters

Set optional values under Apps Script **Project Settings → Script properties**:

| Property | Example | Meaning |
| --- | --- | --- |
| `JOBICY_GEO` | `usa` | Candidate eligibility slug |
| `JOBICY_INDUSTRY` | `engineering` | Job category slug |
| `JOBICY_TAG` | `python` | Keyword, 3–50 characters |

Find current slugs using the official [locations](https://jobicy.com/api/jobs?get=locations) and [industries](https://jobicy.com/api/jobs?get=industries) endpoints. Changing filters affects future imports; previously stored rows remain. For a separate collection, use another spreadsheet and bound script.

## Synchronization behavior

- Each run reads all available cursor pages in the seven-day feed. Job IDs deduplicate and update existing rows.
- Every stored job receives a batch status check, including jobs older than the feed window. `active`, `closed`, and `unknown` remain distinct. Missing from the feed does not mean closed.
- Network, HTTP, malformed-response and incomplete-pagination errors abort before changing job rows. The next scheduled run starts from the first page.
- Successful feed synchronizations are at least one hour apart, including manual runs. Concurrent runs are blocked. Google controls the precise trigger time.
- There are safeguards of 100 feed pages, 10,000 retained rows and four minutes between requests. On reaching a safeguard, narrow the filters or archive the generated tab. Apps Script URL Fetch has its own request timeout and Google execution/URL-fetch quotas; this script cannot override them. No partial feed is deliberately committed.
- API text is written as text, with spreadsheet formula prefixes escaped. Descriptions and credentials are not stored.

Keep original listing links and show **[Jobs powered by Jobicy](https://jobicy.com/)** when sharing or displaying this data publicly. Code is MIT licensed; listing data has separate [fair-use terms](https://github.com/Jobicy/remote-jobs-api#fair-use).

## Local checks

```bash
node --test tests/*.test.mjs
```

These checks execute the script in a JavaScript sandbox with Apps Script adapters. Google authorization, menus, triggers and spreadsheet writes must be checked after installation in your Google account. No Google credentials are included.

[All integrations](https://github.com/Jobicy/jobicy-api-examples) · [API documentation](https://github.com/Jobicy/remote-jobs-api) · [License](LICENSE)
