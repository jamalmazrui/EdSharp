# AudioVault Metadata Enrichment

## Purpose

This package contains the latest assets for building a screen-reader-friendly AudioVault directory and enriching it with IMDb, Wikidata, and Wikipedia metadata.

The final Markdown is intended for optional conversion to HTML with Pandoc.


## Version 6: Resume After a Reboot or Forced Stop

Version 6 adds automatic **log bootstrapping** specifically for interrupted older runs that had a cache and detailed log but did not yet have `AudioVault_Wikimedia_Results.json`.

The latest analyzed run reached 548 fully completed unique productions and had begun title 549 when the computer rebooted. Its log contained 352 `MATCH` decisions and 196 `NO MATCH` decisions. The version-6 parser was tested against that complete uploaded log and recovered all 548 decisions.

When you run `Resume_AudioVault_Wikimedia.cmd`, the program now does the following before making new Internet requests:

1. Reads the existing `AudioVault_Wikimedia_Enrichment.log`.
2. Pairs each prior `TITLE START` with its final `MATCH` or `NO MATCH` line.
3. Reconstructs the exact title/year/Movies-or-TV state key.
4. For prior matches, obtains the selected Wikidata Q-ID and Wikipedia article name from the log.
5. Rebuilds metadata from the existing Wikimedia cache in **cache-only mode**, so bootstrapping itself never generates new network traffic.
6. Immediately saves recovered records into `AudioVault_Wikimedia_Results.json`.
7. Defers any prior match whose required cached data is missing; that title is safely handled through the normal online path later.
8. Skips every title now marked complete in the state file.

This means a reboot should no longer cause hundreds of completed titles to be replayed. After the first version-6 resume, `AudioVault_Wikimedia_Results.json` becomes the primary per-title checkpoint file and is updated after each completed production.

The resume command is:

```powershell
Resume_AudioVault_Wikimedia.cmd
```

Equivalent manual command:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md --wikimedia-delay 7.5 --cache AudioVault_Wikimedia_Cache.json --state AudioVault_Wikimedia_Results.json --log AudioVault_Wikimedia_Enrichment.log --bootstrap-log AudioVault_Wikimedia_Enrichment.log
```

Normally you do not need to specify `--bootstrap-log`: unless `--no-log-bootstrap` is used, the program automatically uses the current `--log` file as its bootstrap source. The explicit option is included in the resume launcher so its purpose is obvious.

### Files to preserve between runs

Do not delete these files while enrichment is unfinished:

- `AudioVault_Directory_IMDb.md`
- `AudioVault_Wikimedia_Cache.json`
- `AudioVault_Wikimedia_Results.json`
- `AudioVault_Wikimedia_Enrichment.log`
- `AudioVault_Directory_Enriched.partial.md`, when present

The cache and state files use atomic temporary-file replacement to reduce the chance that a sudden reboot leaves a half-written JSON file.

## Package Contents

### `AudioVault_Directory.md`

The current screen-reader-friendly AudioVault directory.

### `enrich_audiovault_imdb.py`

The first enrichment pass. It downloads IMDb's official non-commercial datasets and locally adds metadata such as:

- End year
- Genres
- IMDb
- IMDb rating
- IMDb rating votes
- Running time
- Type

This pass uses IMDb's bulk data locally after download, so it does not make thousands of per-title IMDb web requests.

### `enrich_audiovault.py`

The second enrichment pass. It queries Wikidata and English Wikipedia for metadata not already available from the IMDb pass, including:

- Country
- Creator
- Description
- Director
- Wikipedia

It preserves existing metadata, keeps fields alphabetized, caches web responses, and can be interrupted and resumed.

### `Run_AudioVault_Enrichment.cmd`

Windows Command Prompt launcher that runs both passes.

### `Run_AudioVault_Enrichment.ps1`

PowerShell launcher that runs both passes.

## What Has Been Done

1. Consolidated the captured AudioVault HTML pages.
2. Removed navigation, filters, repeated page blocks, footers, and other non-content material.
3. Separated Movies and TV Shows under level-two headings.
4. Made each AudioVault entry a level-three heading with its title linked directly to the audio download.
5. Kept the release year at the end of every heading in parentheses.
6. Sorted entries alphabetically while ignoring an initial A, An, or The.
7. Removed entries explicitly labeled as non-English.
8. Preserved useful AudioVault-specific metadata.
9. Added an IMDb bulk-data enrichment pass.
10. Added a Wikidata/Wikipedia enrichment pass.
11. Added persistent caching and detailed diagnostic logging.
12. Revised Wikimedia throttling after analysis of an actual 5,409-line rate-limit log.

## What the Rate-Limit Log Showed

The earlier Wikimedia defaults were too aggressive.

The log showed:

- 1,255 Wikidata requests.
- 140 Wikidata HTTP 429 responses.
- 309 Wikipedia requests.
- 24 Wikipedia HTTP 429 responses.
- Wikidata rejections typically occurred after only 8 to 10 Wikidata requests.
- The server then commonly required about 43 to 48 seconds of backoff.
- Wikidata and Wikipedia were sometimes rate-limited at essentially the same time.

This behavior is consistent with a shared Wikimedia-side throttle affecting both domains for this client. Treating Wikidata and Wikipedia as independent servers therefore did not help; overlapping them could make the combined burst worse.

## New Wikimedia Throttling Strategy

Version 6 of `enrich_audiovault.py` replaces independent host throttles with a single shared Wikimedia scheduler.

### Initial pacing

The default is now:

```text
7.5 seconds between all uncached Wikimedia requests
```

This is intentionally conservative. It corresponds to about eight requests per minute before considering response time.

Use:

```powershell
--wikimedia-delay 7.5
```

### Adaptive rate control

The program now adjusts itself when the server signals overload.

After an HTTP 429 it:

1. Honors the server's `Retry-After` value when present.
2. Examines how many Wikimedia requests were made during the previous rolling 60 seconds.
3. Increases the shared delay by at least 25 percent.
4. Calculates an additional safe interval from the observed rolling-minute request count.
5. Adds a small random margin to the cooldown.
6. Does not relax the delay until 40 consecutive uncached requests have succeeded.
7. When relaxing, reduces the interval by only 0.5 seconds at a time and never below the configured base delay.

The log records every throttle increase and relaxation.

### Why multithreading was removed from the Wikimedia pass

The previous design used one Wikidata thread and one Wikipedia thread. That was reasonable if the services had independent limits.

The uploaded log showed that they can be rejected together. Therefore the current version deliberately does **not** overlap Wikidata and Wikipedia requests.

Concurrency is useful only when the remote systems are genuinely independent. Here, one shared Wikimedia scheduler is both more polite and more efficient because it avoids synchronized 429s and long forced pauses.

The IMDb pass still gains its speed from bulk datasets and local processing rather than repeated remote lookups.

## More Efficient Wikidata Caching

The earlier cache stored whole request URLs. That meant a common Wikidata entity such as the United States, a film type, or a common genre could be downloaded repeatedly whenever it appeared in a different entity batch.

Version 6 retains the second cache indexed by individual Wikidata Q-ID.

For example, once `Q30` has been downloaded, it can be reused regardless of which later title references it.

This should greatly reduce:

- Total Wikidata requests
- Multi-megabyte responses
- Repeated downloads of common countries, genres, and production types
- Time spent waiting between requests

Existing version-1 cache files are upgraded automatically. Their already downloaded URL responses are preserved.

## Avoiding Redundant Metadata Requests

The Wikimedia pass now inspects metadata already present from the IMDb pass.

If IMDb has already supplied fields such as:

- Genres
- IMDb
- Running time
- Type

the Wikimedia pass does not perform extra label lookups merely to reconstruct those same fields.

It concentrates on metadata still missing, especially Country, Creator, Director, Description, and Wikipedia.

## Diagnostic Log

The Wikimedia pass writes:

```text
AudioVault_Wikimedia_Enrichment.log
```

The log is append-only and records:

- Python version
- Input and output paths
- Cache loading and upgrades
- URL cache hits
- Individual Wikidata entity-cache hits
- Each uncached request
- Current adaptive delay
- Response status, duration, and byte count
- HTTP 429 and temporary server errors
- Retry-After values and backoff
- Rolling-minute throttle calculations
- Throttle increases and relaxations
- Match scores
- Wikipedia descriptions
- Progress counts
- Cached entity counts
- Final totals
- Exception tracebacks

For future debugging, upload these together when possible:

1. `AudioVault_Wikimedia_Enrichment.log`
2. `AudioVault_Wikimedia_Cache.json`

The cache is especially useful because it shows what work was already completed and lets a corrected program resume without throwing that work away.

## Requirements

- Windows 11 or another current Windows version
- Python 3 available through the `py` launcher
- Internet connection
- No third-party Python packages are required

Check Python with:

```powershell
py --version
```

## Simplest Way to Run Everything

Extract all files into one folder and double-click:

```text
Run_AudioVault_Enrichment.cmd
```

The process creates:

```text
AudioVault_Directory_IMDb.md
AudioVault_Directory_Enriched.md
AudioVault_Wikimedia_Cache.json
AudioVault_Wikimedia_Enrichment.log
IMDb_Data\
```

## Run the IMDb Pass Manually

```powershell
py enrich_audiovault_imdb.py AudioVault_Directory.md AudioVault_Directory_IMDb.md
```

To use a different data folder:

```powershell
py enrich_audiovault_imdb.py AudioVault_Directory.md AudioVault_Directory_IMDb.md --data-dir MyIMDbData
```

## Run the Wikimedia Pass Manually

Recommended command:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md --wikimedia-delay 7.5 --cache AudioVault_Wikimedia_Cache.json --log AudioVault_Wikimedia_Enrichment.log
```

The shorter command uses the same 7.5-second default:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md
```

### Wikimedia options


`--bootstrap-log`

Reads an older enrichment log and reconstructs missing per-title resume records from completed `MATCH` and `NO MATCH` decisions. Reconstructed metadata is taken from the existing cache without making Internet requests during bootstrapping.

`--no-log-bootstrap`

Disables automatic log reconstruction. This is mainly a troubleshooting option.

`--state`

Sets the small per-title resume file. Default: `AudioVault_Wikimedia_Results.json`. It is updated after every completed unique production.

`--wikimedia-delay`

Sets the starting global delay between all uncached Wikidata and Wikipedia requests. Default: 7.5 seconds.

`--cache`

Sets the persistent cache file.

`--log`

Sets the detailed append-only log file.

`--timeout`

Sets the network timeout in seconds. Default: 45.

`--retries`

Sets the maximum number of attempts for retryable requests. Default: 8.

`--verbose`

Prints detailed request and throttle messages to the console in addition to writing them to the log.

The older options `--wikidata-delay`, `--wikipedia-delay`, and `--delay` remain accepted for compatibility, but they can only make the shared delay more conservative. The program no longer gives Wikidata and Wikipedia independent request clocks.

## Stopping and Resuming

It is safe to stop the Wikimedia pass with Ctrl+C.

The program saves:

```text
AudioVault_Wikimedia_Cache.json
```

Run the same command again to resume. Cached URL responses and cached individual Wikidata entities are reused.

Do not delete the cache unless you intentionally want to start external lookups from scratch.

## If Rate Limiting Happens Again

Do not immediately delete the cache or restart from scratch.

Upload:

```text
AudioVault_Wikimedia_Enrichment.log
AudioVault_Wikimedia_Cache.json
```

The new log lines beginning with:

```text
THROTTLE INCREASE
THROTTLE RELAX
GLOBAL POLITE WAIT
ENTITY CACHE HIT
```

make it possible to see whether the adaptive scheduler is converging on a stable interval.

If necessary, start more conservatively:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md --wikimedia-delay 10
```

or:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md --wikimedia-delay 12
```

Because of the adaptive logic, a 429 will also raise the interval automatically.

## Converting the Final Markdown with Pandoc

To preserve the manually constructed table of contents:

```powershell
pandoc AudioVault_Directory_Enriched.md --standalone --output AudioVault_Directory.htm
```

Using Pandoc's `--toc` switch may create a second table of contents because the Markdown already contains one.

## Matching Philosophy

The scripts prefer missing metadata over a false match.

Matching considers:

- Normalized title
- Release year
- Movies versus TV Shows
- Wikidata production type
- Candidate-title similarity

Separate AudioVault variants such as UK, US, and TTS remain separate entries but reuse metadata for the same underlying production.

## Suggested Debugging Workflow

If the program behaves unexpectedly:

1. Stop with Ctrl+C.
2. Keep the cache.
3. Keep the log.
4. Upload both files.
5. Also upload `AudioVault_Directory_IMDb.md` if the issue appears to involve title matching rather than networking.

That combination provides enough information to diagnose throttling, matching, caching, and output-generation problems without repeating already completed network work.

## Update after the August 6 completed-run diagnostic

The later uploaded log showed that the new shared 7.5-second Wikimedia scheduler solved the earlier rate-limit problem: the analyzed run made 334 logged HTTP requests without a single logged HTTP 429 or throttle increase. However, the run did **not** reach normal completion. It stopped after starting production 115 of 2,703 (`54` (1998)); the last log record was a global polite wait before the next Wikidata request. There was no `COMPLETE` record and no logged exception.

Because the prior program only wrote the final Markdown after all 2,703 productions were processed, a process termination at that point could leave no newly completed final output even though substantial cache work had been done. Version 5 addresses that failure mode.

### Version 5 resilience changes

- Adds `AudioVault_Wikimedia_Results.json`, a small per-production resume-state file. Every successfully matched or definitively unmatched production is saved immediately. On restart, a `RESULT CACHE HIT` skips that production completely without redoing Wikimedia matching.
- Writes `AudioVault_Directory_Enriched.partial.md` every 100 processed productions and whenever the program catches an interruption or ordinary fatal exception. The partial file contains all metadata completed so far plus the original/IMDb metadata for titles not yet processed.
- Logs `FATAL ERROR` / `UNHANDLED FATAL ERROR` for ordinary unexpected exceptions. This makes future debugging much more informative than a log that simply ends.
- Keeps the existing shared Wikimedia scheduler and 7.5-second default because the completed-run log showed zero 429 responses at that setting.
- Migrates the old cache to version 3. Full `wbgetentities` responses are converted to slim per-QID records and the duplicate giant response blobs are discarded. In testing, the uploaded 156 MB cache shrank to about 1.7 MB while retaining the fields needed by this program.
- Related entity names (country, director, creator, etc.) are now fetched with `props=labels` rather than downloading all claims and references for those entities.
- Full entity API responses are no longer stored twice in both the URL-response cache and Q-ID cache.
- The large network cache is saved less frequently because the small resume-state file now protects completed production work after every title.

### Recommended resume procedure

Keep these files together in the same folder:

- `AudioVault_Directory_IMDb.md`
- `AudioVault_Wikimedia_Cache.json` — use the cache from the previous run; version 5 will migrate and shrink it automatically.
- `enrich_audiovault.py`
- `Resume_AudioVault_Wikimedia.cmd` (or the PowerShell equivalent)

Then double-click `Resume_AudioVault_Wikimedia.cmd`, or run:

```powershell
py enrich_audiovault.py AudioVault_Directory_IMDb.md AudioVault_Directory_Enriched.md --wikimedia-delay 7.5 --cache AudioVault_Wikimedia_Cache.json --state AudioVault_Wikimedia_Results.json --log AudioVault_Wikimedia_Enrichment.log
```

On the first version-5 run, the existing network cache is migrated. The first titles should then pass quickly using cached search/entity/Wikipedia responses. New network requests resume where the cache no longer contains the needed data. From that point forward, `AudioVault_Wikimedia_Results.json` provides direct title-level resumption.

### Files to upload for future debugging

If a later run stops unexpectedly, upload these three files:

1. `AudioVault_Wikimedia_Enrichment.log`
2. `AudioVault_Wikimedia_Results.json`
3. `AudioVault_Wikimedia_Cache.json`

Also upload `AudioVault_Directory_Enriched.partial.md` if you want the partial generated directory evaluated. The most important diagnostic marker for a fully successful run is a final log line beginning with `COMPLETE output=` and reporting `stateCompleted=2703`.
