---
title: "Handover brief: app help guides, 3 August 2026"
subtitle: "Upload this at the start of a new chat to carry the work forward"
date: 2026-08-03
lang: en
---

# Handover brief

Upload this file at the start of a new chat. It supersedes the brief of 2
August. Everything in it was learned by doing, and most of it was learned by
getting something wrong first.

## Who this is for and how to work

Jamal Mazrui is blind and reads everything with JAWS. Every deliverable follows
these without being asked:

- Start each reply with a line reading exactly `Claude:`.
- End each reply with a `## Summary` heading.
- Files at the **root** of a single zip. Never a nested zip.
- Prefer lists to tables.
- Code in **Camel Type**: Hungarian prefixes, lower camel case, functions never
  subprocedures, single-line if-then, for-each loops, alphabetised definition
  blocks, double quotes, constants prefixed `c_`.
- Markdown: UTF-8 **with BOM**, Windows line endings.
- Never a search-results URL in a deliverable. A link only if it has been seen
  to resolve.
- Build scripts always write a log.

**Read the version line first.** When he reports a fault, read the version and
path at the top of the log before diagnosing. Twice this has been a fault already
fixed in a build he was not running.

**He values a clean refusal over another attempt.** Say so when three good-faith
routes have failed, and write the refusal up with its evidence.

## The scope rule

His words, now in `ScopeOfTheseGuides.md`: official help for the **user or
customer** of a site. Not developers, not merchants selling through the site, not
third parties writing about it, and not the site's own content.

Two clauses matter most:

- **Viewing and navigating help is in scope if it is generic** — captions, text
  size, saving to read later, why a video will not play. "What happened in
  Tuesday's hearing" is not. This is what makes a broadcaster volume possible.
- **On a peer marketplace both sides are the user.** eBay and Facebook
  Marketplace: buying *and* selling help are user help, because the same person
  does both. The exclusion is aimed at professional sellers as a separate
  audience — Seller Centre, APIs, bulk tools, store subscriptions.

Where a product serves two audiences, take the one the volume is named for and
say so. Shopify's user *is* the merchant. Slack's admin help was kept because
someone running their own workspace is still a user.

## Twenty-five guides finished

GoldWave, Everything, Slack, Google Drive, Total Recorder, OneDrive, REAPER,
Libsyn, DuckDuckGo, Brave, OSARA, VLC, Dragon, foobar2000, RIM, Scribe, MyChart,
Venmo, Signal, Calm, GoodRx, Dropbox, Zoom (second edition), JetBlue, Walmart.

The largest are Zoom at 836 articles, Dropbox at 1,097, Slack at 501.

## Delivered and waiting to be run

Shopify v2 (browser), the REAPER Accessibility Wiki, Scribd, SlideShare, Google
Classroom, Google Voice, Google Translate, Credit Sesame v2 (slower), IMDb v2,
JetBlue v2, TuneIn, Stitch Fix, MSN, ESPN, CNN, eBay.

`probeEstates2.py` covers 55 products and is the cheapest next move on anything
uncertain.

## Refused, with evidence

Reuters (no reader documentation at all), Winamp (barely developed), WACUP (no
manual), Belarc Advisor (a product page and five questions), **Salesforce** (three
routes, all shells).

Salesforce is the important one because it is a new category: **a site that
refuses can still be gathered through its listing; a site that never renders
cannot be gathered at all.** Brave and Libsyn refused with status codes and
answered through their listings. Salesforce answers politely and the page never
becomes an article.

## The four publishing systems now known

- **Zendesk.** When the pages are walled, `/api/v2/help_center/en-us/articles.json`
  still answers and carries every article's body. This rescued Libsyn, Brave,
  Signal, Calm, GoodRx, Venmo and Credit Sesame. Always try it on a 403.
- **Freshdesk.** `/support/solutions/articles/<n>`, folders at
  `/support/solutions/folders/<n>`, the folder is the category. TuneIn, Stitch Fix.
- **MediaWiki.** `api.php` enumerates a namespace. VLC, and the REAPER wiki.
- **Salesforce Lightning.** Renders nothing for an automated browser. Refuse.

Plus: **Microsoft support** carries its category inside the page as
`ms.collection`; **Google support** keeps `co=` as the platform selector and files
most articles under no category at all, so the help home page is the only
arrangement.

## Chassis rules, each learned from a run that looked fine

- Undo HTML entities in an address **before** parsing it. This cost a whole Zoom run.
- Scan every `href`, not the last per anchor.
- Name which statuses count as a page. Dropbox answers 202 for many articles.
- **A readiness test based on length is always defeated.** Test for what the page
  must contain. Never write it as an `or` with a length check — I did, and saved
  seventeen shells.
- A shell is a page with almost no words *and* almost no links. That is
  "needs a browser", not "thin".
- One browser for the whole crawl, opened once.
- **Refuse against a written list, not a shape — or say the refusal out loud.**
  Refusing by shape silently dropped seven VLC modules named `aa`, `es`, `ts`.
  Dropbox now refuses language-shaped path segments and logs every one.
- A refusal entry must name a whole section, not a fragment. `/login` would have
  killed `/help/login-how-to`.
- **A file name cannot tell a slash from a hyphen.** True addresses come from
  `manifest.txt`. This was learned on DuckDuckGo and repeated on Dropbox.
- On a commercial site, fence **positively**. JetBlue's refusal list was useless
  against a site that invents a page per city pair: 515 of 598 pages were route
  advertising.
- Cap repeating log lines and tally the rest. Dropbox once wrote 27,937 identical
  refusals.
- One page raising an error must never end a crawl.
- Judge a damaged PDF by **density**, not by any single stray byte. Four
  occurrences in 30 MB is chance; 986,000 is a ruined file.

## Build rules

- Renumber headings by nesting depth, ignoring fenced code.
- Unwrap tables: single-row ones become paragraphs, the rest become lists.
  Pandoc's grid tables read aloud as punctuation.
- Strip surviving markup, outside code spans only.
- Resolve cross-references by the publisher's article identifier.
- A link whose visible text is an address gets the article's title instead, or
  becomes plain text. Enforce it once more over the finished document.
- Sections beat categories where a category holds seventy articles — true of
  Slack, Libsyn, Brave, Venmo, Calm and GoodRx.
- **Gates must never read inside fenced code**, and a litter list must name the
  furniture rather than condemn the words. Six times in one day a keyword list
  flagged a publisher's own prose: "still need help", "copyright", "related
  articles", "retrieved from", "have more questions", "https://" in a sentence
  about valid addresses. Every time the answer was to fix the check.

## Tools

- `probeEstates2.py` — 55 products, one run, an evidence-based verdict each.
- `shrinkHarvest.py` — strips a rendered harvest to a fraction without touching
  article text. Took Zoom from 455 MB to 85.
- `scopeAudit.py` — reports out-of-scope candidates and **never refuses**.
- `makeHarvesters.py` + `configs.py` — one chassis, one config per estate, so a
  rule learned once cannot be missing from another script.

## Summary

- Twenty-five guides finished; sixteen scripts delivered and waiting to run.
- Five products refused with written evidence, Salesforce being the instructive
  one.
- Four publishing systems are now known well enough to write a harvester without
  research.
- The scope rule and the chassis rules are both written down so they outlive any
  one conversation.
