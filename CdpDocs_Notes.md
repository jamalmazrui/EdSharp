---
title: "Notes on the CDP documentation volume"
date: 2026-08-11
lang: en
---

# Notes on the CDP documentation volume

## What arrived

- temp.zip held one folder, CdpDocs, with thirteen saved pages, an index page written by getCdpDocs.py, and getCdpDocs.log.
- The log records a run at 14:51 on 11 August 2026: thirteen curated sources, thirteen downloaded, none failed. The whole run took under two seconds, so nothing was rate limited or refused.
- Nothing else was in the archive. No app help harvest arrived with it.

## What was built

- CdpDocs.md, thirteen articles in five categories, about 17,700 words. Every gate green: byte order mark, Windows line endings, contents matching the articles, no bare addresses outside code, no broken internal link, no unbalanced code fence.
- Categories are alpha ordered with Miscellaneous last, and article titles are alpha ordered inside each category, ignoring a leading A, An or The.
- Each article opens with a line naming its publisher, because the sources are not all from the same hand.

## The gaps, stated rather than hidden

- The two official reference pages are shells. Their domain, command and event reference is drawn in the browser, so the saved copies hold only the domain list, which is 56 domain names with their stability markings. That list is worth having and is in the volume; the reference itself is not, and cannot be recovered from these files.
- The route to the real reference is in the note under both articles: a browser running with the remote debugging port open serves the whole protocol it speaks at localhost:9222/json/protocol, and the same definitions are published as browser_protocol.json and js_protocol.json in the ChromeDevTools devtools-protocol repository. That is the better source in any case, because it is the protocol your own Edge speaks rather than the tip of tree.
- The chrome-remote-interface article is thin at 109 words, and the reason is in the page: its recipes live on separate wiki pages listed in a sidebar, and only the wiki home page was downloaded. A second run seeded with the sidebar links would bring back the recipes themselves.
- The Reflect article printed its code with a line-number gutter beside each block. The builder drops a code block whose content is only digits, so the numbers are gone and the code is intact.

## The scope question

- The app help guides cover official help for the user or customer of a site, and exclude developer documentation. This volume is developer documentation from end to end, so it has not been folded into the collection and the README has not been touched. It stands on its own unless you say otherwise.
- Three of the thirteen sources are third parties writing about the protocol rather than its publisher: the Reflect article, the awesome-chrome-devtools list, and the chrome-remote-interface wiki. Each says so at its head.

## Where the pieces are

- buildCdpGuide.py takes a harvest folder and an output path, and writes buildCdpGuide.log beside itself.
- Rerunning it after a second harvest is a matter of adding the new files to the source list at the top of the script; each entry names the file, the title, the category, the publisher, the source address, and which extraction rule reads it.
