---
title: "Zotero Help Guide"
subtitle: "Zotero's official documentation, gathered into one screen-reader-friendly document"
date: 2026-07-30
lang: en
---

# Zotero Help Guide

This guide gathers Zotero's official documentation. Zotero is the free, open-source reference manager: it collects citations from what you are reading, keeps them organised with notes and attached files, and inserts formatted citations and bibliographies into your writing in whichever style a publisher asks for. It was taken from Zotero's own documentation site on 30 July 2026.

Zotero reaches more of the surfaces these guides look for than most products: a program for Windows, Mac and Linux, connectors for Chrome, Edge, Firefox and Safari, a web library at zotero.org, and an app for iPhone and iPad. There is no Alexa skill.

The word processor sections are the ones that matter most when you are actually writing — the Word, LibreOffice and Google Docs plugins, citation styles, and how to build a bibliography — and they sit in the citing section below.

The sections below are mine rather than Zotero's. Zotero's documentation is a wiki with a navigation tree that does not survive being taken out of the site, so the pages are grouped here by subject and that choice is my own.

Two things were left out. The developer documentation, which covers the web API and writing translators — developer material under your standing rule. And the forums, which are users answering users rather than Zotero's own guidance.

Headings run three deep: this title, then a section, then a page, with each page's own subheadings nested beneath. Every page is listed in the table of contents below.

## Contents {#contents}

- [Citing and bibliographies](#citing-and-bibliographies)
    - [Can I prevent the “Add Citation” dialog from the word processor plugins from moving behind the word processor window?](#can-i-prevent-the-add-citation-dialog-from-the-word-processor-plugins-from-movin)
    - [Citation Styles](#citation-styles)
    - [Creating Bibliographies](#creating-bibliographies)
    - [Does Zotero support label/authorship trigraph styles, like [ddb98]?](#does-zotero-support-labelauthorship-trigraph-styles-like-ddb98)
    - [How do I get titles to show up in sentence case in bibliographies?](#how-do-i-get-titles-to-show-up-in-sentence-case-in-bibliographies)
    - [How do I manually add a bibliographic item?](#how-do-i-manually-add-a-bibliographic-item)
    - [How do I prevent title casing of non-English titles in bibliographies?](#how-do-i-prevent-title-casing-of-non-english-titles-in-bibliographies)
    - [How do I unlink all Zotero citations in a document?](#how-do-i-unlink-all-zotero-citations-in-a-document)
    - [How do I use rich text formatting, like italics and sub/superscript, in titles?](#how-do-i-use-rich-text-formatting-like-italics-and-subsuperscript-in-titles)
    - [How do you cite a secondary source in Zotero?](#how-do-you-cite-a-secondary-source-in-zotero)
    - [I need to use Chicago style. Which of the three versions that come with Zotero should I use?](#i-need-to-use-chicago-style-which-of-the-three-versions-that-come-with-zotero-sh)
    - [I’m the publisher/editor of a journal. What can I do to have Zotero support our style?](#im-the-publishereditor-of-a-journal-what-can-i-do-to-have-zotero-support-our-sty)
    - [Legal Citations: Juris-M](#legal-citations-juris-m)
    - [Missing Italics (or Italics-Only) in Word Bibliographies](#missing-italics-or-italics-only-in-word-bibliographies)
    - [Moving Documents with Zotero Citations Between Word Processors](#moving-documents-with-zotero-citations-between-word-processors)
    - [Plugins for Zotero](#plugins-for-zotero)
    - [References appear in the wrong font in Word/LibreOffice](#references-appear-in-the-wrong-font-in-wordlibreoffice)
    - [RTF Scan](#rtf-scan)
    - [Standard Citation Styles](#standard-citation-styles)
    - [Troubleshooting Errors in Word Processor Documents](#troubleshooting-errors-in-word-processor-documents)
    - [Using the Zotero LibreOffice Plugin](#using-the-zotero-libreoffice-plugin)
    - [Using the Zotero Word Plugin](#using-the-zotero-word-plugin)
    - [Using Zotero with Google Docs](#using-zotero-with-google-docs)
    - [What happened to the “classic” citation dialog?](#what-happened-to-the-classic-citation-dialog)
    - [What is the official Harvard style?](#what-is-the-official-harvard-style)
    - [Where is the Zotero toolbar in Word for Mac 2008?](#where-is-the-zotero-toolbar-in-word-for-mac-2008)
    - [Why are my citations underlined with a dashed line?](#why-are-my-citations-underlined-with-a-dashed-line)
    - [Why are Zotero citations or bibliographies always highlighted in gray or another color?](#why-are-zotero-citations-or-bibliographies-always-highlighted-in-gray-or-another)
    - [Why do I see code beginning with ADDIN ZOTERO_ITEM CSL_CITATION in my document instead of formatted citations?](#why-do-i-see-code-beginning-with-addin-zoteroitem-cslcitation-in-my-document-ins)
    - [Why do some citations include first names or initials?](#why-do-some-citations-include-first-names-or-initials)
    - [Why is a citation not updated in my document after editing the item in Zotero?](#why-is-a-citation-not-updated-in-my-document-after-editing-the-item-in-zotero)
    - [Why is Zotero slow to insert citations or update the bibliography?](#why-is-zotero-slow-to-insert-citations-or-update-the-bibliography)
    - [Why isn’t the first letter of a subtitle in uppercase in bibliographies?](#why-isnt-the-first-letter-of-a-subtitle-in-uppercase-in-bibliographies)
    - [Why isn’t Zotero detecting my existing citations?](#why-isnt-zotero-detecting-my-existing-citations)
    - [Why isn’t Zotero detecting my existing Google Docs citations?](#why-isnt-zotero-detecting-my-existing-google-docs-citations)
    - [Word Processor Plugin Shortcuts](#word-processor-plugin-shortcuts)
    - [Word Processor Plugins](#word-processor-plugins)
    - [Zotero and Word Compatibility on Apple Silicon Macs](#zotero-and-word-compatibility-on-apple-silicon-macs)
    - [Zotero does not have permission to control Word](#zotero-does-not-have-permission-to-control-word)
    - [Zotero Word Processor Plugin Troubleshooting](#zotero-word-processor-plugin-troubleshooting)
    - [“Could not find a running Word instance”](#could-not-find-a-running-word-instance)
- [Collecting references](#collecting-references)
    - [Adding Files to your Zotero Library](#adding-files-to-your-zotero-library)
    - [Adding Items to Zotero](#adding-items-to-zotero)
    - [Can I highlight and annotate PDFs with Zotero?](#can-i-highlight-and-annotate-pdfs-with-zotero)
    - [Default translators](#default-translators)
    - [DOI format in APA style](#doi-format-in-apa-style)
    - [How can I import from Citavi?](#how-can-i-import-from-citavi)
    - [How can I quickly switch between Zotero and my browser, PDF viewer, and/or word processor?](#how-can-i-quickly-switch-between-zotero-and-my-browser-pdf-viewer-andor-word-pro)
    - [How do I import a Mendeley library into Zotero?](#how-do-i-import-a-mendeley-library-into-zotero)
    - [How do I import BibTeX or other standardized formats?](#how-do-i-import-bibtex-or-other-standardized-formats)
    - [How do I import from EndNote?](#how-do-i-import-from-endnote)
    - [How do I import references into Zotero?](#how-do-i-import-references-into-zotero)
    - [How do I turn off automatic case changes/capitalization during item import?](#how-do-i-turn-off-automatic-case-changescapitalization-during-item-import)
    - [How does the Import from Clipboard feature work?](#how-does-the-import-from-clipboard-feature-work)
    - [I have bibliographies in Microsoft Word documents, PDFs, and other text files. Can I import them into my Zotero library?](#i-have-bibliographies-in-microsoft-word-documents-pdfs-and-other-text-files-can-)
    - [Importing a Mendeley Library Into Zotero (Alternative Local Method)](#importing-a-mendeley-library-into-zotero-alternative-local-method)
    - [Importing from Other Reference Managers](#importing-from-other-reference-managers)
    - [Is the Zotero web library the same as the Zotero desktop app?](#is-the-zotero-web-library-the-same-as-the-zotero-desktop-app)
    - [Known Translator Issues](#known-translator-issues)
    - [Proxies](#proxies)
    - [Retrieve PDF Metadata](#retrieve-pdf-metadata)
    - [Troubleshooting Problems Saving to Zotero](#troubleshooting-problems-saving-to-zotero)
    - [What are these DOIs doing in my bibliography?](#what-are-these-dois-doing-in-my-bibliography)
    - [Why can’t I access a proxied site when the Zotero Connector is enabled?](#why-cant-i-access-a-proxied-site-when-the-zotero-connector-is-enabled)
    - [Why do attachments have names like “PDF” or “Accepted Version” instead of their filenames in the items list?](#why-do-attachments-have-names-like-pdf-or-accepted-version-instead-of-their-file)
    - [Why do I see highlighted text twice in the PDF reader or in notes created from annotations?](#why-do-i-see-highlighted-text-twice-in-the-pdf-reader-or-in-notes-created-from-a)
    - [Why does Zotero store PDF annotations in its database instead of in the PDF file?](#why-does-zotero-store-pdf-annotations-in-its-database-instead-of-in-the-pdf-file)
    - [Why doesn’t the Zotero Connector offer to save complete data from a webpage?](#why-doesnt-the-zotero-connector-offer-to-save-complete-data-from-a-webpage)
    - [Why is my browser saying the Zotero Connector needs access to my data on all websites?](#why-is-my-browser-saying-the-zotero-connector-needs-access-to-my-data-on-all-web)
    - [Zotero Connector](#zotero-connector)
    - [Zotero Connector and Safari](#zotero-connector-and-safari)
    - [Zotero Connector Preferences](#zotero-connector-preferences)
    - [Zotero Connector: “Is Zotero Running?”](#zotero-connector-is-zotero-running)
    - [Zotero PDF Reader and Note Editor](#zotero-pdf-reader-and-note-editor)
    - [Zotero Translators](#zotero-translators)
- [Getting started](#getting-started)
    - [The Basics](#the-basics)
    - [Can I still use Zotero if I can’t install programs on my computer?](#can-i-still-use-zotero-if-i-cant-install-programs-on-my-computer)
    - [Does Zotero offer installable packages of Zotero for specific Linux distributions?](#does-zotero-offer-installable-packages-of-zotero-for-specific-linux-distribution)
    - [How Do I Install Zotero on a Chromebook?](#how-do-i-install-zotero-on-a-chromebook)
    - [How do I uninstall the Zotero word processor plugins?](#how-do-i-uninstall-the-zotero-word-processor-plugins)
    - [How do I uninstall Zotero?](#how-do-i-uninstall-zotero)
    - [Installation Instructions](#installation-instructions)
    - [Installing the Zotero Word Processor Plugins](#installing-the-zotero-word-processor-plugins)
    - [Manually Installing the Zotero Word Processor Plugin](#manually-installing-the-zotero-word-processor-plugin)
    - [Troubleshooting Errors with Word Processor Plugin Installation](#troubleshooting-errors-with-word-processor-plugin-installation)
    - [Why am I getting a “disk I/O error” at Zotero startup?](#why-am-i-getting-a-disk-io-error-at-zotero-startup)
- [Groups and collaboration](#groups-and-collaboration)
    - [Sharing data directory between Zotero Standalone and Zotero for Firefox](#sharing-data-directory-between-zotero-standalone-and-zotero-for-firefox)
- [Organising your library](#organising-your-library)
    - [Annotation](#annotation)
    - [Collections and Tags](#collections-and-tags)
    - [Do annotations sync?](#do-annotations-sync)
    - [Duplicate Detection](#duplicate-detection)
    - [How can I access my library from multiple computers? Can I store my Zotero library and associated files on an external drive?](#how-can-i-access-my-library-from-multiple-computers-can-i-store-my-zotero-librar)
    - [How can I move my Zotero library to a different computer?](#how-can-i-move-my-zotero-library-to-a-different-computer)
    - [How can I see how many items I have in my Zotero library?](#how-can-i-see-how-many-items-i-have-in-my-zotero-library)
    - [How can I see what collections my item is in?](#how-can-i-see-what-collections-my-item-is-in)
    - [How do I export my Zotero library?](#how-do-i-export-my-zotero-library)
    - [How do I get my Zotero collection to work with an Exhibit or Citeline presentation?](#how-do-i-get-my-zotero-collection-to-work-with-an-exhibit-or-citeline-presentati)
    - [How do I merge tags?](#how-do-i-merge-tags)
    - [How do I print my Zotero references and notes?](#how-do-i-print-my-zotero-references-and-notes)
    - [How do I sort my notes by page number?](#how-do-i-sort-my-notes-by-page-number)
    - [Library Item Overview](#library-item-overview)
    - [Note Templates](#note-templates)
    - [Notes](#notes)
    - [Related Items](#related-items)
    - [Search](#search)
    - [Searching](#searching)
    - [Sorting](#sorting)
    - [What ways can I organize and manage my collections?](#what-ways-can-i-organize-and-manage-my-collections)
    - [Where did the Extract Annotations button go?](#where-did-the-extract-annotations-button-go)
- [Other help pages](#other-help-pages)
    - [About Zotero](#about-zotero)
    - [Can I open snapshots in a new tab or window? (Zotero for Firefox)](#can-i-open-snapshots-in-a-new-tab-or-window-zotero-for-firefox)
    - [Contact Zotero](#contact-zotero)
    - [Does Zotero support non-Western characters?](#does-zotero-support-non-western-characters)
    - [Feeds](#feeds)
    - [File Handling Issues](#file-handling-issues)
    - [File Renaming](#file-renaming)
    - [Forum Guidelines](#forum-guidelines)
    - [Frequently Asked Questions](#frequently-asked-questions)
    - [Getting Help](#getting-help)
    - [Google Scholar (or some other site) locked me out after using Zotero to save items. What happened?](#google-scholar-or-some-other-site-locked-me-out-after-using-zotero-to-save-items)
    - [How can I open multiple instances of Zotero and save to or cite from a specific instance?](#how-can-i-open-multiple-instances-of-zotero-and-save-to-or-cite-from-a-specific-)
    - [How can I switch the Zotero desktop app to a different Zotero account?](#how-can-i-switch-the-zotero-desktop-app-to-a-different-zotero-account)
    - [How can I use Zotero with a SOCKS proxy?](#how-can-i-use-zotero-with-a-socks-proxy)
    - [How can subsequent occurences of the same author replaced by a fixed term/symbol?](#how-can-subsequent-occurences-of-the-same-author-replaced-by-a-fixed-termsymbol)
    - [How do I add a letter or memo?](#how-do-i-add-a-letter-or-memo)
    - [How do I add an archival or other unpublished source?](#how-do-i-add-an-archival-or-other-unpublished-source)
    - [How do I add an edited volume or a book chapter?](#how-do-i-add-an-edited-volume-or-a-book-chapter)
    - [How do I attach a file or web page to an item?](#how-do-i-attach-a-file-or-web-page-to-an-item)
    - [How do I change the font size of text in Zotero?](#how-do-i-change-the-font-size-of-text-in-zotero)
    - [How do I change the size of my Zotero pane?](#how-do-i-change-the-size-of-my-zotero-pane)
    - [How do I find the browser console?](#how-do-i-find-the-browser-console)
    - [How do I label different creator roles, such as Director or Producer, for films and other media?](#how-do-i-label-different-creator-roles-such-as-director-or-producer-for-films-an)
    - [How do I open and close my Zotero pane?](#how-do-i-open-and-close-my-zotero-pane)
    - [How does Zotero parse things in the name fields?](#how-does-zotero-parse-things-in-the-name-fields)
    - [How Zotero Support Works](#how-zotero-support-works)
    - [I have two Zotero libraries. How can I combine them?](#i-have-two-zotero-libraries-how-can-i-combine-them)
    - [Journal Abbreviations](#journal-abbreviations)
    - [Knowledge Base](#knowledge-base)
    - [Language Support](#language-support)
    - [Licensing](#licensing)
    - [Locate Menu](#locate-menu)
    - [My Publications](#my-publications)
    - [My snapshot is unreadable because flash ads appear on top of article text. How do I get a better snapshot?](#my-snapshot-is-unreadable-because-flash-ads-appear-on-top-of-article-text-how-do)
    - [A snapshot only captures the first of multiple pages of a New York Times article. How do I get Zotero to capture the full article?](#a-snapshot-only-captures-the-first-of-multiple-pages-of-a-new-york-times-article)
    - [Sometimes the icon for a book/article/etc doesn’t appear in the address bar right away. What’s going on?](#sometimes-the-icon-for-a-bookarticleetc-doesnt-appear-in-the-address-bar-right-a)
    - [Tips and Tricks](#tips-and-tricks)
    - [Video Tutorials](#video-tutorials)
    - [Viewing Site Certificate Information](#viewing-site-certificate-information)
    - [What connections do I need to allow through a firewall for Zotero to work properly?](#what-connections-do-i-need-to-allow-through-a-firewall-for-zotero-to-work-proper)
    - [What do I do if my Zotero database is corrupted?](#what-do-i-do-if-my-zotero-database-is-corrupted)
    - [What does “Zotero” mean?](#what-does-zotero-mean)
    - [What is the difference between Links and Snapshots?](#what-is-the-difference-between-links-and-snapshots)
    - [What version of Zotero do I have?](#what-version-of-zotero-do-i-have)
    - [Where are the link and snapshot buttons in Zotero 2.0?](#where-are-the-link-and-snapshot-buttons-in-zotero-20)
    - [Why am I getting an “unresponsive script warning”?](#why-am-i-getting-an-unresponsive-script-warning)
    - [Why can’t Zotero find a linked file?](#why-cant-zotero-find-a-linked-file)
    - [Why can’t Zotero find a stored file?](#why-cant-zotero-find-a-stored-file)
    - [Why do I no longer see a Save to Zotero option in the Firefox open/save dialog?](#why-do-i-no-longer-see-a-save-to-zotero-option-in-the-firefox-opensave-dialog)
    - [Why doesn’t Zotero have a “watch folder” feature?](#why-doesnt-zotero-have-a-watch-folder-feature)
    - [Why don’t I see a Zotero icon in the address bar when viewing a webpage?](#why-dont-i-see-a-zotero-icon-in-the-address-bar-when-viewing-a-webpage)
    - [Why don’t I see any tabs in the item pane after selecting an item?](#why-dont-i-see-any-tabs-in-the-item-pane-after-selecting-an-item)
    - [Why is there no save button in my browser toolbar?](#why-is-there-no-save-button-in-my-browser-toolbar)
    - [Why isn’t Zotero detecting updates?](#why-isnt-zotero-detecting-updates)
    - [Zotero 2.1](#zotero-21)
    - [Zotero and the Text Encoding Initiative (TEI)](#zotero-and-the-text-encoding-initiative-tei)
    - [Zotero and Wikipedia/ Wikidata](#zotero-and-wikipedia-wikidata)
    - [Zotero Beta Builds](#zotero-beta-builds)
    - [The Zotero Data Directory](#the-zotero-data-directory)
    - [Zotero Documentation](#zotero-documentation)
    - [Zotero doesn’t open when I click the Zotero status bar icon or select “Zotero” from the Firefox Tools menu.](#zotero-doesnt-open-when-i-click-the-zotero-status-bar-icon-or-select-zotero-from)
    - [Zotero Item Types and Fields](#zotero-item-types-and-fields)
    - [Zotero Keyboard Shortcuts](#zotero-keyboard-shortcuts)
    - [Zotero Privacy Policy](#zotero-privacy-policy)
    - [Zotero Security](#zotero-security)
    - [Zotero System Requirements](#zotero-system-requirements)
    - [ZOTERO TERMS OF SERVICE](#zotero-terms-of-service)
    - [Zotero Wiki Contributor License Agreement](#zotero-wiki-contributor-license-agreement)
    - [“[domain] uses an invalid security certificate. The certificate is not trusted because […]”](#domain-uses-an-invalid-security-certificate-the-certificate-is-not-trusted-becau)
    - [“The add-on could not be downloaded because of a connection failure on www.zotero.org.”](#the-add-on-could-not-be-downloaded-because-of-a-connection-failure-on-wwwzoteroo)
- [Preferences and troubleshooting](#preferences-and-troubleshooting)
    - [Advanced](#advanced)
    - [Cite](#cite)
    - [Debug Output Logging](#debug-output-logging)
    - [Export](#export)
    - [General](#general)
    - [Hidden Preferences](#hidden-preferences)
    - [How can I use multiple profiles in Zotero?](#how-can-i-use-multiple-profiles-in-zotero)
    - [How do I change Zotero settings?](#how-do-i-change-zotero-settings)
    - [I upgraded to Zotero 5.0 and now my data is missing! How do I get it back?](#i-upgraded-to-zotero-50-and-now-my-data-is-missing-how-do-i-get-it-back)
    - [Overriding Security Certificate Errors in Zotero](#overriding-security-certificate-errors-in-zotero)
    - [Profile directory location](#profile-directory-location)
    - [Reporting Zotero Problems](#reporting-zotero-problems)
    - [Reports](#reports)
    - [Sort Order](#sort-order)
    - [Why am I getting a database version error?](#why-am-i-getting-a-database-version-error)
    - [Why is Zotero telling me that some data could not be downloaded?](#why-is-zotero-telling-me-that-some-data-could-not-be-downloaded)
    - [Zotero Settings](#zotero-settings)
    - [“[domain] uses an invalid security certificate.”](#domain-uses-an-invalid-security-certificate)
- [Syncing and storage](#syncing-and-storage)
    - [Can I store my Zotero data directory in a cloud storage folder?](#can-i-store-my-zotero-data-directory-in-a-cloud-storage-folder)
    - [List of WebDAV services](#list-of-webdav-services)
    - [Sync](#sync)
    - [Syncing](#syncing)
    - [Why am I getting “The attached file could not be found” when I try to open a file in Zotero?](#why-am-i-getting-the-attached-file-could-not-be-found-when-i-try-to-open-a-file-)
    - [Why aren’t changes I make syncing between multiple devices and/or zotero.org?](#why-arent-changes-i-make-syncing-between-multiple-devices-andor-zoteroorg)
    - [Why do I keep getting file sync errors while syncing?](#why-do-i-keep-getting-file-sync-errors-while-syncing)
    - [Why does Zotero keep asking me to reconcile the same conflicts whenever I sync?](#why-does-zotero-keep-asking-me-to-reconcile-the-same-conflicts-whenever-i-sync)
    - [Why is Zotero still saying that my storage is full after I upgraded my storage plan or deleted files?](#why-is-zotero-still-saying-that-my-storage-is-full-after-i-upgraded-my-storage-p)
    - [Zotero Storage FAQ](#zotero-storage-faq)
    - [Zotero Sync Reset Options](#zotero-sync-reset-options)
    - [“Error connecting to server. Check your Internet connection.”](#error-connecting-to-server-check-your-internet-connection)
- [Zotero on iPhone and iPad](#zotero-on-iphone-and-ipad)
    - [Zotero for iOS](#zotero-for-ios)
    - [Zotero for Mobile](#zotero-for-mobile)

## Citing and bibliographies {#citing-and-bibliographies}
### Can I prevent the “Add Citation” dialog from the word processor plugins from moving behind the word processor window? {#can-i-prevent-the-add-citation-dialog-from-the-word-processor-plugins-from-movin}
#### Can I prevent the “Add Citation” dialog from the word processor plugins from moving behind the word processor window?#

Yes, by changing the hidden preference extensions.zotero.integration.keepAddCitationDialogRaised.

### Citation Styles {#citation-styles}

Zotero ships with several popular citation styles for creating citations and bibliographies, and over 10,000 additional styles can be found in the Zotero Style Repository. All these styles are written in the [Citation Style Language](https://citationstyles.org/) (CSL), a format also supported by many other programs.

#### Installing Additional Styles#
##### Zotero Style Repository#

You can install styles from the Zotero Style Repository by going to the Cite pane of the Zotero settings and clicking on the “Get additional styles…” option in the Zotero Style Manager. Search for the style you want and click the style title to install it into Zotero. You can also visit the Zotero Style Repository webpage in your browser with the Zotero Connector installed to install styles directly into Zotero.

The repository allows you to search by style name and filter by style type and academic field of study. By checking the box “Show only unique styles”, duplicate styles that share the exact same format are hidden (e.g., for the journal-specific styles “Nature”, “Nature Biotechnology”, “Nature Chemistry”, etc., only the *independent* “Nature” style is shown).

##### Alternative Installation Methods#

You can also install CSL styles (with a “.csl” extension) from local files on your computer (e.g., styles that you edit yourself or that you download from another website). In the Zotero Style Manager, click the ‘+’ button, then find the style file on your computer.

#### Managing and Editing Styles#

You can remove installed styles by clicking the ‘-’ button in the Zotero Style Manager. From this tab, you can also preview style output for the selected items in Zotero and edit installed styles.

#### Reporting Style Errors#

If a CSL style doesn’t give the expected output, first make sure that you are running the latest version of Zotero and have the most recent version of the style installed from the Zotero Style Repository. Once you have made sure that the style deviates from the style guide, instructions for authors, or published examples, report the error in the Zotero Forums. For your post, use the title “Style Error: \[Name of style\]”, and give a link to, or excerpt from, the style guide that shows that the CSL style is wrong. You can also try to edit the style yourself.

#### Requesting New Styles#

If you can’t find the style you’re looking for in the Zotero Style Repository, feel free to [request a style](https://github.com/citation-style-language/styles/wiki/Requesting-Styles). When requesting styles, please provide formatted references for the Campbell/Pedersen article and the Mares chapter listed on the linked page. Please also provide a link to a free-to-access article using the style (if available). You can also try to create the style yourself.

#### Questions#

Still have questions? Check the following FAQ entries, or, if these don’t answer your question, use the Zotero Forums:

-   Can I use Zotero in one language and create bibliographies in another?
-   Does Zotero support label/authorship trigraph styles, like \[ddb98\]?
-   DOI format in APA style
-   How can subsequent occurences of the same author replaced by a fixed term/symbol?
-   How do I get titles to show up in sentence case in bibliographies?
-   How do I prevent title casing of non-English titles in bibliographies?
-   How do I use rich text formatting, like italics and sub/superscript, in titles?
-   How do you cite a secondary source in Zotero?
-   How does Zotero parse things in the name fields?
-   I need to use Chicago style. Which of the three versions that come with Zotero should I use?
-   I’m the publisher/editor of a journal. What can I do to have Zotero support our style?
-   Journal Abbreviations
-   Missing Italics (or Italics-Only) in Word Bibliographies
-   References appear in the wrong font in Word/LibreOffice
-   Standard Citation Styles
-   What are these DOIs doing in my bibliography?
-   What is the official Harvard style?
-   Why do some citations include first names or initials?
-   Why isn’t the first letter of a subtitle in uppercase in bibliographies?

### Creating Bibliographies {#creating-bibliographies}
#### Word Processor Integration#

Using Microsoft Word, LibreOffice, or Google Docs? Zotero’s word processor integration allow you to add citations and bibliographies directly from your documents.

#### Quick Copy#

If you just want to quickly add references to a paper, email, or blog post, Zotero’s Quick Copy is the easiest way to go. Simply select items in the center column and drag them into any text field. Zotero will automatically create a formatted bibliography for you. To copy citations instead of references, hold down Shift at the start of the drag.

To configure your Quick Copy preferences, open the Zotero settings and select Export. From this tab you can do the following:

-   Select the export format
-   Set up site-specific export settings
-   Choose whether you want Zotero to include HTML markup when copying

You can also use Edit → Copy Bibliography or press Ctrl/Cmd-Shift-C to copy bibliography entries to your system clipboard and then paste them into documents. To copy citations instead of bibliography entries, use Edit → Copy Citation or Ctrl/Cmd-Shift-A.

In addition to bibliographic output, Quick Copy also supports export formats such as BibTeX and RIS.

#### Right-Click to Create Citation/Bibliography#

To create a bibliography or a citations list in Zotero, highlight one or more items, right-click them, and select “Create Bibliography from Item(s)…”. Then select a citation style for your citation/bibliography format and choose either to create a list of *Citations/Notes* or a *Bibliography*. Then choose one of the following four ways to create your citation/bibliography:

-   *Save as RTF* will save the bibliography as a rich-text file.
-   *Save as HTML* will save the bibliography as an HTML file for viewing in a web browser. This format will also embed metadata allowing other Zotero users viewing the document to capture bibliographic information.
-   *Copy to Clipboard* will copy the bibliography to your clipboard for pasting into any text field.
-   *Print* will send your bibliography straight to a printer.

#### RTF Scan#

With RTF Scan, you can write in plain text, and use Zotero to finalize your citations and bibliographies in the style you want.

### Does Zotero support label/authorship trigraph styles, like [ddb98]? {#does-zotero-support-labelauthorship-trigraph-styles-like-ddb98}
#### Does Zotero support label/authorship trigraph styles, like \[ddb98\]?#

Not yet fully. Zotero’s citation processor, [citeproc-js](https://github.com/juris-m/citeproc-js) will automatically create values for the “citation-label” variable in CSL styles consisting of four inital letters from author names See DIN-1505-2 for an example style using this format.

The formatting for automatic citation labels cannot currently be customized in Zotero. You can manually enter citation labels for individual items in your Zotero library by adding them to the “Extra” field with this format:

    citation-label: Smi01

Some discussions about this topic are available [here](https://github.com/citation-style-language/schema/issues/41) and [here](https://github.com/citation-style-language/styles/issues/678).

### How do I get titles to show up in sentence case in bibliographies? {#how-do-i-get-titles-to-show-up-in-sentence-case-in-bibliographies}
#### How do I get titles to show up in sentence case in bibliographies?#

Some citation styles, such as APA, require the use of sentence case for titles (e.g., “Oxidation and reduction of iron by acidophilic bacteria”). Others, like the Chicago Manual of Style, require title case (“Oxidation and Reduction of Iron by Acidophilic Bacteria”).

Unfortunately, it’s not possible for Zotero or any other tool to automate conversion to sentence case in a way that’s 100% reliable, since a computer can’t know for sure that something is a proper noun. (Think, for example, of the word “united”, which can be an adjective, verb, name of a company, or part of a country name.)

The solution is to **store titles in sentence case** in your Zotero library and let Zotero automate the conversion to title case if necessary. For example, when using Chicago Manual of Style, Zotero will automatically convert titles to title case, whereas when using APA it won’t try to enforce sentence casing and will simply print titles as they appear in your library.

##### Fixing Incorrect Capitalization#

If you find that a title is incorrectly showing in title case in your bibliography, make sure that it is stored in sentence case in your Zotero library. You can automate part of the conversion into sentence case by right-clicking on the unselected title in the Info pane and choosing “Sentence case”. You will then need to manually capitalize any proper nouns. You only need to do this once, and then Zotero will always use the correct case for that item, no matter what citation style you’re using.

It is also possible to prevent title casing of non-English titles.

##### Sentence Case and Subtitles#

Some styles that require sentence case, such as APA, also require that the first letter of the subtitle following a colon also be uppercase (“Age and environmental sustainability: A meta-analysis”). Zotero will automatically capitalize the subtitle for styles that require it. If the title contains multiple colons or dashes, store the main title (before the colon) in the Short Title field to indicate where the subtitle begins.

### How do I manually add a bibliographic item? {#how-do-i-manually-add-a-bibliographic-item}
#### How do I manually add a bibliographic item?#

You can manually add items by clicking the green circle with a plus icon above the middle column, selecting the appropriate item type, and filling in the bibliographic details in the right column. New items will appear in the “My Library” folder and can be organized into other folders.

### How do I prevent title casing of non-English titles in bibliographies? {#how-do-i-prevent-title-casing-of-non-english-titles-in-bibliographies}
#### How do I prevent title casing of non-English titles in bibliographies?#

Some CSL citation styles, such as the Chicago Manual of Style styles, convert titles to title case. However, title casing is specific to English. To prevent non-English titles from being title cased, specify the language of the corresponding item in your Zotero library using the “language” field.

Use two-letter language codes, e.g. “de” for German, “fr” for French, or “ja” for Japanese (four-letter codes can also be used, e.g. “de-DE” and “ja-JP”; see 0 for a list of locale codes). English items can be marked as such using “en”, “en-GB” (British English) or “en-US” (American English).

Titles should generally always be stored in sentence case; Zotero can automatically transform titles into title case, but items cannot be reliably transformed to sentence case (e.g., while treating abbreviations and proper nouns correctly). See Sentence Casing for more information.

### How do I unlink all Zotero citations in a document? {#how-do-i-unlink-all-zotero-citations-in-a-document}

Before submitting a document that you created with Zotero, you should always unlink Zotero citations so that only flat text is left in the document. This avoids distracting highlights and can prevent problems with publishing pipelines.

**Always unlink citations in a copy of the document.** Keep the original version of the document with active Zotero citations so that, if you need to make changes in response to comments (or even, say, reformat the document in another citation style), you can do so.

#### Unlinking with Zotero#

To unlink citations with Zotero, simply use the plugin’s Unlink Citations button in your copy of the document.

#### Unlinking Manually#

If for some reason you’re not able to use Zotero to unlink citations, you can unlink citations directly in Microsoft Word. In a copy of the document, select all text (Ctrl-A/Cmd-A) and press Ctrl-Shift-F9/Cmd-Shift-Fn-F9. (On a Mac, you can also press Cmd-6.)

If you have a document in Google Docs that you’re unable to unlink with Zotero, you can download it as a .docx, open it in Word, and perform the same step.

Note that this method will flatten all fields in the document, not just Zotero citations.

### How do I use rich text formatting, like italics and sub/superscript, in titles? {#how-do-i-use-rich-text-formatting-like-italics-and-subsuperscript-in-titles}
#### How do I use rich text formatting, like italics and sub/superscript, in titles?#

You can apply rich text formatting by manually adding the following HTML-like tags to fields in your Zotero library:

-   0 and 1 for *italics*
-   0 and 1 for **bold**
-   0 and 1 for _(subscript)
-   0 and 1 for ^(superscript)
-   0 and 1 for SMALLCAPS
-   0 and 1 to suppress capitalization rules (e.g., for foreign phrases within English titles)

Zotero will automatically replace these tags by the specified formatting in bibliographic output. E.g. “*Pseudomonas aureofaciens* nov. spec. and its pigments” will become “*Pseudomonas aureofaciens* nov. spec. and its pigments”.

Note that if rich text formatting has to be applied indiscriminately to entire fields (e.g. a style guide may dictate that titles should be in italics), you can modify the relevant Citation Style Language (CSL) style (see the CSL documentation and the [CSL field formatting options](http://citationstyles.org/downloads/specification.html#formatting)).

A future version of Zotero will allow visual rich-text editing without manually adding HTML tags.

### How do you cite a secondary source in Zotero? {#how-do-you-cite-a-secondary-source-in-zotero}
#### How do you cite a secondary source in Zotero?#

Sometimes, one may wish to cite a secondary source. For example,

1.  British Foreign Office, FO371, FO Minute, 30 March 1943, A3068/4/2. Vol 22507, quoted in Bryce Wood, Dismantling the Good Neighbor Policy, (Austin: University of Texas Press, 1985), 18
2.  (Johnson, 1982, as cited in Smith, 2004)

In such cases, one should cite the source that was actually consulted for the citation (in the above examples, Wood \[1985\] and Smith \[2004\]). Enter the remaining content in the prefix field of the Zotero work processor plugin. For example,

1.  British Foreign Office, FO371, FO Minute, 30 March 1943, A3068/4/2. Vol 22507, quoted in
2.  Johnson, 1982, as cited in

This way, the original sources that were not consulted don’t appear in the bibliography, which is generally regarded as good practice.

### I need to use Chicago style. Which of the three versions that come with Zotero should I use? {#i-need-to-use-chicago-style-which-of-the-three-versions-that-come-with-zotero-sh}
#### I need to use Chicago style. Which of the three versions that come with Zotero should I use?#

Zotero ships with three variants of the Chicago style (formatted examples of each style are shown below). The *author-date* format is most popular in the physical, natural, and social sciences, whereas researchers in literary, historical, and artistic fields mostly use note-based styles. For the note-based variants, the notes can either be self-explanatory and come with or without a bibliography (*full note*), or serve as a reference to a bibliographic entry (*note*).

-   Chicago Manual of Style 16th edition (author-date)
    -   **in-text citation**: (Adams 2002, 12)
    -   **bibliographic entry**: Adams, Douglas. 2002. *The Ultimate Hitchhiker’s Guide to the Galaxy*. New York: Del Rey.
-   Chicago Manual of Style 16th edition (full note)
    -   **note**: Douglas Adams, *The Ultimate Hitchhiker’s Guide to the Galaxy* (New York: Del Rey, 2002), 12.
    -   **bibliographic entry**: Adams, Douglas. *The Ultimate Hitchhiker’s Guide to the Galaxy*. New York: Del Rey, 2002.
-   Chicago Manual of Style 16th edition (note)
    -   **note**: Adams, *The Ultimate Hitchhiker’s Guide to the Galaxy*, 12.
    -   **bibliographic entry**: Adams, Douglas. *The Ultimate Hitchhiker’s Guide to the Galaxy*. New York: Del Rey, 2002.

### I’m the publisher/editor of a journal. What can I do to have Zotero support our style? {#im-the-publishereditor-of-a-journal-what-can-i-do-to-have-zotero-support-our-sty}
#### I’m the publisher/editor of a journal. What can I do to have Zotero support our style?#

Fantastic! You are the perfect person to get your style supported and approved.

Zotero uses the [Citation Style Language](http://citationstyles.org/) (CSL). By helping to create a CSL style, you also create a style that can be used by [Mendeley](http://www.mendeley.com/), [Papers3](http://www.mekentosj.com/papers/), and many other reference management programs.

If the Zotero Style Repository already contains a style for your journal, but is found to contain errors, please point out these errors on the Zotero Forums. Otherwise, either create a CSL style yourself or [request](https://github.com/citation-style-language/styles/wiki/Requesting-Styles) one at the Zotero Forums.

Once you have a CSL style that is of sufficient quality, consider linking to it on your webpage (or host it yourself) in the “Instructions for Authors” section.

### Legal Citations: Juris-M {#legal-citations-juris-m}
#### Legal Citations: Juris-M#

Zotero has only limited support for creating citation for US and UK legal cases and legislation. Support for legal citations from other jurisdictions is even more limited. For legal scholars or researchers making heavy source of legal sources, we recommend considering [Juris-M](http://juris-m.github.io/), a community-driven unofficial version of Zotero specifically designed for compliance with legal citation requirements.

#### Legal Citations: Zotero#

While we recommend Juris-M for heavy or frequent legal citations, it is possible to create proper citation for basic legal citations in Zotero, particularly if only a few such citations are needed.

##### US Legal Citation#

The two main item types supported for US legal citations are cases and legislation (“Statute” in Zotero).

##### US Legal Cases#

The most common norms for US legal citation, as prescribed in the **Bluebook** and followed by most general purpose styles like the *Chicago Manual of Style* distinguish between citations to cases that are reported in federal reporters and unreported cases. Zotero’s citation styles distinguish between these two based on the presence or absence of a reporter in the data.

*Example: Reported Cases*
Correct citation: Eaton v. IBM Corp., 925 F. Supp. 487 (S.D. Tex. 1996).
Data entry:

*Example: Unreported Cases*
Correct citation: Gilliard v. Oswald, No. 76–2109 (2d Cir. March 16, 1977).
Data entry:

Citations to cases reported from databases like WestLaw or Lexis are not currently properly supported.

##### US Legislation#

Similar to cases, legal citation practice distinguishes between legislation that is cited in the US Code (U.S.C.) and no yet codified legislation that’s cited by it’s public law number and it’s reference in the US Statutes at large (Stat.). Zotero distinguishes between the two based on the presence of a public law number.

*Example: Legislation in US Code*
Correct citation: Homeland Security Act of 2002, 6 U.S.C. § 101 (2002).
Data entry:

*Example: Legislation in Statutes at large*
Correct citation: Homeland Security Act of 2002, Pub. L. No. 107–296, 116 Stat. 2135 (2002).
Data entry:

##### UK and Australian Legal Citations#

By far the most common legal citation style in the UK is OSCOLA, the *Oxford University Standard for Citation of Legal Authorities*. The citation style for OSCOLA available in Zotero was commissioned by members of the Oxford legal faculty and is able to produce almost all item types listed in the guide, but does require very specific (and at times somewhat idiosyncratic) data entry in Zotero. To help with this task a public Zotero group contains sample references based on the examples given in the the OSCOLA guide.

Similarly, the *Australian Guide to Legal Citation* (AGLC) is the dominant legal style in Australia. The AGLC style available in Zotero has been written to accommodate almost all item types within AGLC, using similar data-entry conventions as for OSCOLA. A public group with AGLC examples is forthcoming.

### Missing Italics (or Italics-Only) in Word Bibliographies {#missing-italics-or-italics-only-in-word-bibliographies}

On rare occasions, you will find that your Zotero bibliography in Word is missing italics, even though those are required for book titles in the citation style you’ve chosen. Sometimes you will find the reverse: the entire bibliography is in italics.

First, verify that this is not an issue with the citation style. The citation should display correct formatting when hovering your mouse over it at the Zotero style repository.

If the style is displaying correctly there, you are very likely facing a [known issue with Microsoft Word](http://shaunakelly.com/word/styles/stylesoverridedirectformatting.html), which will change the formatting of the entire bibliography when more than half of the first entry in the bibliography is italicized.

#### Workaround#

Insert a “dummy” reference that will appear at the beginning of the bibliography and contains little italicized text (e.g. short book title). For bibliographies sorted alphabetically by first author, for example, insert a fake reference by Abu Aardvark. You would then remove this reference after finalizing the document and converting Zotero citations to plain text (we highly recommend saving a copy of the document before doing this).

Unfortunately, for numeric citation styles where the bibliography is sorted in order of first occurrence, inserting a fake reference would adjust the numbering for all other citations. Removing the fake reference at the end, would then make the bibliography start from 2 rather than 1. The best workaround we can suggest in this case is to duplicate the first reference in your Zotero library (so you still have an untouched original), reduce the length of the field that is italicized (e.g. delete some of the book title), and re-insert the shortened version of the reference into your document instead of the original. The reference would then have to be manually fixed in the document after finalizing it.

### Moving Documents with Zotero Citations Between Word Processors {#moving-documents-with-zotero-citations-between-word-processors}
#### Moving Documents with Zotero Citations Between Word Processors#

If you use the Zotero word processor plugin to add citations to your document and then open the document in another word processor, the Zotero citation links will be lost. To retain active Zotero citations when moving between programs, you can use the plugin to convert the document to a temporary format that can be safely transferred and then restore it in another supported word processor.

##### Word to Google Docs#

1.  In Word, use File → Save As… to create a copy of the document as a .docx with a new filename (e.g., “My Document - Transfer.docx”).
2.  Click Document Preferences in the Zotero plugin and select “Switch to a Different Word Processor…”.
3.  After the document has been converted, save the changes (File → Save).
4.  Use File → Open… from within a Google Doc to upload the file.
5.  Select Refresh from the Zotero menu in the opened Google Doc to continue using the document.

##### Google Docs to Word#

1.  In the Google Doc, use File → Make a Copy… to create a copy of the document.
2.  In the new document, select “Switch word processors…” from the Zotero menu.
3.  Select File → Download as → Microsoft Word (.docx) and save the converted file.
4.  Open the downloaded file in Word and click Refresh in the Zotero plugin to continue using the document.

##### LibreOffice to Google Docs#

1.  In LibreOffice, use File → Save As… to create a copy of the document as an .odt with a new filename (e.g., “My Document - Transfer.odt”).
2.  Click Document Preferences in the Zotero plugin and select “Switch to a Different Word Processor…”.
3.  After the document has been converted, save the changes (File → Save).
4.  Use File → Open… from within a Google Doc to upload the file.
5.  After opening the file, use File → Save as Google Docs to switch from .docx mode to Google Docs mode.
6.  Select Refresh from the Zotero menu in the opened Google Doc to continue using the document.

##### Google Docs to LibreOffice#

1.  In the Google Doc, use File → Make a Copy… to create a copy of the document.
2.  In the new document, select “Switch word processors…” from the Zotero menu.
3.  Select File → Download as → OpenDocument Format (.odt) and save the converted file.
4.  Open the downloaded file in LibreOffice and click Refresh in the Zotero plugin to continue using the document.

##### Word and LibreOffice#

You can store the citations in your document in a way that is compatible with both Word and LibreOffice by selecting “Bookmarks” in the plugin’s Document Preferences. This allows you to work on the same document with both Word and LibreOffice without going through the conversion procedure. However, storing citations as bookmarks does not work with footnote styles and may occasionally lead to citation corruption.

If you don’t intend to use both Word and LibreOffice to edit the document, you should use the conversion procedure below.

##### Word to LibreOffice#

1.  In Word, make sure citations are stored as Fields in the plugin’s Document Preferences.
2.  Use File → Save As… to create a copy of the document as a .odt with a new filename (e.g., “My Document - Transfer.odt”). Ignore any incompatibility warnings.
3.  Click Document Preferences in the Zotero plugin and select “Switch to a Different Word Processor…”.
4.  After the document has been converted, save the changes (File → Save).
5.  Open the converted file in LibreOffice and click Refresh in the Zotero plugin to continue using the document.

##### LibreOffice to Word#

1.  In LibreOffice, use File → Save As… to create a copy of the document as a .docx with a new filename (e.g., “My Document - Transfer.docx”).
2.  Click Document Preferences in the Zotero plugin and select “Switch to a Different Word Processor…”.
3.  After the document has been converted, save the changes (File → Save).
4.  Open the converted file in Word and click Refresh in the Zotero plugin to continue using the document.

### Plugins for Zotero {#plugins-for-zotero}

An active community of Zotero users has developed plugins to provide enhancements, new features, and interfaces with other programs.

We don’t currently provide a list of available plugins, but most plugins are announced and discussed in the Zotero Forums. An official plugin directory is planned.

To install a plugin in Zotero, download its 0 file to your computer. Then, in Zotero, go to “Tools → Plugins” and drag the 1 onto the Plugins window.

Be aware that plugins have full access to your Zotero and your computer. You should only install plugins from developers you trust.

Note: Word processor plugins for Word, LibreOffice, and Google Docs are bundled with Zotero. Extensions for Firefox, Chrome, and Edge are available from the download page.

### References appear in the wrong font in Word/LibreOffice {#references-appear-in-the-wrong-font-in-wordlibreoffice}
#### References appear in the wrong font in Word/LibreOffice#

The word processor plug-ins apply the “Default” (LibreOffice) or “Normal”/”Standard” (Word) style to generated citations and the paragraph in which they are inserted. The bibliography is rendered in a different style — “Bibliography” (Word) or “Reference1” (LibreOffice). As a result, inserting a citation with Zotero can remove formatting (such as indents, line-spacing, font size, etc.) from an entire paragraph and the citation can appear in an unwanted font or format. You can correct the formatting by adjusting the “Default/Normal/Standard” style (for citations) and the “Bibliography/Reference1” style (for the bibliography) in your word processor.

##### Solutions#
##### Word for Windows#

Place your cursor in a Zotero citation or bibliography. Then, choose the drop-bown button in the bottom-right corner of the [Quick Style picker](http://www.addbalance.com/word/ribbonsHome.htm#Styles) on the “Home” tab and choose “Modify Style”. Make the necessary changes to the font and paragraph formatting for the “Normal” or “Bibliography” styles and click “OK”. You can also change the formatting for individual citations using the options in the “Font” and “Paragraph” groups of the “Home” tab. See [here](http://www.lostintechnology.com/how-to/how-to-change-the-default-settings-in-microsoft-word-2007) for a guide with screenshots.

##### Word for Mac#

Place your cursor in a Zotero citation or bibliography. Then, click the [Styles Pane](https://support.office.com/en-us/article/Customize-styles-in-Word-for-Mac-1ef7d8e1-1506-4b21-9e81-adc5f698f86a) button at the end of the “Home” tab. Click the Current Style title at the top of the pane and choose “Modify Style”. Make the necessary changes to the font and paragraph formatting for the “Normal” or “Bibliography” styles and click “OK”. You can also change the formatting for individual citations using the options in the “Font” and “Paragraph” options in the “Format” menu. See [here](https://support.office.com/en-us/article/Customize-styles-in-Word-for-Mac-1ef7d8e1-1506-4b21-9e81-adc5f698f86a) for a guide with screenshots.

##### LibreOffice#

In LibreOffice, open the styles manager in “Format” → “Styles and Formatting” or by hitting F11. Right click on “Default” or “Bibliography”, select “Modify”, and make the desired changes to this style.

### RTF Scan {#rtf-scan}

Zotero’s RTF Scan feature allows users to create a fully cited document without having to use the word processor plugin. Many writers know the creator and date of a work they wish to cite off the top of their heads. Using the plugin might slow them down. Zotero can still format all the citations after the fact, however.

To use RTF Scan, create a new document in the Rich Text Format (RTF) and start writing. Whenever you wish to create a citation, write it in one of the following formats:

      {Smith, 2009}
      Smith {2009}
      {Smith et al., 2009}
      {John Smith, 2009}
      {Smith, 2009, 10-14}
      {Smith, "Title", 2009}
      {Jones, 2005; Smith, 2009}

You can also install the RTF Scan citation style into Zotero and use Quick Copy to easily copy citations in the expected format into your document without typing.

If you wish a bibliography to appear somewhere other than at the end of the document, type {Bibliography} where you wish it to appear.

Once you have finished writing, save the document (make sure it’s .RTF) and open Zotero. From the Tools menu, select “RTF Scan…”. Under Input File, select the document you’ve just created. In Output File, specify the name and location where you want the new, formatted file to be saved. Click Continue.

The Verify Cited Items screen will tell you which citations were mapped properly and which remain ambiguous. To fix an improperly mapped citation, click the icon to its right and select the correct citation in the dialog that appears. Zotero will provide suggestions for citations it is unsure of. Clicking the icon with the green arrow to the right of the suggestion will map it to that citation. Once all the citations are mapped properly, you can click Continue.

At the next screen, select the citation style you wish to use and click Continue. Zotero will then create your properly cited document. In the Output File, all the citations should be formatted properly for the selected style and, if the style calls for it, a bibliography will be included at the end, unless you specified another location.

It is important to note that, if you selected a citation format which calls for footnotes or endnotes, they will only appear properly if you open the Output File in a full-featured word processor, such as Microsoft Word or LibreOffice.

A more robust version of the RTF Scan feature that reduces the possibility of ambiguous citations and allows more flexibility for including prefixes, suffixes, and locator (page, chapter, verse, etc.) numbers in citations is provided by the RTF/ODF-Scan for Zotero plugin. This plugin requires LibreOffice for use.

### Standard Citation Styles {#standard-citation-styles}

Professional associations produce style guides in order to standardise citation methods in their field. These standards should be adhered to, unless there is a **very** good reason to use or create another style.

Below is a list of some professional association style guides. Corresponding CSL styles for most of them are at the Zotero style repository.

If you know of other official association style guides, please add them
(this is a wiki, and you are welcome to contribute).

Many style guides are in the form of books which are not freely available; your library may hold a copy, and there are also webpages like [this](http://www.library.american.edu/subject/citation.html) and [this](http://ia.juniata.edu/citation/) which give free and simple guidelines on using the styles correctly.

The use of unique identifiers, such as the DOI, is increasingly encouraged in citation styles.

#### International Standards#

-   [ISO 690](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=4888) ([Wikipedia page](http://en.wikipedia.org/wiki/ISO_690))
-   [ISO 832](http://www.iso.org/iso/catalogue_detail.htm?csnumber=5195)

#### National Standards#

-   [American National Standards Institute & National Information Standards Organization](http://www.niso.org/standards/z39-29-2005/) (free PDF available)
-   [British Standards Institution](http://www.bsi-global.com/) - standards [BS 1629:1989](http://www.bsi-global.com/en/Shop/Publication-Detail/?pid=000000000001551118), [BS 5605:1990](http://www.bsi-global.com/en/Shop/Publication-Detail/?pid=000000000001583347), [BS 6371:1983](http://www.bsi-global.com/en/Shop/Publication-Detail/?pid=000000000001547540)
-   [Commonwealth of Australia](http://publications.gov.au/styleManual.html)

#### Standards in Law#

-   [Bluebook Standard System of American Legal Citation](http://www.legalbluebook.com/)
-   [Indigo Book: A Manual of Legal Citation](https://law.resource.org/pub/us/code/blue/IndigoBook.html)
-   [Oxford University Standard for Citation of Legal Authorities (OSCOLA)](http://www.law.ox.ac.uk/publications/oscola.php) (free)

#### Standards in the Humanities#

-   [Modern Languages Association](http://www.mla.org/style) ([Wikipedia page](http://en.wikipedia.org/wiki/The_MLA_Style_Manual))
-   [Unified Style for Linguistics](http://linguistlist.org/pubs/tocs/JournalUnifiedStyleSheet2007.pdf) (PDF, free)
-   [American Political Science Association](http://www.ipsonet.org/data/files/APSAStyleManual2006.pdf) (PDF, free)
-   [Modern Humanities Research Association](http://www.mhra.org.uk/Publications/Books/StyleGuide/index.html)

#### Standards in Sciences#
##### Chemical, Physical and Life Sciences#

-   [Council of Science Editors](http://www.councilscienceeditors.org/publications/style.cfm)
-   [American Chemical Society](http://pubs.acs.org/page/books/styleguide/index.html) ([Wikipedia page](http://en.wikipedia.org/wiki/ACS_style))
-   [American Institute of Physics](http://www.aip.org/pubservs/style/4thed/AIP_Style_4thed.pdf) (PDF, free)

##### Engineering, Computer Science and Technology#

-   [IEEE](http://www.ieee.org/portal/cms_docs_iportals/iportals/publications/authors/transjnl/stylemanual.pdf) (PDF, free)

##### Social and Behavioural Sciences#

-   [American Psychological Association](http://www.apastyle.org/) ([Wikipedia page](http://en.wikipedia.org/wiki/APA_style), [free APA tutorial](http://www.apastyle.org/learn/tutorials/basics-tutorial.aspx))
-   [American Sociological Association](http://www.asanet.org/galleries/default-file/asaguidelinesnew.pdf) (PDF, free)
-   [American Anthropological Association](http://www.aaanet.org/publications/guidelines.cfm) (free)

##### Medical and Biomedical#

-   [American Medical Association](http://www.amamanualofstyle.com//oso/public/index.html)
-   [International Committee of Medical Journal Editors](http://www.icmje.org/urm_main.html/) (free) based on:
-   [National Library of Medicine](http://www.ncbi.nlm.nih.gov/bookshelf/br.fcgi?book=citmed) (free)

#### Commercial Style guides#

In addition a number of commercial style guides are published, for example.

-   [Chicago Manual of Style](http://www.chicagomanualofstyle.org)
-   [Turabian](http://www.press.uchicago.edu/books/turabian/index.html)

***
More examples of styles in use in particular fields can be seen at the [Wikipedia page on Style guides](http://en.wikipedia.org/wiki/Style_guide#Examples)

### Troubleshooting Errors in Word Processor Documents {#troubleshooting-errors-in-word-processor-documents}

Follow the steps below for Microsoft Word, Google Docs, or LibreOffice.

#### Microsoft Word#

**June 10, 2026: The June 9 Windows security updates from Microsoft broke Word integration on Windows. We've fixed the issue in Zotero 9.0.5, available via Help → "Check for Updates…".**

If you get an error trying to use Zotero in a **new, empty document**, see Word Processor Plugin Troubleshooting.

If you can insert citations into a new, empty Word document but get an error using Zotero in an **existing document**, follow these steps:

1.  Restart both Zotero and Word.
2.  Make sure you’re using the latest versions of Zotero and Word.
3.  While troubleshooting, disable the Track Changes feature in Word, as it can have complicated effects when working with Zotero. If Track Changes is enabled when you insert or modify a Zotero citation, it may mark many or all of the Zotero citations in your document as changed or cause field codes to be displayed. On rare occasions, Track Changes may cause Zotero to think a citation is corrupted. If you had Track Changes enabled previously, try accepting all changes to see if that resolves the issue.
4.  If using Windows, in Word Options → Advanced, make sure “Typing replaces selected text” is checked.
5.  If you use any clipboard enhancing software on macOS such as TextExpander, LaunchBar, Pure Paste, etc., temporarily disable it.
6.  Check for citations in image captions. Zotero won’t let you insert them, but if you copied a citation to a caption, that’s most likely the source of the problem. Delete it.
7.  Try copying and pasting the document content into a new document to see if the problem goes away. You may need to click the “Set Document Preferences” button before your old citations will be recognized.
8.  Make a copy of your document — by duplicating the file itself, not by copying and pasting the content — to use for debugging.
9.  If using OneDrive on Windows, save the copy of the document to your local hard drive, or try renaming the file to remove any spaces in the filename. OneDrive is known to interfere with the plugin for some people.
10. Open the copied file and check if you get the error after switching to a different bibliography style.
11. If the document has a bibliography, delete it completely and check if you still get the error.
12. While debugging, if you are using Fields mode in the Word plugin, it may help to display field codes rather than formatted text. To do this, press Alt/Option-F9 (or Alt/Option-Fn-F9) in Word.
13. **Isolate the problematic citations.** In the copy of your document, delete half of the contents at a time and see if the error still occurs. If not, use Undo to restore the deleted section and then try deleting the other half. Repeat the halving process on the section that fails, or pick one at random if both do. Continue this until you find the smallest possible section, ideally with a single citation, that must be present for the problem to occur. Remove the isolated citations from the original document and the problem should go away (unless there are multiple broken citations, in which case you’ll need to repeat the process). Unless the error still occurs if you completely clear the contents of the document, **this final step will by definition identify the problem.**

If you encounter a broken document, please create a new thread in the Zotero Forums so we can attempt to fix the issue. Include a Report ID from Zotero, your operating system and Word versions, and the steps you’ve taken to try to fix the error. You should also send the document excerpt from the final step and a link to your forum thread to support@zotero.org so that we can try to reproduce the problem.

#### Google Docs#

If you get an error trying to use Zotero in a **new, empty document**, see Google Docs Troubleshooting.

If you can insert citations into a new, empty Google Doc but get an error using Zotero in an **existing document**, follow these steps:

1.  Restart both Zotero and your browser.
2.  Make sure you’re using the latest versions of Zotero and the Zotero Connector.
3.  Disable any other browser extensions and reload Google Docs.
4.  Try using File → “Make a copy” to see if the problem goes away in a new document. You may need to click the “Set Document Preferences” button before your old citations will be recognized.
5.  In the copy of the document, check if you get the error after switching to a different bibliography style.
6.  If the document has a bibliography, delete it and check if you still get the error.
7.  **Isolate the problematic citations.** In the copy of your document, delete half of the contents at a time and see if the error still occurs. If not, use Undo to restore the deleted section and then try deleting the other half. Repeat the halving process on the section that fails, or pick one at random if both do. Continue this until you find the smallest possible section, ideally with a single citation, that must be present for the problem to occur. Remove the isolated citations from the original document and the problem should go away (unless there are multiple broken citations, in which case you’ll need to repeat the process). Unless the error still occurs if you completely clear the contents of the document, **this final step will by definition identify the problem.**

If you encounter a broken document, please create a new thread in the Zotero Forums so we can attempt to fix the issue. Include a Debug ID from the Zotero for reproducing the problem and the steps you’ve taken to try to fix the error. You should also make a sharing link for the document excerpt from the final step and email it to support@zotero.org with a link to your forum thread so that we can try to reproduce the problem.

#### LibreOffice#

If you get an error trying to use Zotero in a **new, empty document**, see Word Processor Plugin Troubleshooting.

If you can insert citations into a new, empty LibreOffice document but get an error using Zotero in an **existing document**, follow these steps:

1.  Restart both Zotero and LibreOffice.
2.  Make sure you’re using the latest versions of Zotero and LibreOffice.
3.  While troubleshooting, disable the Track Changes feature in LibreOffice, as it can have complicated effects when working with Zotero. If Track Changes is enabled when you insert or modify a Zotero citation, it may mark many or all of the Zotero citations in your document as changed or cause field codes to be displayed. On rare occasions, Track Changes may cause Zotero to think a citation is corrupted. If you had Track Changes enabled previously, try accepting all changes to see if that resolves the issue.
4.  Check for citations in image captions. Zotero won’t let you insert them, but if you copied a citation to a caption that’s most likely the source of the problem. Delete it.
5.  Make sure your LibreOffice text style uses the same text style in the “Next Style” property.
6.  Try copying and pasting the document content into a new document to see if the problem goes away. You may need to click the “Set Document Preferences” button before your old citations will be recognized.
7.  Make a copy of your document — by duplicating the file itself, not by copying and pasting the content — to use for debugging.
8.  Open the copied file and check if you get the error after switching to a different bibliography style.
9.  If the document has a bibliography, delete it and check if you still get the error.
10. While debugging, if you are using Reference Marks mode in the LibreOffice plugin, it may help to display field codes rather than formatted text by pressing Ctrl-F9.
11. **Isolate the problematic citations.** In the copy of the document, delete half of the contents at a time and see if the error still occurs. If not, use Undo to restore the deleted section and then try deleting the other half. Repeat the halving process on the section that fails, or pick one at random if both do. Continue this until you find the smallest possible section, ideally with a single citation, that must be present for the problem to occur. Remove the isolated citations from the original document and the problem should go away (unless there are multiple broken citations, in which case you’ll need to repeat the process). Unless the error still occurs if you completely clear the contents of the document, **this final step will by definition identify the problem.**

If you encounter a broken document, please create a new thread in the Zotero Forums so we can attempt to fix the issue. Include a Report ID from Zotero, your operating system and LibreOffice versions, and the steps you’ve taken to try to fix the error. You should also send the document excerpt from the final step and a link to your forum thread to support@zotero.org so that we can try to reproduce the problem.

### Using the Zotero LibreOffice Plugin {#using-the-zotero-libreoffice-plugin}

These are instructions for using the Zotero LibreOffice Plugin. For plugins for Word or Google Docs, see Word Processor Plugins.

#### Zotero Plugin Toolbar#

Installing the Zotero LibreOffice plugin adds a Zotero toolbar to LibreOffice.

The Zotero toolbar contains these icons:

[TABLE]

#### Citing#

You can begin citing with Zotero by clicking the “Add/Edit Citation” () button. Pressing the button brings up the citation dialog.

The citation dialog is used to select items from your Zotero library, and create a citation.

Start typing part of a title, the last names of one or more authors, and/or a year in the dialog box. Matching items will instantly appear below the dialog box.

Matching items will be shown for each library in your Zotero database (My Library and any groups you are part of). Items you have already cited in the document will be shown at the top of the list under “Cited”.

Select an item by clicking on it or by pressing Enter/Return when it is highlighted. The item will appear in the dialog box in a shaded bubble. Press Enter/Return again to insert the citation and close the Add Citation box.

In the Add Citation dialog box, you can click on the bubble for a cited item, then click “Open in My Library (or the Group Library’s name)” to view the item in Zotero. Items that are orphaned (not connected to any items in your Zotero database) will not have an “Open in My Library” button. Orphaned items can exist if they were inserted by a collaborator from their My Library or a group you don’t have access to or if you they were deleted from your Zotero library.

#### Bibliography#

Clicking the “Add/Edit Bibliography” () button inserts a bibliography at the cursor location.

You can edit which items appear in the bibliography by clicking the “Add/Edit Bibliography” button again, which will open the bibliography editor. See below. Manual edits made to the bibliography in LibreOffice will be overwritten the next time Zotero refreshes the document.

#### Document Preferences#

The “Document Preferences” window lets you set the following document-specific preferences:

1.  The citation style.
2.  The language to use to format citations and bibliography.
3.  For note-based styles (e.g., “Chicago Manual of Style (Note)”), whether citations are inserted in footnotes or endnotes.
    -   Note that Word, not Zotero, controls the style and format of footnotes and endnotes.
4.  Whether to store citations as **ReferenceMarks** or **Bookmarks**.
    -   Unless you need to collaborate with colleagues using Word, you should always choose ReferenceMarks.
5.  For styles that abbreviate journal titles (e.g., “Nature”), whether to use the **MEDLINE** abbreviations list to abbreviate titles.
    -   If this option is selected (the default), the contents of the “Journal Abbr” field in Zotero will be ignored.

#### Customizing Cites#

Citations can be customized in various ways.

If a citation is simply incorrect or missing data, start by making sure that the item metadata in Zotero is correct and complete, and then click Refresh in the plugin to update your document with any changes.

Other customizations can be made via the citation dialog. Click an existing citation in your document and click Add/Edit Citation to open the citation dialog, and then click the citation bubble to open the cite options window, where you can make the following changes.

##### Page and Other Locators#

In some cases you want to cite a certain part of an item, e.g. a certain page, page range or volume. This additional cite-specific information (e.g. “pp. 4-7” in the cite “Doe et al. 2001, p. 4-7”) is called the “locator”.

The cite options windows has a drop-down list of the different locator types (“Page” is the default), and a text box in which you can enter the locator value (e.g. “4-7”). To cite a locator other than the ones listed (e.g., “Table), use the Suffix field.

You can also add page numbers from the keyboard as you insert citations. Search for an item, press Enter once to add to the citing dialog, and then, before pressing Enter again to insert it into the document, simply type “p.34” or similar, and the page number will be added to the citation.

##### Prefix and Suffix#

The “Prefix” and “Suffix” text boxes allow you to specify text to respectively precede and follow the automatically generated cite. For example, instead of “Tribe 1999”, you might want “cf. Tribe 1999, see also…”.

Any text in the prefix and suffix fields can be formatted with the HTML tags 0 (for italics), 1 (bold), 2 (subscript), and 3 (superscript). For example, typing “4cf5. the classic example” will be displayed as “*cf*. the classic example”.

Prefixes and suffixes can be applied to each item in a citation to create complex citations. For example: “(see Smith 1776 for the classic example; Marx 1867 presents and alternate view)”. Modifying citations by entering text into the Prefix and Suffix fields is always preferable to directly typing in the citation fields in LibreOffice. Manual modifications will prevent Zotero from automatically updating the citation.

##### Omitting Authors: Using Authors in the Text#

With author-date styles, authors are often moved into the text and omitted from the following parentheses-enclosed citation, e.g.: “…according to Smith (1776) the division of labor is crucial…”. To omit the authors from the cite, check the “Omit Author” box (this will result in a cite like “(1776)” instead of “(Smith, 1776)”) and write the author’s name (“Smith”) as part of the regular text in your document.

##### Citations with Multiple Cited Items#

To create a citation containing multiple cites (e.g., “\[2,4-6\]” for numeric styles or “(Smith 1776, Schumpeter 1962)” for author-date styles), add them one after the other in the Add Citation box. After selecting the first item, don’t press Enter/Return, but type the author, title, or year of the next item.

Some citation styles require that items within one in-text citations are ordered either alphabetically (e.g., “(Doe 2000, Grey 1994, Smith 2008)”) or chronologically (“(Grey 1994, Doe 2000, Smith 2008)”). Zotero will follow these sort rules automatically.

-   To disable automatic sorting of the cites in the citation, drag the citations to rearrange them in the Add Citation box. You can also click the “Z” icon on the left side of the Add Citation box and uncheck the “Keep Sources Sorted” option. *This option only appears for citation styles that specify a sort order for citations.* To restore automatic sorting, re-check the “Keep Sources Sorted” option.

##### Switching to the “Classic View”#

You can switch to the “Classic View” citation dialog by clicking the “Z” icon on the left side of the Add Citation box, and selecting “Classic View”. To permanently switch to the classic view check the “Use classic Add Citation view” checkbox in the Cite pane of Zotero preferences.

#### Editing the Bibliography#

After you’ve inserted the bibliography using the “Add/Edit Bibliography” () button, click the button again to open the Edit Bibliography window.

In this window, you can add uncited sources to your bibliography (e.g., items included in a review but not cited in the paper) or remove items that are cited in text but which should not be included in the bibliography (e.g., personal communications).

While it is also possible to edit the text or formatting of bibliography references in this window, doing so is discouraged. References edited here will not be automatically updated by Zotero if you change the data in your library. Editing references here is also somewhat unreliable; several users have reported that modifications made here sometimes do not persist when Zotero references, among other issues.

If you need to edit items in your bibliography, it is best to do this as a final step before submitting the document. First, save a backup copy of the document. Then, click the “Unlink Citations” button () to disconnect your document from Zotero and convert all citations and the bibliography to regular text. Finally, make your adjustments to the bibliography text.

This process can be used for a variety of minor modifications to the bibliography, including:

-   Adding asterisks \* before references included in a review or meta-analysis
-   Setting the names of particular authors in bold, italics, or all caps
-   Adding annotations or comments about an item
-   Adding headings for bibliography subsections (e.g., primary versus secondary sources)

**Note:** General corrections to style formatting should be made in the CSL Citation Style, not here. Corrections to item data should be made in your Zotero library, not here.

#### Keyboard Commands#

The Zotero LibreOffice plugin can be used with just the keyboard for improved accessibility and faster use.

-   Keyboard shortcuts can be set up for all the buttons in the Zotero tab.
-   In the Citation dialog
    -   Use the up and down arrow keys to move between search results. Press Enter to select an item.
    -   Type “p.45-48” or “:45-48” after a citation to cite a specific page or page range.
    -   Type “ibid” to automatically select the last cited work. This works with all citation styles, regardless of whether “ibid” is actually used in citations. If you use Zotero in a language other than English, use the corresponding abbreviation instead of ibid., e.g. “ebd.” in German.
    -   Press Ctrl/Cmd-↓ (down arrow key) to open the cite options dialog for the citation selected with the cursor. Use Tab and Shift-Tab to move between the different elements, use the up and down arrow keys to change the locator type in the locator drop-down list, and the space bar to toggle the “Suppress Author” checkbox.

#### Troubleshooting#

If you run into problems while trying to use the Zotero LibreOffice plugin, make sure to check out the word processor plugin troubleshooting page.

### Using the Zotero Word Plugin {#using-the-zotero-word-plugin}

These are instructions for using the Zotero Word Plugin. For plugins for LibreOffice or Google Docs, see Word Processor Plugins.

#### Zotero Plugin Tab#

Installing the Zotero Word plugin adds a Zotero tab to Microsoft Word.

The Zotero tab contains these icons:

[TABLE]

#### Citing#

You can begin citing with Zotero by clicking the “Add/Edit Citation” () button. Pressing the button brings up the citation dialog.

The citation dialog is used to select items from your Zotero library, and create a citation.

Start typing part of a title, the last names of one or more authors, and/or a year in the dialog box. Matching items will instantly appear below the dialog box.

Matching items will be shown for each library in your Zotero database (My Library and any groups you are part of). Items you have already cited in the document will be shown at the top of the list under “Cited”.

Select an item by clicking on it or by pressing Enter/Return when it is highlighted. The item will appear in the dialog box in a shaded bubble. Press Enter/Return again to insert the citation and close the Add Citation box.

In the Add Citation dialog box, you can click on the bubble for a cited item, then click “Open in My Library (or the Group Library’s name)” to view the item in Zotero. Items that are orphaned (not connected to any items in your Zotero database) will not have an “Open in My Library” button. Orphaned items can exist if they were inserted by a collaborator from their My Library or a group you don’t have access to or if you they were deleted from your Zotero library.

#### Bibliography#

Clicking the “Add/Edit Bibliography” () button inserts a bibliography at the cursor location.

As you use the plugin, Zotero will automatically update the bibliography based on the citations in the document.

In the rare case that you want to add items to the bibliography that you haven’t cited in the document, you can click “Add/Edit Bibliography” button again, which will open the bibliography editor. Manual edits made to the bibliography in Word will be overwritten the next time Zotero refreshes the document.

#### Document Preferences#

The “Document Preferences” window lets you set the following document-specific preferences:

1.  The citation style.
2.  The language to use to format citations and bibliography.
3.  For note-based styles (e.g., “Chicago Manual of Style (Note)”), whether citations are inserted in footnotes or endnotes.
    -   Note that Word, not Zotero, controls the style and format of footnotes and endnotes.
4.  Whether to store citations as **Fields** or **Bookmarks**.
    -   Unless you need to collaborate with colleagues using LibreOffice, you should always choose Fields.
5.  For styles that abbreviate journal titles (e.g., “Nature”), whether to use the **MEDLINE** abbreviations list to abbreviate titles.
    -   If this option is selected (the default), the contents of the “Journal Abbr” field in Zotero will be ignored.

#### Customizing Citations#

Citations can be customized in various ways.

If a citation is simply incorrect or missing data, start by making sure that the item metadata in Zotero is correct and complete, and then click Refresh in the plugin to update your document with any changes.

Other customizations can be made via the citation dialog. Click an existing citation in your document and click Add/Edit Citation to open the citation dialog, and then click the citation bubble to open the cite options window, where you can make the following changes.

##### Page and Other Locators#

In some cases you want to cite a certain part of an item, e.g. a certain page, page range or volume. This additional cite-specific information (e.g. “pp. 4-7” in the cite “Doe et al. 2001, p. 4-7”) is called the “locator”.

The cite options windows has a drop-down list of the different locator types (“Page” is the default), and a text box in which you can enter the locator value (e.g. “4-7”). To cite a locator other than the ones listed (e.g., “Table), use the Suffix field.

You can also add page numbers from the keyboard as you insert citations. Search for an item and, before or after selecting it, but before pressing Enter to insert the citation into the document, type “p.34”, “p34”, or even just “34”. The page number will be added to the citation.

##### Prefix and Suffix#

The “Prefix” and “Suffix” text boxes allow you to specify text to respectively precede and follow the automatically generated cite. For example, instead of “Tribe 1999”, you might want “cf. Tribe 1999, see also…”.

Any text in the prefix and suffix fields can be formatted with the HTML tags 0 (for italics), 1 (bold), 2 (subscript), and 3 (superscript). For example, typing “4cf5. the classic example” will be displayed as “*cf*. the classic example”.

Prefixes and suffixes can be applied to each item in a citation to create complex citations. For example: “(see Smith 1776 for the classic example; Marx 1867 presents and alternate view)”. Modifying citations by entering text into the Prefix and Suffix fields is always preferable to directly typing in the citation fields in Word. Manual modifications will prevent Zotero from automatically updating the citation.

##### Narrative Citations with Omit Author (“According to Smith (1776)”)#

With author-date styles, authors are often moved into the text and omitted from the following parentheses-enclosed citation, e.g.: “According to Smith (1776), the division of labor is crucial…”. To omit the authors from the cite, check the “Omit Author” box, which will result in a cite like “(1776)” instead of “(Smith, 1776)”, and write the author’s name (“Smith”) as part of the regular text in your document.

##### Citations with Multiple Cited Items#

To create a citation containing multiple cites (e.g., “\[2,4-6\]” for numeric styles or “(Smith 1776, Schumpeter 1962)” for author-date styles), add them one after the other in the Add Citation box. After selecting the first item, don’t press Enter/Return, but type the author, title, or year of the next item.

Some citation styles require that items within one in-text citations are ordered either alphabetically (e.g., “(Doe 2000, Grey 1994, Smith 2008)”) or chronologically (“(Grey 1994, Doe 2000, Smith 2008)”). Zotero will follow these sort rules automatically.

-   To disable automatic sorting of the cites in the citation, drag the citations to rearrange them in the Add Citation box. You can also click the “Z” icon on the left side of the Add Citation box and uncheck the “Keep Sources Sorted” option. *This option only appears for citation styles that specify a sort order for citations.* To restore automatic sorting, re-check the “Keep Sources Sorted” option.

##### Switching to the “Classic View”#

You can switch to the “Classic View” citation dialog by clicking the “Z” icon on the left side of the Citation box, and selecting “Classic View”. To permanently switch to the classic view check the “Use classic Add Citation view” checkbox in the Cite pane of Zotero preferences.

##### Other Changes#

If your citation still isn’t displaying the way you want, you can edit the citation directly in your document, but note that doing so will prevent Zotero from being able to automatically update the citation to reflect other changes in the document (e.g., for ‘ibid.’ or given name disambiguation). After you make a manual edit, Zotero will ask you to confirm that you want to keep the edit and prevent the citation from being updated automatically going forward. It may be preferable to instead make notes in the text of changes you want to make, wait until you’re ready to submit the document, and make the changes in a copy of the document after using Unlink Citations.

If you believe there’s an error in a citation style, post to the Zotero Forums so that we can investigate and, if necessary, correct the style. If a style is updated, your document will automatically update to reflect any changes the next time you refresh the document.

#### Editing the Bibliography#

As you add and remove citations in the document, Zotero will automatically update the bibliography to reflect your changes. Generally, that’s all you have to do.

In rare cases, however, you may want to add uncited sources to your bibliography (e.g., items included in a review but not cited in the paper) or remove items that are cited in text but which should not be included in the bibliography (e.g., personal communications). To do this, click the “Add/Edit Bibliography” () button again to open the Edit Bibliography window:

You can then use the arrows to add or remove items.

While it is also possible to edit the text or formatting of bibliography references in this window, doing so is discouraged. References edited here will not be automatically updated by Zotero if you change the data in your library. Editing references here is also somewhat unreliable; several users have reported that modifications made here sometimes do not persist when Zotero references, among other issues.

If you need to edit items in your bibliography, it is best to do this as a final step before submitting the document. First, save a backup copy of the document. Then, click the “Unlink Citations” button () to disconnect your document from Zotero and convert all citations and the bibliography to regular text. Finally, make your adjustments to the bibliography text.

This process can be used for a variety of minor modifications to the bibliography, including:

-   Adding asterisks \* before references included in a review or meta-analysis
-   Setting the names of particular authors in bold, italics, or all caps
-   Adding annotations or comments about an item
-   Adding headings for bibliography subsections (e.g., primary versus secondary sources)

**Note:** General corrections to style formatting should be made in the CSL Citation Style, not here. Corrections to item data should be made in your Zotero library, not here.

#### Keyboard Commands#

The Zotero Word plugin can be used with just the keyboard for improved accessibility and faster use.

-   Keyboard shortcuts can be set up for all the buttons in the Zotero tab.
-   In the Citation dialog
    -   Use the up and down arrow keys to move between search results. Press Enter to select an item.
    -   Type “p.45-48” or “:45-48” after a citation to cite a specific page or page range.
    -   Type “ibid” to automatically select the last cited work. This works with all citation styles, regardless of whether “ibid” is actually used in citations. If you use Zotero in a language other than English, use the corresponding abbreviation instead of ibid., e.g. “ebd.” in German.
    -   Select a citation with the mouse or left/right arrow keys and press down arrow or space bar to open the cite options dialog. Use Tab and Shift-Tab to move between the different elements, use the up and down arrow keys to change the locator type in the locator drop-down list, and the space bar to toggle the “Omit Author” checkbox.

#### Troubleshooting#

If you run into problems while trying to use the Zotero Word plugin, make sure to check out the word processor plugin troubleshooting page.

### Using Zotero with Google Docs {#using-zotero-with-google-docs}

Zotero’s powerful Google Docs support helps you easily add citations and bibliographies to the documents you create in Google Docs.

You can quickly search for items in your Zotero library, add page numbers and other details, and insert citations. When you’re done, a single click inserts a formatted bibliography based on the citations in your document. Zotero supports complex style requirements such as *Ibid.* and name disambiguation, and it keeps your citations and bibliography updated as you make changes to items in your library. If you need to switch citation styles, you can easily reformat your entire document in any of the over 10,000 citation styles that Zotero supports.

Google Docs support is provided by the Zotero Connector for Chrome, Firefox, Edge, and Safari and requires the Zotero program to function.

Using another word processor? Zotero also integrates with Word and LibreOffice.

#### Citation Interface#

The Zotero Connector adds a Zotero menu to the Google Docs interface:

It also adds a toolbar button for one-click citing:

In the Zotero menu, you’ll find the following options:

[TABLE]

#### Authorization#

Interacting with the Zotero functionality for the first time in a document will prompt you to authorization the plugin to access your Google account. Be sure to:

1\. Select the Google account you used to create the document or that has been given editing access by the document’s creator. This is unrelated to any Zotero account you may have, which isn’t required to use Zotero or Google Docs integration.

2\. Grant Zotero the permission to “See, edit, create and delete all your Google Docs documents”. Zotero requires this permission to be able to insert and modify citations into your document. The plugin doesn’t do anything else with your document content and doesn’t access documents other than the ones on which it’s triggered. The integration works entirely locally on your computer, so even when you trigger the plugin on a given document, nothing is sent to Zotero servers.

Once you’ve authorized the plugin to access your document, you can begin inserting citations from your Zotero libraries.

#### Citing#

You can begin citing by clicking the (“Add/Edit Zotero Citation”) button in the Google Docs toolbar or by selecting “Add/Edit Citation” from the Zotero menu, both of which will bring up the citation dialog.

The citation dialog is used to select items from your Zotero library and create a citation.

Start typing part of a title, the last names of one or more authors, and/or a year in the dialog box. Matching items will instantly appear below the dialog box.

Matching items will be shown for each library in your Zotero database (My Library and any groups you are part of). Items you have already cited in the document will be shown at the top of the list under “Cited”.

Select an item by clicking on it or by pressing Enter when it is highlighted. The item will appear in the dialog box in a shaded bubble. Press Enter again to insert the citation and close the Add Citation box.

In the Add Citation box, you can click on the bubble for a cited item and then click “Open in My Library” (or another library name) to view the item in Zotero. Items that are orphaned (i.e., not connected to any items in your Zotero database) will not have an “Open in My Library” button. Orphaned items can exist if they were inserted by a collaborator from their My Library or a group you don’t have access to or if they were deleted from your Zotero library.

#### Bibliography#

Clicking the “Add/Edit Bibliography” menu option inserts a bibliography at the cursor location.

You can edit which items appear in the bibliography by clicking the “Add/Edit Bibliography” button again, which will open the bibliography editor. See Editing the Bibliography below for more info. Manual edits made to the bibliography in the document will be overwritten the next time Zotero refreshes the document.

#### Collaboration#

Google Docs is designed to let you collaborate on documents, and Zotero’s integration is no different. You and your coauthors can all insert and edit citations in a shared document, and you don’t even need to be in a Zotero group. If you’re planning a large collaborative project, though, we recommend using a group library, which not only makes it easy to collect and manage materials but will also allow all collaborators to change cited item metadata (authors, title, date of publication, etc.). If someone cites an item from their personal library, only they will be able to update the metadata for that item.

We recommend that anyone making changes to the document have the Zotero Connector installed. (The Zotero app itself is necessary only if inserting or editing citations.) If someone cuts and pastes an active citation without the Zotero Connector, the citation will be unlinked from Zotero and disappear from the bibliography, and the next person refreshing the document with the Zotero Connector will receive a warning about unlinked citations. While people without the Connector can theoretically edit non-citation parts of the document, we don’t recommend it due to the risk of accidental citation unlinking.

When working collaboratively on a document, you and your coauthors should avoid inserting or editing citations at the same time. The Zotero Connector has mechanisms in place to prevent document and citation corruption from concurrent citation editing, but due to technical limitations they do not provide perfect safety.

#### Document Preferences#

The “Document Preferences” window lets you set the following document-specific preferences:

1.  The citation style.
2.  The language to use to format citations and bibliography.
3.  For styles that abbreviate journal titles (e.g., “Nature”), whether to use the **MEDLINE** abbreviations list to abbreviate titles.
    -   If this option is selected (the default), the contents of the “Journal Abbr” field in Zotero will be ignored.
4.  Whether Zotero should automatically update citations for disambiguations, ibid and numbering, or whether updating should be delayed until a manual refresh. Note that if you enable this mode, Zotero will insert your citations with a gray background to indicate that the citation text is not final. The citation will be finalized and the gray background removed once you click “Refresh” in the Zotero menu.

#### Saving for Publication#

When you’re ready to submit your document, use File → “Make a copy…” and, in the new document, use Zotero → “Unlink Citations” to convert the citations and bibliography to plain text. You can then download that second document (e.g., as a PDF), while keeping active citations in the original document in case you need to make further changes. Zotero will prompt you to create a copy if you try to download your original document.

#### Customizing Cites#

Citations can be customized in various ways.

If a citation is simply incorrect or missing data, start by making sure that the item metadata in Zotero is correct and complete, and then click Refresh in the plugin to update your document with any changes.

Other customizations can be made via the citation dialog. Click an existing citation in your document and click Add/Edit Citation to open the citation dialog, and then click the citation bubble to open the cite options window, where you can make the following changes.

##### Page and Other Locators#

In some cases you want to cite a certain part of an item, e.g. a certain page, page range or volume. This additional cite-specific information (e.g. “pp. 4-7” in the cite “Doe et al. 2001, p. 4-7”) is called the “locator”.

The cite options windows has a drop-down list of the different locator types (“Page” is the default), and a text box in which you can enter the locator value (e.g. “4-7”). To cite a locator other than the ones listed (e.g., “Table), use the Suffix field.

You can also add page numbers from the keyboard as you insert citations. Search for an item, press Enter once to add to the citing dialog, and then, before pressing Enter again to insert it into the document, simply type “p.34” or similar, and the page number will be added to the citation.

##### Prefix and Suffix#

The “Prefix” and “Suffix” text boxes allow you to specify text to respectively precede and follow the automatically generated cite. For example, instead of “Tribe 1999”, you might want “cf. Tribe 1999, see also…”.

Any text in the prefix and suffix fields can be formatted with the HTML tags 0 (for italics), 1 (bold), 2 (subscript), and 3 (superscript). For example, typing “4cf5. the classic example” will be displayed as “*cf*. the classic example”.

Prefixes and suffixes can be applied to each item in a citation to create complex citations. For example: “(see Smith 1776 for the classic example; Marx 1867 presents and alternate view)”. Modifying citations by entering text into the Prefix and Suffix fields is always preferable to directly typing in the citation fields in the document. Manual modifications will prevent Zotero from automatically updating the citation.

##### Omitting Authors: Using Authors in the Text#

With author-date styles, authors are often moved into the text and omitted from the following parentheses-enclosed citation, e.g.: “…according to Smith (1776) the division of labor is crucial…”. To omit the authors from the cite, check the “Omit Author” box (this will result in a cite like “(1776)” instead of “(Smith, 1776)”) and write the author’s name (“Smith”) as part of the regular text in your document.

##### Citations with Multiple Cited Items#

To create a citation containing multiple cites (e.g., “\[2,4-6\]” for numeric styles or “(Smith 1776, Schumpeter 1962)” for author-date styles), add them one after the other in the Add Citation box. After selecting the first item, don’t press Enter/Return, but type the author, title, or year of the next item.

Some citation styles require that items within one in-text citations are ordered either alphabetically (e.g., “(Doe 2000, Grey 1994, Smith 2008)”) or chronologically (“(Grey 1994, Doe 2000, Smith 2008)”). Zotero will follow these sort rules automatically.

-   To disable automatic sorting of the cites in the citation, drag the citations to rearrange them in the Add Citation box. You can also click the “Z” icon on the left side of the Add Citation box and uncheck the “Keep Sources Sorted” option. *This option only appears for citation styles that specify a sort order for citations.* To restore automatic sorting, re-check the “Keep Sources Sorted” option.

##### Switching to the “Classic View”#

You can switch to the “Classic View” citation dialog by clicking the “Z” icon on the left side of the Citation box, and selecting “Classic View”. To permanently switch to the classic view check the “Use classic Add Citation view” checkbox in the Cite pane of Zotero preferences.

#### Editing the Bibliography#

After you’ve inserted the bibliography using the “Add/Edit Bibliography” option, select it again to open the Edit Bibliography window.

In this window, you can add uncited sources to your bibliography (e.g., items included in a review but not cited in the paper) or remove items that are cited in text but which should not be included in the bibliography (e.g., personal communications).

While it is also possible to edit the text or formatting of bibliography references in this window, doing so is discouraged. References edited here will not be automatically updated by Zotero if you change the data in your library.

If you need to edit items in your bibliography, it is best to do this as a final step before submitting the document. First, make a copy of the document. Then, in the copy, use the “Unlink Citations” menu option to disconnect your document from Zotero and convert all citations and the bibliography to regular text. Finally, make your adjustments to the bibliography text.

This process can be used for a variety of minor modifications to the bibliography, including:

-   Adding asterisks before references included in a review or meta-analysis
-   Setting the names of particular authors in bold, italics, or all caps
-   Adding annotations or comments about an item
-   Adding headings for bibliography subsections (e.g., primary versus secondary sources)

**Note:** General corrections to style formatting should be made in the CSL citation style, not in this window. Corrections to item data should be made in your Zotero library.

#### Keyboard Shortcuts#

You can use keyboard shortcuts for improved accessibility and faster citing.

-   Press Ctrl-Command-C (Mac) or Ctrl-Alt-C (Windows/Linux) to insert a citation. You can configure this from the Advanced pane of the Zotero Connector preferences.
-   In the citation dialog
    -   Use the up and down arrow keys to move between search results. Press Enter to select an item.
    -   Type “p.45-48” or “:45-48” after a citation to cite a specific page or page range.
    -   Type “ibid” to automatically select the last cited work. This works with all citation styles, regardless of whether “ibid” is actually used in citations. If you use Zotero in a language other than English, use the corresponding abbreviation instead of ibid., e.g. “ebd.” in German.
    -   Press Ctrl/Cmd-↓ (down arrow key) to open the cite options dialog for the citation under the cursor. Use Tab and Shift-Tab to move between the different elements, use the up and down arrow keys to change the locator type in the locator drop-down list, and the space bar to toggle the “Suppress Author” checkbox.

#### Limitations#

While we’ve tried to create the same experience available in Word and LibreOffice, there are some limitations to be aware of when working in Google Docs:

-   As noted above, anyone making changes to the document should have the Zotero Connector installed. (The Zotero app itself is necessary only if inserting or updating citations.) Citations that are cut and pasted without the Connector installed will be unlinked.
-   Dragging citations within the document will cause the citations to become unlinked. Cutting and pasting is fine as long as the Zotero Connector is installed.
-   If someone views the document without having the Zotero Connector installed, or if you download the document instead of first making a copy and unlinking citations, active citations in the document will show up as links leading to URLs such as 0.
-   Citation inserts and edits slow down significantly as the number of citations increases. With 100+ citations, a single citation update can take up to 10 seconds, so for longer documents you’ll want to disable automatic citation updates in the Zotero document preferences.
-   Google Docs provides limited facilities for text formatting. Styles that use small caps fonts will not use a true small caps formatting style in Google Docs and will instead fall back to the “Alegreya Sans SC” font. Citations that have been inserted with automatic citation updates disabled will be inserted with a gray background instead of dashed underlining like in Word and LibreOffice.

#### Troubleshooting#
##### Menu doesn’t appear#

If nothing appears when you click the Zotero menu, or you see a thin gray line, try restarting Zotero and your browser.

If that doesn’t help, disable all other browser extensions, reload Google Docs, and try again. In particular, the Google Docs Offline extension has been reported as interfering with Zotero’s Google Docs integration.

In some browsers, you may need to give the Zotero Connector permission to run. While Google Docs support only requires access to 0 and 1, if you’re going to be using Zotero, you’ll want to use the Zotero Connector to save to Zotero, and for that to work it needs to be able to run on all sites. (See why this is safe.) In Chrome or Edge, right-click on the Save to Zotero button in your browser toolbar, select “This Can Read and Change Site Data”, and choose “On All Sites”. In Safari, go to the Websites tab of the Safari settings, click on Zotero Connector in the left column, and make sure any sites that show up and “For other websites” at the bottom are all set to “Allow”.

If it’s still not working, try in a new browser profile (e.g., a new Chrome profile) or in a different browser.

##### Citation dialog doesn’t appear after clicking Add/Edit Citation#

If you can open the Zotero menu but the citation dialog doesn’t appear after you click Add/Edit Citation, make sure that a dialog isn’t appearing behind your other browser or Zotero windows.

If you’re sure that’s not the problem, generate a Debug ID for reloading Google Docs and clicking Add/Edit Citation and post it to the Zotero Forums along with a description of the problem.

##### Unlinked Citations#

See Why isn’t Zotero detecting my existing Google Docs citations?

##### “The Google account you selected does not have permission to edit this document”#

You likely selected the wrong Google account. See Authentication.

##### Other problems#

If you encounter other problems citing in Google Docs, let us know in the Zotero Forums. Provide a Debug ID from the Zotero Connector for reloading Google Docs and trying to perform the relevant action.

You should always troubleshoot in a new, empty document or with a copy of the original document, using File → “Make a copy…”. If something isn’t working in a particular document, the document version history may allow you to revert to an earlier version. Some of the Debugging Broken Documents steps may also be useful in Google Docs.

### What happened to the “classic” citation dialog? {#what-happened-to-the-classic-citation-dialog}

The “classic” citation dialog was the original citation dialog in Zotero, introduced in 2006. It was replaced as the default dialog in 2011 by the “red bar” citation dialog, which allowed for faster searching and citing via the keyboard. The “classic” dialog remained an option for people who preferred to choose citations by browsing their collections rather than searching.

In 2026, Zotero 8 introduced a new unified citation dialog, replacing the “red bar” dialog, the “classic” dialog, and the Add Note dialog (the “yellow bar”). The new dialog includes a Library mode that directly replaces the classic dialog, bringing all the efficiency-enhancing features of the previous default dialog to people who want to browse by collection.

We know that switching to a new interface after many years can be a little jarring, but we believe that most people who give the new dialog a chance and spend a few days with it will quickly find that it’s worth the change.

#### Frequently Asked Questions#
##### Can I switch back to the classic dialog?#

No. The classic dialog has been removed and won’t be returning. Many features have been added to Zotero’s word processor integration over the years, and the classic dialog wasn’t able to take advantage of any of them. Maintaining a separate, parallel dialog also diverts development time from improvements to the parts of Zotero that everyone uses. The new dialog was designed to combine the library browser that classic-dialog users want with all the modern features.

If you’re having trouble adjusting to the new dialog, let us know in the Zotero Forums. We’re actively improving both the documentation and the dialog itself in response to feedback.

##### How do I browse my collections like before?#

Select “Library” mode in the bottom right of the window. Library mode shows the same collections and items lists you had in the classic dialog.

List mode, by comparison, works similarly to the previous (“red-bar”) default dialog. It’s the fastest mode when you generally know exactly what you’re searching for and don’t need to limit your search to a given collection.

By default, Zotero will open in the last mode you used. If you prefer to always open in one or the other, set Citation Dialog Mode in the Cite tab of the settings.

##### How do I add a page number to a citation?#

After choosing an item, you can type a page number directly into the main input field, rather than clicking around the window:

You can also add other locator types by typing the short or full name (e.g., “chap4” or “chapter 4”).

##### Why does it take more steps to do everything in the new dialog?#

It doesn’t! The new dialog requires fewer steps to perform most common actions, greatly improves keyboard usage, and provides new features to speed up citing.

Bear in mind that the classic dialog hasn’t been the default citation dialog in Zotero for 15 years. Citing hasn’t been slower for all the people using the default dialog that entire time — it’s been faster.

See Comparing Common Actions for a detailed comparison between the classic and new dialogs.

##### How do I cite multiple items at once?#

Use Ctrl/Cmd or Shift to select multiple items in the items list, and then press Enter/Return to add all of them to your citation.

##### How do I reorder items?#

Instead of clicking up/down arrows to move items into the right order, you can simply drag them to where you want them. To move with the keyboard, use left/right-arrow to select an item and Shift-left/right to move it.

##### What are these random items appearing at the top of the dialog?#

The new dialog lets you quickly add citations for selected items and open documents.

When you open the dialog, the selected item in your library, or the selected document tab in the reader, will be shown in a section at the top of the window. To select the first item, you can simply press Enter/Return on your keyboard or click it with your mouse. You can also choose among other open documents.

This means that, if you just want to insert a citation for the PDF you’re reading in Zotero, you can just click Add/Edit Citation and press Enter/Return twice to insert it into your document.

##### What does “Cited Items” mean?#

If you’ve already cited an item in your library, you can type its author or title, and it will show up in a Cited Items section at the top of the dialog. This also helps you avoid accidentally creating duplicate references for items that are duplicated in your library.

##### What happened to the citation editor?#

In the new dialog, there’s no text field to make manual edits to citations. It’s been possible to edit citations directly in the document since the introduction of the red-bar dialog in 2011, which is why the red bar never included such a text field.

More importantly, though, such manual edits should be avoided in almost all cases, since they prevent Zotero from updating the citation as you edit metadata, add other citations, or change citation styles. Instead, customize the citation via the citation dialog, which will allow Zotero to continue to update the citation as necessary. If you’re finding yourself regularly wanting to edit citations, post to the Zotero Forums with examples, and we may be able to suggest a better approach that avoids breaking citation updates.

#### Comparing Common Actions#

Below, we’ve detailed the steps necessary to perform some common actions in the classic and new dialogs.

##### Adding a single item with a page number#

0

**Classic dialog:**

1.  Type “smith”
2.  Click item
3.  Click locator field
4.  Type “123”
5.  Press Enter/Return or click Done

**New dialog:**

1.  Type “smith”
2.  Click “+” next to item or double-click row
3.  Type “123”
4.  Press Enter/Return or click Accept (checkmark)

**Classic dialog, keyboard-only:**

1.  Type “smith”
2.  Tab twice to items list and ↓ to item
3.  Tab four (!) times to locator field
4.  Type “123”
5.  Press Enter/Return

**New dialog, keyboard-only:**

1.  Type “smith”
2.  ↓ to item + Enter/Return, or just Enter/Return if only result
3.  Type “123”
4.  Press Enter/Return

**Winner:** New dialog, for both speed and much easier keyboard use

##### Adding a single item with a unique search, with a chapter number#

0

**Classic dialog:**

1.  Type “gatsby”
2.  Click item
3.  Click the locator dropdown and select Chapter
4.  Click in the locator field
5.  Type “3”
6.  Press Enter/Return or click Done

**Classic dialog, keyboard-only:**

1.  Type “gatsby”
2.  Tab five (!) times to locator menu
3.  Navigate dropdown with keyboard and choose Chapter
4.  Tab once to locator field
5.  Type “3”
6.  Press Enter/Return

**New dialog:**

1.  Type “gatsby chap3”
2.  Press Enter/Return twice

**Winner:** New dialog

##### Adding a multi-item citation with page numbers and a prefix#

0

**Classic dialog:**

1.  Type “smith”
2.  Click item
3.  Click in locator field
4.  Type “123”
5.  Click Multiple Sources…
6.  Double-click search box to select all text
7.  Type “jones”
8.  Click item
9.  Click right-arrow button
10. Click in Prefix field
11. Type “see also”
12. Click in locator field
13. Type “234”
14. Click OK

**Classic dialog, keyboard only:**

*Not possible — right-arrow button isn’t keyboard accessible*

**New dialog:**

1.  Type “smith”
2.  Click “+” next to item or double-click row
3.  Type “123”
4.  Type “jones”
5.  Click “+” next to item or double-click row
6.  Click citation bubble
7.  Type “234”
8.  Click in Prefix field or press Tab
9.  Type “see also”
10. Press Enter/Return twice

**New dialog, keyboard only:**

1.  Type “smith”
2.  Press ↓ and Enter/Return, or press Enter/Return if only result
3.  Type “123”
4.  Type “jones”
5.  Press ↓ and Enter/Return, or press Enter/Return if only result
6.  Press ← + ↓ to open citation bubble
7.  Type “234”
8.  Tab once to Prefix field
9.  Type “see also”
10. Press Enter/Return twice

**Winner:** New dialog, for both speed and keyboard accessibility

##### Adding a citation for the current PDF tab in Zotero#

**Classic dialog:**

1.  Remember the details of the PDF you have open
2.  Type search term
3.  Click item
4.  Click Done

**New dialog:**

1.  Press Enter/Return twice

**Winner:** New dialog!

### What is the official Harvard style? {#what-is-the-official-harvard-style}

“Harvard style” is a popular name for the author-date type of parenthetical referencing.

This way of referencing does not originate from Harvard University.\[1\] There are many variants of “Harvard style” in use, and there is no “official” Harvard Style.

If you are looking for a style called “Harvard”, you can check the many “Harvard” styles available in the Zotero Style Repository to see if there’s one for your institution.

If there isn’t a style for your institution, you can use the [CSL Search by Example tool](https://editor.citationstyles.org/searchByExample) to find a style that matches your institution’s style guide.

If you’re not able to find a good match, you can [request a new style](https://github.com/citation-style-language/styles/wiki/Requesting-Styles).

\[1\] Anecdotal evidence suggests the name originates with an English visitor of the Harvard University Museum of Comparative Zoology. See Chernin, Eli (1988). [“The ‘Harvard system’: a mystery dispelled”](http://www.pubmedcentral.nih.gov/picrender.fcgi?artid=1834803&blobtype=pdf), *British Medical Journal*. October 22, 1988, pp. 1062–1063.

### Where is the Zotero toolbar in Word for Mac 2008? {#where-is-the-zotero-toolbar-in-word-for-mac-2008}
#### Where is the Zotero toolbar in Word for Mac 2008?#

The Zotero word processor plugin for Word for Mac 2008 doesn’t offer a toolbar, instead adding a “Zotero” entry to the AppleScript menu (the manuscript icon to the right of the Help menu):

Most menu commands are also available through shortcut keys.

Word for Mac 2008 lacks support for Visual Basic for Applications (VBA), making it impossible to create a toolbar. VBA support was restored in Word for Mac 2011, and the Zotero plugin for Word 2011 and 2016 includes a toolbar (Word 2011) or Zotero tab (Word 2016).

### Why are my citations underlined with a dashed line? {#why-are-my-citations-underlined-with-a-dashed-line}

As you insert citations into your document, Zotero often needs to update other citations in the document and the bibliography to reflect the new citation. For example, if you insert a citation by an author with the same last name as another citation in the document, the style guidelines may require both names to include first initials to help readers tell the authors apart.

When the citations in your document are underlined with a dashed line, it means you have these automatic citation updates disabled. The underline warns you that the citation may not be fully up to date or correctly formatted. In large documents, citation updates can take a while, and Zotero likely prompted you to disable automatic updating to speed up your writing.

To update all of the citations in your document with automatic citations disabled, click the plugin’s Refresh button. You only need to do this once when you’re ready to submit the document (though you may want to do it occasionally as a test to make sure you won’t run into any unexpected issues before a deadline).

To manually enable or disable automatic citation updates, open the plugin’s Document Preferences window and check or uncheck “Automatically update citations”. To avoid accidentally submitting a paper with unformatted citations, we recommend leaving automatic updates enabled unless you find that inserts are taking too long for a given document.

### Why are Zotero citations or bibliographies always highlighted in gray or another color? {#why-are-zotero-citations-or-bibliographies-always-highlighted-in-gray-or-another}

By default, Zotero stores the reference data for citations and the bibliography in [Fields](https://support.office.com/en-us/article/Insert-fields-in-Word-c429bbb0-8669-48a7-bd24-bab6ba6b06bb) (Word) or [Reference Marks](https://help.libreoffice.org/Writer/About_Fields) (LibreOffice), which stores items’ reference data hidden behind the formatted text.

Word and LibreOffice will highlight Fields/Reference Marks on your screen to indicate that the text is automatically generated. This can help you avoid accidentally manually typing in the fields (to edit the text shown in a Zotero citation, see Customizing Cites). These highlights are only shown on screen and won’t appear if you print or save the document as a PDF.

You can change the settings for highlighting Fields/Reference Marks in your word processor:

-   **[Word for Windows](https://helpdeskgeek.com/office-tips/show-field-shading-in-word-and-convert-the-fields-to-plain-text/):** In Word Options, open “Advanced”, then set “Field shading” to “Never”, “Always”, or “When selected”.
-   **Word for Mac:** Open Word -> Preferences -> View and set “Field shading” to “Never”, “Always”, or “When selected”.
-   **[LibreOffice](https://help.libreoffice.org/Writer/About_Fields):** Open Tools -> Options -> LibreOffice -> Application Colors and check/uncheck the “Field shadings” box. You can also control the color used for field shadings.

### Why do I see code beginning with ADDIN ZOTERO_ITEM CSL_CITATION in my document instead of formatted citations? {#why-do-i-see-code-beginning-with-addin-zoteroitem-cslcitation-in-my-document-ins}
#### Why do I see code beginning with ADDIN ZOTERO\_ITEM CSL\_CITATION in my document instead of formatted citations?#

By default, Zotero stores the reference data for citations and the bibliography in [Fields](https://support.office.com/en-us/article/Insert-fields-in-Word-c429bbb0-8669-48a7-bd24-bab6ba6b06bb) (Word) or [Reference Marks](https://help.libreoffice.org/Writer/About_Fields) (LibreOffice), which stores items’ reference data hidden behind the formatted text.

If Word or LibreOffice is showing the field codes, rather than the formatted text, you can hide the field codes by pressing Alt-F9 (Option-Fn-F9 on a Mac) in Word or Ctrl-F9 in LibreOffice. In Word, you can also select one or more citations (or do a Select All), right-click, and choose Toggle Field Codes.

If you see field codes showing repeatedly, check your word processor settings for displaying fields. In [Word for Windows](http://mikefrobbins.com/2010/05/10/how-to-toggle-field-codes-off-or-on-in-word/), open Word Options, then choose “Advanced” and uncheck the “Show field codes instead of their values” box. In Word for Mac, open Word -> Preferences -> View and uncheck the “Field codes instead of values” box in the “Show in Document” area. In [LibreOffice](https://help.libreoffice.org/Writer/Field_Names), open Tools -> Options -> LibreOffice Writer -> View, and uncheck the “Field codes” box in the “Display” area.

When using Track Changes, Word or LibreOffice will sometimes show you the changed and original versions of changed field codes. If this occurs, click “Accept” twice to quickly accept the changed citations and hide the field codes again.

### Why do some citations include first names or initials? {#why-do-some-citations-include-first-names-or-initials}
#### Why do some citations include first names or initials?#

Sometimes you’ll see Zotero produce citations like “(J. Doe, 2004)” even though your citation style normally only shows the author’s last name (“Doe, 2004”). In these cases, Zotero is disambiguating different authors according to the rules of your selected citation style. This is generally what you want it to do, and if you think otherwise, you should carefully review your style’s requirements. (APA style, for example, requires this sort of disambiguation when you cite different authors with the same last name.)

Disambiguation can also occur if a certain author is inconsistently named in your Zotero library. For example, Zotero treats the names

-   Jeff Smith
-   J. Smith
-   J. R. Smith

as distinct individuals and will disambiguate them according to the style rules. You can fix this by going through your library and changing all names that refer to the same person to the exact same form, which will allow Zotero to disambiguate authors correctly in this style and other styles you use in the future.

If you don’t want to update the names in your library, you can also simply use a style that doesn’t disambiguate names. A step-by-step guide to disable given name disambiguation in a CSL citation style can be found here.

Zotero and CSL support [sophisticated disambiguation rules](http://citationstyles.org/downloads/specification.html#disambiguation). If you think citations are being disambiguated incorrectly, please post to the Zotero Forums and provide documentation in the form of a style guide or a published article in the publication in question. *Note that these changes do not address the issue of inconsistently named authors.*

##### Other Causes#
##### Deleted Items#

If you’ve made sure that the author’s name is formatted consistently in Zotero, one or more of your items may be pointing to an item that you’ve deleted from Zotero. When this happens, Zotero uses metadata embedded in the document instead. To check whether an item is still linked to your Zotero library, click on the citation, click Add/Edit Citation, click the blue bubble in the citation dialog (red bar, not classic), and look for an “Open in \[Library\]” button at the bottom of the popup. If the button doesn’t appear, the citation is no longer linked to Zotero, and you will need to delete the citation and reinsert it, being sure to select from the appropriate library rather than the “Cited” section in the citation dialog search results.

If you’re having trouble finding the relevant citations, it may help to display Word field codes and search for the title or author in the field code. Each citation’s field code will include a 0 field with a URL like 1. If two URLs don’t match, those are pointing at different items. Follow the steps above to identify or reinsert a citation from your library and then make sure that all similar citations in the document match the URL from that citation.

##### Duplicate Items#

If you have more than one copy of an item and cite both in separate places, Zotero will treat the authors as separate and disambiguate them in the text. Within the same library, you can fix this by merging the items in Zotero, either via the Duplicate Items view or by selecting the items in the main items list, right-clicking, and choosing “Merge Items…”, and then clicking the plugin’s Refresh button. If you’ve cited items from different libraries, use Add/Edit Citation as described in the previous section to identify the associated item for each citation and make sure each one is pointing to the same item in just one of the libraries.

##### Disabled Automatic Citation Updates#

If you’re seeing citations with dashed underlines, you’ve disabled Automatic Citation Updates. You can either press Refresh to update citations manually or re-enable Automatic Citation Updates in the plugin’s document preferences.

### Why is a citation not updated in my document after editing the item in Zotero? {#why-is-a-citation-not-updated-in-my-document-after-editing-the-item-in-zotero}

When you make changes to item data (title, author, date, etc.) in your Zotero library, those changes will be reflected in citations to those items in your word processor the next time you use the Zotero plugin’s Refresh button.

If your citations are flat text and aren’t being detected at all, see Existing Citations Not Detected.

If you still have active citations (e.g., highlighted in gray when you click on them in Word or LibreOffice, or showing the “Edit in Zotero” popup in Google Docs) but changes you make in Zotero aren’t being reflected in the document, either you edited the citation text manually and told Zotero not to make further updates or the citation in your document is no longer linked to the item in your Zotero library.

Click the citation and click Add/Edit Citation. If you edited the citation, Zotero will notify you that the citation was modified and give you the option to reset it, after which the citation will update automatically. You should generally avoid manual edits in the document and customize the citation to add page numbers, prefixes, etc.

If the citation hasn’t been modified, the citation dialog will open with the citation shown. To check whether the citation is still linked to your library, click the blue bubble and look for the “Open in My Library \[or the group library name\]” button in the popup:

If the “Open in…” button doesn’t appear, the citation isn’t linked to any of your Zotero libraries and Zotero is using the item metadata embedded in the document to generate the citation and bibliography entry. You will need to delete the citation from your document and reinsert it, being sure to select from the library section of the citation dialog search results rather than from the Cited section.

Citations can become orphaned for a number of reasons:

1.  You have duplicate items in your library, cite one of the duplicates, and then delete it rather than merging the items
2.  You delete an item from your library and then reimport it
3.  You cite an item with Mendeley and then edit the document with Zotero (though Zotero can relink Mendeley Desktop citations after you import your Mendeley library)

### Why is Zotero slow to insert citations or update the bibliography? {#why-is-zotero-slow-to-insert-citations-or-update-the-bibliography}
#### Why is Zotero slow to insert citations or update the bibliography?#

When you insert a citation into a document using Zotero’s word processor plugin, Zotero needs to scan the entire document for citations to ensure correct formatting. Citation style requirements such as *ibid* or name disambiguation mean that the format of a given citation may depend on the citations that precede it, and bibliographies depend on the presence of, and in some cases the order of, all citations in the document, including any that may have been deleted or moved around since a citation was last inserted.

In longer documents, scanning the entire document can take multiple seconds or even minutes, and these updates can become disruptive to the writing process. Word for Mac 2008 and Google Docs are especially prone to slow down due to technical limitations of those programs.

To speed up your writing, you can disable automatic updates and defer citation updating until a manual refresh is triggered. With automatic updates disabled, citation inserts will remain instantaneous regardless of the size of the document.

To disable automatic updates, click the Document Preferences button in the word processor plugin and uncheck “Automatically Update Citations”:

To illustrate how citation inserting works with updates disabled, let’s look at an example. Say we’ve added a citation for a paper by Jessica Smith using APA style:

If we then insert a paper by James Smith, Zotero will create a citation in the default format required by the style without taking into account other citations in the document:

Zotero adds a dashed underline below newly added citations to remind us that they haven’t been updated (though keep in mind that existing citations later in the document might also now be incorrect).

It also replaces the bibliography with a warning:

We can quickly insert citations this way without waiting for each update. When we’re ready to submit our document, we click the Zotero plugin’s Refresh button:

Zotero scanned the document and updated the citations and bibliography to conform to the style rules, which in this case require disambiguation for lead authors with the same last name.

To avoid accidentally submitting a paper with unformatted citations, we recommend leaving automatic updates enabled unless you find that inserts are taking too long for a given document.

Alternative methods to speed up citing if you want to keep automatic updates enabled are to split long documents into chapters or to use a less-demanding citation style, such as Annual Reviews (author-date) during writing to increase the speed of citation inserts.

### Why isn’t the first letter of a subtitle in uppercase in bibliographies? {#why-isnt-the-first-letter-of-a-subtitle-in-uppercase-in-bibliographies}
#### Why isn’t the first letter of a subtitle in uppercase in bibliographies?#

Some styles that require sentence case, such as APA, also require that the first letter of the subtitle following a colon also be uppercase (“Age and environmental sustainability: A meta-analysis”).

Zotero will automatically uppercase the subtitle for APA style and other styles based on it. If you encounter a style that should have uppercase subtitles but doesn’t currently, please report in on the Zotero forums.

### Why isn’t Zotero detecting my existing citations? {#why-isnt-zotero-detecting-my-existing-citations}

If you’ve previously used one of Zotero’s word processor plugins to insert citations into a document and later find that 1) the plugin says “You must insert a citation before performing this operation”, 2) the bibliography doesn’t contain all citations in the document, and/or 3) references in a numeric citation style start from 1 instead of from an appropriate higher number, the existing citations in the document may no longer be active fields.

To check whether fields in a document are active, click them and looking for a gray highlight (Word/LibreOffice) or “Edit with Zotero” popup (Google Docs). If you then click Add/Edit Citation, the Zotero citation dialog should appear with the citation shown. If the citation dialog is empty, the citation is no longer active. In Word and LibreOffice, you can also try toggling field codes.

Your citations may have become inactive for a few reasons:

1.  You used the “Unlink Citations” button. This button will disconnect your document from Zotero and convert all citations and bibliographies to plain text.
2.  You (or someone else) saved the document in an unsupported file type:
    -   When using Word, you need to save your document as .docx. If you save as .odt, active citations will be lost.
    -   When using LibreOffice, if you are storing citations as ReferenceMarks (the default), you must save your document as .odt. If storing citations as Bookmarks, you must save your document as .docx.
3.  You (or someone else) opened and saved your document using an unsupported word processor or without following proper steps to move active citations between supported word processors:
    -   **Google Docs:** If you want to open a Word or LibreOffice document in Google Docs, or vice versa, you must follow an extra step to transfer the document. Directly opening a Word or LibreOffice document in Google Docs, or vice versa, will break existing citations.
    -   **Other online word processors**: Most online word processors do not support Fields/ReferenceMarks/Bookmarks. Opening an existing document in these tools will break connections with Zotero. Microsoft’s Word Online does support Fields and so can be used safely with Word documents containing Zotero fields, though a Zotero plugin is not currently available for Word Online.
    -   **Pages:** Apple Pages does not support Fields/ReferenceMarks/Bookmarks. Opening a document in Pages will break connections with Zotero.
    -   **Word:** If you open an .odt file (created by LibreOffice) in Word, Zotero references stored as ReferenceMarks (the default) will be broken. To share a document between Word and LibreOffice users, change the “Store Citations as:” option in the Zotero Document Preferences to Bookmarks. (Bookmarks can cause errors if accidentally modified, so they should only be used if compatibility between Word and LibreOffice is necessary. You can also choose to transfer the document instead.)
    -   **LibreOffice:** If you open a .docx or .doc file (created by Word) in LibreOffice, Zotero references stored as Fields (the default) will be broken. To share a document between Word and LibreOffice users, change the “Store Citations as:” option in the Zotero Document Preferences to Bookmarks. (Bookmarks can cause errors if accidentally modified, so they should only be used if compatibility between Word and LibreOffice is necessary. You can also choose to transfer the document instead.)
4.  If you’re using Google Docs, see Why isn’t Zotero detecting my existing Google Docs citations? for other possible reasons.

If you find that your citations have been flattened, your only options are to restore the document from a backup, re-insert the citations using the plugin, or manually edit the document’s citations going forward and generate a final bibliography from a collection in Zotero without using the plugin. If reinserting citations, it may help to adjust Word field settings to always highlight fields in gray rather than only doing so when they are selected.

If your bibliography is flattened but your citations are still active, you can simply insert a new bibliography by clicking Add/Edit Bibliography.

### Why isn’t Zotero detecting my existing Google Docs citations? {#why-isnt-zotero-detecting-my-existing-google-docs-citations}

Google Docs integration uses a different mechanism for storing citations than is used in the Word and LibreOffice plugins, and it can be easier for citations to accidentally become unlinked. The most common reason for unlinking is collaborators editing the document without having the Zotero Connector installed. To ensure Zotero citations stay linked, everyone editing the Google Doc should install the Zotero Connector in their browser. (The Zotero app itself is only required for inserting and editing citations.) Citations can also become unlinked if you drag them with the mouse within the Google Doc rather than cutting and pasting them.

To restore unlinked citations, you have two options:

1.  Use the Google Docs version history and restore to an earlier version of the document before the citation was unlinked. The version history can be found under Google Docs → File → Version History → “See version history”. Note that the warning about unlinked citations will show up on the next operation *after* the operation that caused the problem, so restoring to a previous version that doesn’t immediately show citations as unlinked isn’t sufficient — you need to press Refresh in the Zotero menu and confirm that the warning doesn’t reappear. If it does, revert to an earlier version.
2.  Use the Zotero plugin to reinsert the missing unlinked citations.

If all your collaborators have the Zotero Connector installed but you still see citations becoming unlinked, try to identify the actions that lead to the links breaking via the Google Docs version history and post details to the Zotero Forums. Again, be sure to press Refresh after restoring to confirm that the citations aren’t already unlinked in that version.

For information regarding other word processor plugins, see Existing Citations Not Detected.

### Word Processor Plugin Shortcuts {#word-processor-plugin-shortcuts}

In most word processors it is possible to assign keyboard shortcuts to the various functions of the Zotero word processor plugin toolbar (i.e. add citation, edit citation etc.). How to do this depends on your word processor.

#### Word for Windows#

1.  Click the Office Symbol at the top right of the program.
2.  Click “Word Options.”
3.  Select the “Customize Tab.”
4.  Click the “Customize” button at the bottom of the window next to “Keyboard Shortcuts.”
5.  Select the “Macros” category in the box on the left.
6.  Locate the Zotero items in the box on the right and select one to assign a keyboard shortcut.
7.  If it already has a shortcut it will show in the “Current Keys” box.
8.  To assign one, place the cursor in the “Press new shortcut key” field.
9.  Press the keyboard combination you want to assign, e.g. when assigning a combination to add a new Zotero citation, press ctrl-alt-A
10. If the “Currently assigned to” field says “\[unassigned\]”, you can use this shortcut without any conflicts with other commands.
11. Repeat the process for the other Zotero commands.

See [this blog post](http://diyivorytower.wordpress.com/2012/03/06/create-zotero-hotkeys-in-word/) for longer, illustrated instructions.

#### Word for Mac 2016 and newer#

1.  Open Tools -> Customize keyboard…
2.  Select the “Macros” category in the box on the left.
3.  Locate the Zotero items in the box on the right and select one to assign a keyboard shortcut.
4.  Press the keyboard combination you want to assign, e.g. when assigning a combination to add a new Zotero citation, press ctrl-option-A
5.  If it already has a shortcut it will show in the “Current Keys” box.
6.  If the “Currently assigned to” field says “\[unassigned\]”, you can use this shortcut without any conflicts with other commands.
7.  Repeat the process for the other Zotero commands.

See here for more details.

#### LibreOffice#

1.  Open Tools -> Customize…
2.  Click on the key to you want to use, from the list of assignable keys in the top portion of the dialog.
3.  Find the appropriate Zotero macro in the bottom part of the dialog. The macros are in LibreOffice Macros -> user -> Zotero -> Zotero. Click on the desired action in the list (probably ZoteroAddCitation)
4.  Click the “Modify” button.
5.  Click “OK” and enjoy!

#### Google Docs#

See Google Docs Keyboard Shortcuts.

### Word Processor Plugins {#word-processor-plugins}

Of the different ways to automatically generate bibliographies (as well as in-text citations and footnotes), the easy-to-use word processor plugins are the most powerful. These plugins, available for Microsoft Word, LibreOffice, and Google Docs, create dynamic bibliographies: insert a new in-text citation in your manuscript, and the bibliography will be automatically updated to include the cited item. Correct the title of an item in your Zotero library and with a click of a button the change will be incorporated in your documents.

To get started with these plugins, see the following pages:

-   Using the Zotero Word Plugin
-   Using the Zotero LibreOffice Plugin
-   Using Zotero with Google Docs
-   Troubleshooting

Third-party plugins are also available for integrating Zotero with other word processors and writing systems.

#### Plugin Installation#

The word processor plugins are bundled with Zotero and should be installed automatically for each supported word processor on your computer when you first start Zotero.

You can reinstall the plugins later from the Cite -> Word Processor Plugins pane of the Zotero preferences. If you’re having trouble, see Manually Installing the Zotero Word Processor Plugin or Word Processor Plugin Troubleshooting.

### Zotero and Word Compatibility on Apple Silicon Macs {#zotero-and-word-compatibility-on-apple-silicon-macs}
#### Zotero and Word Compatibility on Apple Silicon Macs#

To use Zotero with Microsoft Word on an Apple Silicon Mac, please make sure you’re running macOS 11.4 or later. A bug in macOS 11.3 and earlier can cause Zotero to freeze when using the plugin.

(If you previously set Word to open under Rosetta to work around this bug, you should select Word in Finder, go to File → Get Info, and uncheck “Open using Rosetta”, and then restart Word.)

### Zotero does not have permission to control Word {#zotero-does-not-have-permission-to-control-word}

Recent macOS versions include new security measures when a program attempts to control another program on the system. Zotero’s Word for Mac plugin requires the permission for Zotero to be able to control Word. When you first interact with the Zotero Word plugin, you’ll get a prompt asking for this permission:

If you press “Don’t allow”, Zotero won’t be able to provide the Word plugin functionality and every subsequent attempt to use the plugin will trigger the “Missing Permissions” prompt, until you follow the steps in the prompt:

#### macOS 13 Ventura and later#

1.  Open System Settings
2.  Select “Privacy & Security” in the left column
3.  Select “Automation”
4.  Find “Zotero” and click the arrow to expand it
5.  Make sure “Microsoft Word” is enabled under “Zotero”
6.  Restart Word

#### macOS 12 Monterey and earlier#

1.  Open System Preferences
2.  Select “Security & Privacy”
3.  Find and select “Automation” on the left
4.  Check the checkbox for “Microsoft Word” under “Zotero”
5.  Restart Word

#### Word 2011#

If you’re running Word 2011, be sure you’ve updated to the latest version, 14.7.7, for compatibility with the new permissions system in Mojave and above.

### Zotero Word Processor Plugin Troubleshooting {#zotero-word-processor-plugin-troubleshooting}

This page describes some of the reported issues with the Zotero word processor plugins, together with possible solutions.

#### All Plugins/Platforms#
##### Zotero toolbar doesn’t appear#

In most cases, the Zotero plugin should appear automatically in the Word ribbon or the LibreOffice toolbar after installing Zotero and restarting the word processor. If you don’t see a Zotero tab or toolbar, follow these steps:

1.  Close Word or LibreOffice.
2.  Open the Zotero settings, click the Cite tab, and scroll down to the Word Processors section. Click “Install Microsoft Word Add-in” or “Install LibreOffice Add-in”, and then restart your word processor. If you get an error, see Troubleshooting Errors with Word Processor Plugin Installation.
3.  If the installation completes but the Zotero tab or toolbar still doesn’t appear in your word processor, follow the manual installation instructions.

If you still don’t see the plugin after performing a manual installation, follow these OS-specific troubleshooting steps:

-   Word for Windows
-   Word 2016+ for Mac
-   LibreOffice

##### Fixing broken documents#

If you can insert a Zotero citation in a new, empty document but get an error in an existing document, see Troubleshooting Errors in Word Processor Documents.

##### Formatting issues#

Citations and bibliographies generated by the word processor plugins might appear in a different style (font, font-size, etc) than the surrounding text. The appearance of the generated text can be changed by changing the default style. For example, in Word, open the Styles Manager in Format → Styles or by clicking the “Styles Pane” or “Manage Styles” buttons on the “Home” tab of the ribbon. In LibreOffice, open the Styles Manager in Format → Styles and Formatting or by pressing F12. Right-click on “Default”, select “Modify”, and make the desired changes to this style.

Bibliography formatting is controlled by the citation style you select in Zotero document preferences and should conform to the requirements of the style in use. The formatting of the bibliography can be modified by editing the “Bibliography” (Word) or “Bibliography 1” (LibreOffice) word processor style.

##### Field codes instead of citation/bibliography text#

See Field Codes.

##### Citations/bibliography highlighted#

See Citations Highlighted.

##### Citations converted to plain text#

See Existing Citations Not Detected.

#### Word#
##### Windows#
##### Error or non-responsive plugin buttons#

-   *“Zotero experienced an error updating your document.”*
-   *“Word could not communicate with Zotero. Please ensure Zotero is running and try again.”*

If you see one of the above errors in a **new, empty document** (and not just a specific existing document), or if no citation dialog appears when you click the plugin buttons, try the following steps, testing the plugin after each one.

1.  Close Word, restart Zotero, and then start Word.
2.  If the plugin still isn’t working, restart your computer. Start Zotero before starting Word.
3.  If the plugin still isn’t working, go to your Word Startup folder, delete Zotero.dotm, and restart Word to make sure that the plugin is completely gone from Word. (If the “Zotero” tab isn’t disappearing, see the steps in the linked section.) Once the “Zotero” tab is gone from Word, reinstall the plugin.
4.  Make sure that you’re running Zotero as the same user as Word. Specifically, check to make sure neither program is running as administrator (right-click → Properties → Compatibility). For security and stability reasons, you should typically not run any software as administrator.
5.  Temporarily disable any security software you’re running, which could interfere with the connection between Word and Zotero.
6.  If you’ve set up Zotero to use multiple Zotero profiles (most people haven’t), you might have configured Zotero to launch with the 0 command-line option. This will prevent the plugins from functioning, and if you’ve done this you should remove the 1 command-line option from the shortcut used to launch the Zotero profile. This flag should never be used with Zotero.

##### Zotero tab does not appear in the Word ribbon#

First, make sure you’ve tried the general troubleshooting steps.

Other things to try:

**Check if the Zotero plugin is correctly installed and enabled**

Go to File → Options → Add-ins and look for Zotero.dotm in the list:

-   If Zotero.dotm appears under Inactive Application Add-ins, select “Word Add-ins” in the Manage drop-down at the bottom, click “Go…”, and make sure Zotero.dotm is ticked.
-   If Zotero.dotm appears under Disabled Application Add-ins, select “Disabled Items” in the Manage drop-down at the bottom, and click “Go…”. In the window that pops up, select Zotero.dotm and click Enable.

**Check Trust Center**

1.  Go to File → Options → Trust Center and click “Trust Center Settings…” in the right-hand pane.
2.  Under “Add-ins”, make sure that “Require Application Add-ins to be signed by Trusted Publisher” and “Disable all Application Add-ins” are **unchecked**.
3.  Restart Word.
4.  If the Zotero tab is still not present, go back into the Trust Center Settings, open the “Macro Settings” pane, and select “Disable all macros with notification”. Restart Word and see if you get a notification asking for macro permissions. Click “Enable Content”.

**Check institutional policies**

If the Zotero tab still isn’t showing up and this is an institutional computer, you should contact your IT department, as they may be blocking Word templates from running or modifying the user interface (e.g., the Microsoft Intune policy “Disable UI extending from documents and templates”, which changes the 0 registry key).

**Reset Normal.dotm**

If you have modified you or someone else modified Normal.dotm template on your machine, the Zotero tab may fail to appear. You can reset it to default settings by temporarily removing it from its default location.

1.  Close Word.
2.  Press Windows+R, paste \0 into the dialog and click OK. An Explorer window will appear.
3.  Rename the \0 file to \1 or temporarily move it somewhere else, like your desktop.
4.  Restart Word and see if the Zotero tab appears.

##### Zotero tab has an empty label#

If after installing the plugin the Zotero tab has an empty label, it means Word Macro settings are preventing the Zotero plugin from running.

1.  In Word, go File → Options.
2.  Select the Trust Center panel and click the “Trust Center Settings…” button
3.  Select the Trusted Locations panel and make sure that the Word default location startup is in this list.

You can find your Word Startup location in Word Options → Advanced, by clicking the “File Locations…” button at the bottom of the panel.

##### Error “Could not find a running Word instance”#

See Could not find a running Word instance.

##### Run-time error ‘5097’: Word has encountered a problem.#

This issue affects the users of the Windows 10 October 2018 Update. To fix it you will need to change your regional format to English:

1.  In “Windows settings” go to the page for “Region”, which has settings for “Regional format”
2.  Change that setting to “English (United States)” or “English (United Kingdom)”
3.  Restart Word

##### “This command is not available because no document is open”#

Zotero’s Word add-on may not work for documents in OneDrive. If you encounter this error, move your document to a different folder. Other cloud-syncing services such as Dropbox or Google Drive are not affected. (Note that Zotero documents should not be opened or edited in Google Drive’s word processor as this will break Zotero citations. See this thread for details and developments.)

##### Mac#
##### Zotero tab does not appear in the Word 2016+ ribbon#

First, make sure you’ve tried the general troubleshooting steps.

Make sure you’re running the latest stable version of Word.

If the tab still does not appear, check whether the plugin was installed in Word correctly:

1.  Go to Tools → Templates and Add-ins.
2.  Make sure that “Zotero.dotm” is present under Global Templates and Add-ins and is checked.

If Zotero.dotm still doesn’t show up, you may not have correctly performed the manual installation steps to copy Zotero.dotm to your current Word Startup folder.

##### “Word could not communicate with Zotero. Please ensure that Zotero is open and try again.”#

First, make sure the Zotero app is open and running on your computer. Note that this is the Zotero program, not the Zotero website or Zotero Connector in your browser.

If Zotero is open and you’re still receiving this error, you likely don’t have the latest version of the Zotero plugin in Word. Go to your Word Startup folder, delete Zotero.dotm, restart Word, and confirm that the Zotero tab is gone from Word. If it doesn’t disappear, follow the instructions for identifying your *active* Word Startup folder, delete Zotero.dotm from there as well, and reset your Word Startup folder location to the default to avoid future problems. Once the Zotero tab is gone from Word, follow the steps under \#Zotero toolbar doesn’t appear.

If you modified your Zotero installation to work with multiple instances, you may need to do additional configuration. This does not apply to most users.

##### No response from plugin#

If you get no response when you attempt to use the Word plugin, try the following steps:

1.  Restart Word and try again.
2.  If the plugin still isn’t working in Word, go to your Word Startup folder, delete Zotero.dotm, and restart Word to make sure that the plugin is completely gone from Word. Then reinstall the plugin.

##### Run-time error ‘5’: Invalid procedure call or argument#

Some macOS audio plugins Voxengo SPAN or Voloco may interfere with the Zotero Word plugin. They can be found at 0. Try temporarily removing the plugins and test if the issue persists.

##### Clipboard software interference#

If you have any software installed that interacts with the MacOS clipboard, e.g. by providing clipboard history, it may interfere with the Zotero plugin. If you are having problems inserting citations, you should temporarily disable the clipboard software. If that doesn’t resolve the issue, temporarily remove the software and restart your Mac.

##### Linux#

The Zotero Word for Windows plugin does not work out of the box under WINE, CrossOver Office, or other compatibility environments for Linux. We do not support running the Word for Windows plugin under Linux, and advise users to use LibreOffice instead. If you absolutely must run Zotero in WINE, this forum thread has some helpful tips.

As of March 2016, the following steps should work getting Office 2010 run with Zotero.

1\. Install Office 2010 and Zotero on Wine

2\. Change directory to */home/%user%/.wine/drive\_c/users/%user%/Application Data/Microsoft/Word/STARTUP*/. Substitute “Zotero.dot” file with [this](https://drive.google.com/file/d/0B-FGaNOW8dnudEpuZDdnNnBVUmc/edit?usp=sharing) file

3\. Open Zotero.

4\. Open Microsoft Word.

Thanks to Sudarlin Laoddang for providing these instructions on his [blog](http://sudarlin.blogspot.pt/2015/03/running-zotero-ms-word-di-wine.html?showComment=1458032475297#c6126731001028222510).

#### LibreOffice#
##### All Platforms#

Zotero requires LibreOffice 5.2 or later. If you are using an older version, upgrade to the current version of LibreOffice. See System Requirements. Apache OpenOffice and NeoOffice are based on older versions of LibreOffice and are not supported by Zotero.

##### Installation error#

At the last step of LibreOffice Integration installation, you may see the message

*“Installation could not be completed because an error occurred.”*

See the installation troubleshooting instructions.

##### Toolbar is missing#

Check if there’s an entry for Zotero under View → Toolbars. If not, look for the Zotero LibreOffice Integration plugin in Tools → Extension Manager. If it’s in not there, return to Zotero and the Cite pane of the Zotero settings. In the “Word Processors” settings, click the “Install LibreOffice Add-in” button. If you get an error, refer to the installation troubleshooting instructions

##### Buttons are unresponsive after updating LibreOffice#

Unresponsive Zotero toolbar buttons are an indication of LibreOffice not having access to a working Java JRE. Refer to the installation troubleshooting instructions.

##### Add Extension(s)…does not exist error#

When attempting to manually install Zotero LibreOffice Integration, you may see the message

*Add extension(s): <DIRECTORY>/Zotero\_LibreOffice\_Integration.oxt does not exist*

We believe this is caused by a corrupt LibreOffice profile directory. Move or delete the [LibreOffice profile directory](https://wiki.documentfoundation.org/UserProfile), then follow the instructions above to reinstall the Zotero LibreOffice extension. (This will revert any LibreOffice settings you have customized to their default state.)

### “Could not find a running Word instance” {#could-not-find-a-running-word-instance}

This page only applies to the error “Could not find a running Word instance” when using the Zotero plugin in Word on Windows.

For general word processor troubleshooting, see Word Processor Plugin Troubleshooting.

Perform the following steps:

1.  Make sure you’re running **at least Zotero 7.0.3**, which includes a change that can avoid this issue.
2.  Make sure neither Zotero nor Word are running as administrator (right-click → Properties → Compatibility → “Run this program as an administrator” is unchecked) and that you’re not running Windows using the hidden Administrator account.
3.  If using OneDrive, save the copy of the document to your local hard drive, or try renaming the file to remove any spaces in the filename. OneDrive is known to interfere with the plugin for some people.
4.  Make sure you have [User Account Control](https://support.microsoft.com/en-us/windows/user-account-control-settings-d5b2046b-dcb8-54eb-f732-059f321afe18) enabled. You can change the UAC behavior by opening the Control Panel → System and Security → Change User Account Control settings. Make sure your level is at least “Notify me only when programs try to make changes to my computer (do not dim my desktop)” or higher, and **then restart your computer**.
5.  If re-enabling UAC and restarting doesn’t resolve the problem, you can try temporarily changing Zotero to run in Windows 8 compatibility mode. We’ve received several reports that this fixes the problem, and that compatibility mode can then be disabled without the problem recurring. To test, right-click on the Zotero shortcut, go to Properties → Compatibility, and set “Run this program in compatibility mode for” to “Windows 8”. Then restart Zotero and test the Word plugin. You should then disable compatibility mode, as permanently running a program under compatibility mode can result in lower performance, security vulnerabilities, and other unexpected problems.
6.  Temporarily disable any security software you’re running, which could interfere with the connection between Word and Zotero.

#### Post on the Zotero Forums#

If none of these steps help resolve the issue, please create a new thread in the Zotero Forums for further troubleshooting. Be sure to include a Report ID from Zotero, your operating system and Word versions, and the steps you’ve taken to try to fix the error.

## Collecting references {#collecting-references}
### Adding Files to your Zotero Library {#adding-files-to-your-zotero-library}

In addition to item metadata, notes, and tags, Zotero can also be used for managing files. This page describes the different ways you can add files to your Zotero library, and how added files are stored and synced.

#### Child versus Standalone Attachment Files#

Files can be added either as *standalone items* or as *child items* to regular Zotero bibliographic metadata items. It is generally always a good idea to work with files as child items. Standalone files cannot be used with many of Zotero’s features, including citing, My Publications, and most types of searching, because they lack bibliographic metadata.

If you save a PDF directly to your library, Zotero will attempt to retrieve metadata for it and create a parent item automatically. If the item can’t be recognized, you’ll be left with a standalone attachment. You can add a parent by either saving an item from the web and dragging the PDF on top of it (if a PDF isn’t attached automatically) or by right-clicking on the PDF, choosing Create Parent Item, and entering an identifier such as a DOI or ISBN. If all else fails, you can click Manual Entry after selecting Create Parent Item and manually enter metadata for the item.

#### Stored Files and Linked Files#

Files can be added to your Zotero library as either stored files or linked files.

##### Stored Files#

Stored files, which are the default, are stored within the Zotero data directory, and Zotero will automatically manage them, including deleting them if you delete the attachment item in Zotero. If you use file syncing, Zotero will automatically sync stored files between devices and make them available in your online library on zotero.org. If you add a stored file from a file on your computer, the file is copied to the Zotero data directory, so you may wish to delete the original to avoid confusion.

When using Zotero file syncing, you can choose to download stored files only as needed, avoiding the need to download all files to every device. An upcoming version of Zotero will allow you to choose how long to keep synced files on a given computer in order to limit disk space usage, temporarily redownloading files when you need them.

To use stored files outside of Zotero, you can use Zotero’s search and organization abilities to quickly find the relevant items and then either drag the attachments straight from Zotero (e.g., into an email) or right-click and choose Show File to view the files in your file manager. If you prefer to find files without going through Zotero, you can use your operating system’s search features (e.g., Spotlight on macOS) or create a smart folder in your OS to show a list of all PDFs within your Zotero data directory and interact with the files directly. Zotero automatically renames files based on the parent item’s bibliographic data, so you can easily find files by title, author, or year even from outside Zotero.

We strongly recommend using stored files for the most seamless experience.

##### Linked Files#

With linked files, Zotero only stores a link to the location of the original file on your computer. Linked files are not synced, nor are they deleted if the attachment item is deleted in Zotero. They also can’t be used within a group library, as there’s no guarantee that other group members would have access to the same file location. Linked files are also not supported in the [Zotero iOS app](https://apps.apple.com/us/app/zotero/id1513554812) or the upcoming [Android app](https://play.google.com/store/apps/details?id=org.zotero.android).

You can add a linked file by selecting an existing item and choosing “Attach Link to File…” from the Add Attachment menu in the Zotero toolbar or, to use PDF metadata retrieval, by selecting “Link to File…” from the New Item menu. You can also use the appropriate OS-specific modifier key for linking files while dragging in a file from the filesystem.

If you sync linked files using an external tool (Google Drive, etc.) for use on multiple computers, it is a good idea to set the linked attachment base directory so that the files can be found by Zotero on each computer even if the containing folder is at a different location in the filesystem.

Given the advanced nature of linked-file workflows, and the differences on individual systems, we’re not able to help troubleshoot problems with specific setups.

If you wish to convert linked files to stored files in order to allow Zotero to manage them, you can do so from the Tools → Manage Attachments menu.

#### Adding Files#
##### Adding Files via the Browser#

Zotero can automatically save associated web page snapshots and PDFs when you use the Zotero Connector save button in your web browser (whether associated snapshots and PDFs are saved can be changed in the Zotero preferences). Such snapshots and PDFs are stored as copies in Zotero data directory, and appear as child items of the saved item.

##### Adding Files via the Zotero window#
##### Drag and Drop#

Files can be copied into your library by dragging a file from your operating system’s file browser into the Zotero window, and either dropping it onto a collection in the left pane, or onto the center pane. Files dropped onto an existing regular Zotero item in the center pane are added as child items. Files dropped onto a collection, or in an empty space or between items in the center pane, are added as standalone items.

You can also drag and drop an existing standalone file item in Zotero onto a regular Zotero item to create a child item.

##### Adding linked files#

-   By default, files dragged into Zotero are added as **copies** of the original files. To instead add **links** to the original files, hold down 0+1 (Windows/Linux) or 2+3 (Mac) while dropping. (On macOS, it may be necessary to allow the Zotero window to come to the front before letting go for the modifier key to take effect.)

##### New Item Button#

File copies and file links can be created by clicking the “New Item” () button at the top of the center column and selecting “Store Copy of File…” or “Link to File…”, respectively. This creates standalone items.

##### Attachment Menu#

When you have selected a single item in the center pane, you can click the “Add Attachment” paperclip button at the top of the center column. Select either “Attached Stored Copy of File…” or “Attach Link to File…” to add files as attachments to the item.

You can also “Attach Link to URI…” to add a link to a web page (0 or 1) or to another program on your computer (e.g., OneNote 2 or Evernote 3).

These options are also available when you right-click an item and choose “Add Attachment”.

#### Accessing Files#

Files in your library can be accessed by double-clicking the item in the center pane. Alternatively, you can right-click the item and select “View PDF” or “View File”.

To locate a stored (copied) or linked file, right-click the item in the Zotero pane and select “Show File”. Copied files are stored in the Zotero data directory, and each file has its own subdirectory, which is named with a random 8-character string.

#### Web Snapshots#

Zotero can archive a webpage by creating a snapshot — an offline file reflecting the state of the page at the time the snapshot was taken. If the Zotero Connector does not recognize data on a page, you can save the page as a Web Page item with an attached snapshot by clicking the Zotero save button in the browser toolbar. You can also take a snapshot of any page by right-clicking (click-and-hold in Safari) on the Zotero save button and choosing “Save to Zotero (Web Page with Snapshot)”.

By default, Zotero will save snapshots when importing items from webpages. You can disable this setting in the Zotero preferences.

### Adding Items to Zotero {#adding-items-to-zotero}

This page describes the various ways to add items (e.g., books, journal articles, web pages, etc.) as items in Zotero. To learn more about adding files (such as PDFs or images), please see the files page.

#### Via your web browser#

**To use Zotero properly, you need to install the Zotero Connector for Chrome, Firefox, Edge, or Safari, in addition to the Zotero desktop app.**

The Zotero Connector’s save button is the most convenient and reliable way to add items with high-quality bibliographic metadata to your Zotero library. As you browse the web, the Zotero Connector will automatically find bibliographic information on webpages you visit and allow you to add it to Zotero with a single click.

For example, if you’re on the main page for a journal article, the save button will change to the icon of a journal article (circled in red):

On a library catalog entry for a book, the save button will show a book icon:

Clicking the save button will create an item in Zotero with the information it has identified.

On many sites, Zotero will also save any PDF accessible from the page or an open-access PDF that can be found for the saved item.

##### Generic Webpages#

Some webpages don’t provide any information that Zotero can recognize. On these pages, the save button will show a gray webpage icon. If you click the save button on these pages, Zotero will import the page as a “Web Page” item with a title, URL, and access date. See Saving Webpages below.

**Firefox:**

**Safari:**

##### PDFs#

If you are viewing a PDF file in your browser, the save button will show a PDF icon. Clicking this button will import the PDF file alone into your library and then automatically attempt to retrieve information about it. While this will often produce good results, it is usually better to use the save button from the publication’s abstract page or catalog entry, as described above, if there is one.

If you save a PDF directly and Zotero isn’t able to retrieve metadata, it will leave the PDF as a standalone attachment. To add metadata, you’ll need to create a parent item, either by saving a regular bibliographic item as described above and dragging the PDF on top of it or by right-clicking on the PDF, choosing Create Parent Item, and entering an identifier such as a DOI or ISBN. If all else fails, you can click Manual Entry after selecting Create Parent Item and manually enter metadata for the item.

##### Multiple Results#

On some webpages that contain information about multiple items (e.g., a list of Google Scholar search results), the save button will show a folder icon. Clicking this folder icon will open a window where you can select the items that you want to save to Zotero:

##### Saving to a Specific Collection or Library#

After you click the save button, a popup will appear indicating which Zotero collection the item is being saved to. If you want to save the item to a different collection or library, you can change the selection there, as well as enter tags to assign to the new item.

##### Data Quality and Choosing a Translator#

The quality of the data Zotero imports is determined by the information supplied on the webpage. Some websites include high-quality data for tools like Zotero in the page itself (“embedded metadata”). Other websites provide only limited metadata (e.g., only the title of a blog post) or no metadata at all. For many sites, Zotero has website-specific “translators” to obtain the best quality metadata. Zotero recognizes almost all library catalogs, most news sites, research databases, and scientific publishers. (For more information, see our compatible websites list.) Metadata for the same item may vary in quality across sites providing it. For example, importing an item from the publisher’s website will generally yield much better data than importing from Google Scholar.

Zotero will generally choose the best translator available for each site automatically. You can choose an alternative translator by right-clicking on the Zotero save button (or the page background in Safari) and choosing one of the available options. If a website isn’t importing properly, please report it on the Zotero Forums and provide the webpage URL.

#### Add Item by Identifier#

You can quickly add items to your library if you already know their ISBN, DOI, PubMed ID, arXiv ID, or ADS Bibcode. Click the Add Item by Identifier button () in the toolbar, type or paste in the identifier, and press Enter/Return. To add more than one item, separate identifiers by spaces, commas, or line breaks.

To look up metadata, Zotero uses Library of Congress, [WorldCat](http://www.worldcat.org/), and other catalogs for ISBNs, [CrossRef](http://www.crossref.org/) and other registries for DOIs, [NCBI PubMed](http://www.ncbi.nlm.nih.gov/pubmed/) for PubMed IDs, [arXiv.org](https://arxiv.org/) for arXiv IDs, and [ADS](https://ui.adsabs.harvard.edu/) for ADS Bibcodes.

#### Adding PDFs and Other Files#

As explained above, when possible, we recommend saving items using the Save to Zotero button in your browser from the primary webpage (e.g, a journal article’s abstract page) rather than adding PDFs directly. The Save to Zotero button will usually save high-quality metadata and also automatically download the relevant PDF if you have access to it.

If there’s no primary webpage, you can click the Save to Zotero button while viewing the PDF in your browser to save the PDF directly.

If you have a local PDF or other file on your computer — for example, if you received a file via email — you can drag it to Zotero, either onto an existing item to create a child attachment or between items to create a standalone attachment. You can also add an attachment by clicking “Add Attachment” in the Zotero toolbar and choosing one of the options.

##### Standalone Attachments and Parent Items#

Attachments can be either child items or standalone attachments. Standalone attachments can’t have bibliographic metadata or child notes, so in most cases you’ll want to convert them to child items under regular parent items.

When you add a PDF directly, Zotero will initially save it as a standalone attachment and then automatically attempt to retrieve metadata for it and create a parent item. This should work well for most academic PDFs (though it may sometimes yield lower-quality metadata than using the Save to Zotero button on the article page). For other documents, while Zotero can sometimes extract basic information (title, author), you shouldn’t expect that — anything can be distributed as a PDF, but that doesn’t mean there’s any standard metadata available for it.

If Zotero isn’t able to retrieve metadata for the PDF, you’ll be left with just the standalone attachment. You have a few options:

-   If you can find a source for metadata online, you can save a regular bibliographic item by using the Save to Zotero button on the article page and drag the attachment item onto the new item.
-   If you have a DOI, ISBN, or other identifier, you can right-click on the attachment item, choose Create Parent Item, and enter the identifier to retrieve metadata.
-   If all else fails, you can click Manual Entry in the Create Parent Item window to enter metadata manually.

#### Saving Webpages#

With Zotero, you can create an item from any webpage by clicking the save button in the browser toolbar. If the page isn’t recognized by a translator, you’ll see the gray webpage icon. If the page does have a recognized translator, you can force Zotero to save a Web Page item instead by right-clicking (click-and-hold in Safari) on the Zotero save button and choosing “Save to Zotero (Web Page with/without Snapshot)”

**Firefox:**

**Safari:**

If “Automatically take snapshots when creating items from web pages” is enabled in the General tab of the Zotero preferences, a copy (or snapshot) of the webpage will be saved to your computer and added as a child item. You can also save a snapshot with this setting disabled by right-clicking (click-and-hold in Safari) on the Zotero save button and choosing the relvant option. To view the saved copy, double-click the item or the snapshot in Zotero.

Double-clicking a Web Page item without a snapshot in your library will take you to the original webpage. Double-clicking a Web Page item with a snapshot will display the snapshot instead. You can also visit the original webpage by clicking the ”URL:” label to the left of the 0 field in Zotero’s right-hand pane.

#### Importing from Other Tools#

See Importing from Other Reference Managers.

#### Large-Scale Imports from Databases#

If you are importing a large number of items from scholarly databases (e.g., if you are conducting a systematic review), databases such as Google Scholar, ProQuest, Web of Science, and others, may lock you out if you use the Zotero save button too frequently or with too many items at once. In such cases, it is better to export the items as a batch in one of the standardized formats listed above (e.g., BibTeX and RIS are common choices) and import this file into Zotero. Web of Science and ProQuest offer the ability to select multiple items from a search results list and export as a batch to various formats. In Google Scholar, you need to first save the items to your Google Scholar library (using the ☆ icon in the search results), then select and export them from the Google Scholar “My Library” page.

#### Manually Adding Items#

Zotero is designed to help you avoid manual entry whenever possible. As a rule, you should save items to Zotero via your web browser rather than creating them manually. When you save from the web, Zotero will automatically extract high-quality metadata and download PDFs when available, saving you time and reducing errors. Even if you need to make manual corrections, it’s best to start with the version that Zotero saves rather than creating an item completely from scratch.

But if you really need to add something manually — for example, a source that isn’t available anywhere online — you can do so by clicking the green “New Item” () button at the top of the middle pane and selecting the desired item type from the drop-down menu. (The top level of the menu shows recently created item types. The complete list of item types, minus Web Page, can be found under “More”.) An empty item of the selected item type will now appear in the center column. You can then manually enter the item’s bibliographic information via the right-hand pane.

**Note:** Since it’s almost always better to visit a webpage in your browser and use the “Save to Zotero” button, the Web Page item type is not included in the “New Item” menu. However, if you really want to create a webpage item by hand, you create an empty item of another type and switch the item type to Web Page in the right-hand pane.

#### Editing Items#

When you have selected an item in the center pane, you can view and edit its bibliographic information via the Info tab of the right-hand pane. Most fields can be clicked and edited. Changes are saved automatically as they are made. Some fields have special features, which are discussed below.

##### Names#

Each item can have zero or more creators, of different types, such as authors, editors, etc. To change the creator type, click the creator field label (e.g., 0). A creator can be deleted by clicking the minus button at the end of the creator field, and additional creator fields can be added by clicking the plus button at the end of the last creator field. Creators can be reordered by clicking a creator field label and selecting “Move Up” or “Move Down”.

Each name field can be toggled between single and two field mode by clicking the “Switch to single field” / “Switch to two fields” buttons at the end of the creator field. Single field mode should be used to institutions (e.g., when the author is “Company A”), while two field mode (last name, first name) should be used for personal names. If a person has only one name (e.g., “Socrates”), enter this as a Last Name in two field mode. You can switch the order of two field author names by right-clicking on the name and choosing “Swap First/Last Names”

To quickly enter additional creators, type Shift-Enter/Retun to move immediately to a new creator field.

##### Journal Abbreviations#

Journal articles are often cited with the abbreviated journal title. Zotero stores the journal title and journal title abbreviation in separate fields (“Publication” and “Journal Abbr”, respectively). While some citation styles require different abbreviations, most of the variation is in whether or not the abbreviation contain periods (e.g., “PLoS Biol” or “PLoS Biol.”). Because removing periods is more accurate than adding them, we recommend that you store title abbreviations in your Zotero library with periods. Zotero can then reliably strip out the periods in rendered bibliographies when the chosen citation style calls for it.

##### Titles#

We recommend that you always store titles in your Zotero library in sentence case. See Sentence Casing for more information.

##### Links#

Clicking the label of the URL (“URL:”) and DOI (“DOI:”) fields will open up the (DOI-resolved) URL in your web browser.

##### Extra#

The Extra field can be used for storing custom item metadata or data that doesn’t have a dedicated field in Zotero. If you need to cite an item using a field not supplied by Zotero, you can also store such data in Extra. See Citing Fields from Extra for more details on how to cite these fields. For example, to add a DOI to a Book Section item, add this to the top of Extra: 0

#### Verify and Edit Your Records#

**When using Zotero — or any other reference manager — for citing, you should always check items for accuracy after saving them to your library.**

Zotero will accurately import metadata supplied by most bibliographic databases, library catalogs, publisher sites, and webpages. It will even make adjustments to the metadata to compensate for known quirks (e.g., author names in all upper case) in what the supplier provides.

That said, sometimes the metadata that Zotero receives is incomplete or incorrect. For example, one major academic search site often provides the wrong serial name with otherwise correct metadata. Another scholarly research site’s metadata can omit some of the authors’ names or present them in the wrong order. Even major publishers sometimes omit important metadata fields.

Some metadata is provided with only author last names and one or two initials when the authors’ full names are provided on the full-text version of the article. (For author names to be properly disambiguated in author-date styles, the author’s name must be consistently and identically entered across all items they contributed to.)

Publishers have different conventions for the casing of titles. No software can accurately and reliably convert title case to sentence case, so you should always store titles in sentence case and let Zotero convert them to title case as necessary.

You should be aware of these issues and verify that the items in your library are accurate and in the correct format so that Zotero can produce well-formed citations. One of the primary benefits of using a reference manager is that, once you’ve corrected item data once, your citations will always be correct going forward, in any citation style, no matter how many times you cite them.

If you do consistently receive incorrect information from a particular source, you should report it — with an example URL or identifier — in the Zotero Forums, as Zotero developers may be able to update Zotero to automatically correct the incorrect data.

### Can I highlight and annotate PDFs with Zotero? {#can-i-highlight-and-annotate-pdfs-with-zotero}
#### Can I highlight and annotate PDFs with Zotero?#

Zotero will open PDFs with your computer’s default PDF viewer. You can annotate PDFs with a variety of programs, including Adobe Reader/Professional, FoxIt, Nuance PowerPDF, PDF-XChange, macOS Preview, Microsoft Reader. So long as the program saves the annotations to the PDF file, the annotations will sync with zotero.org and your other devices.

### Default translators {#default-translators}
#### Default translators#

Zotero has four translators that attempt to find useful bibliographic data on pages that are not recognized by any of the more specific site translators. You can tell what translator detected bibliographic data on a page by placing the cursor over the document icon in the address bar; the name of the translator will show up as a tooltip. A similar name is usually saved to the “Library Catalog” field of created items.

-   DOI. Zotero tries to detect one or more pieces of text that could be DOIs. The DOI translator provides fairly passable support for many academic databases and sites that don’t have a dedicated Zotero translator. Metadata for DOIs is retrieved through [CrossRef](http://www.crossref.org); the data never includes full text or abstracts. Sometimes DOI-based saves will fail because of incomplete data in the CrossRef database.
-   COinS can include information on a limited number of item types. It never supports full text.
-   Embedded Metadata and unAPI *can* support all of Zotero’s item types and fields, as well as full text attachments. Embedded Metadata, however, will often detect minimal metadata on webpages, and the items it saves from such pages are frequently not very useful.

For details on how Zotero uses default translators, see the developer’s documentation.

### DOI format in APA style {#doi-format-in-apa-style}
#### DOI format in APA style#

Since the Digital Object Identifier (DOI) system was introduced in 2000, guidelines on how to best display DOIs have evolved. Effective March 2017, Crossref, an influential DOI registration agency, now [recommends](https://www.crossref.org/display-guidelines/) the following format:

0

Note the use of “https” instead of “http”, and “doi.org” instead of “dx.doi.org”.

While the 6th edition of the [Publication Manual of the American Psychological Association](http://www.apa.org/pubs/books/4200066.aspx), published in 2010, recommended the original 0 format, APA has [updated their guidelines](http://blog.apastyle.org/apastyle/2017/03/doi-display-guidelines-update-march-2017.html) to follow the Crossref guidelines for displaying DOIs with this latest format.

The CSL styles for the 6th edition of APA has therefore been updated to use the 0 format as well.

#### A brief history of DOI formats#

As Crossref [explains](https://www.crossref.org/display-guidelines/#why-not-use-doi-or-doi) in their guidelines, the original concise 0 format was recommended with the hope that web browsers would one day automatically recognize and hyperlink these DOIs.

Since this didn’t happen, Crossref [changed its recommendation](https://www.crossref.org/news/2011-08-02-crossref-revises-doi-display-guidelines/) in 2011 to show DOIs as regular URLs instead, in the format 0. APA started to recommend this newer (but nowadays also outdated) format in its 2012 companion guide, [APA Style Guide to Electronic References](http://www.apastyle.org/products/4210512.aspx), and discussed this change in a 2014 [blog post](http://blog.apastyle.org/apastyle/2014/07/how-to-use-the-new-doi-format-in-apa-style.html).

More recently, there has been a strong movement to move the web over from HTTP to the more secure HTTPS protocol. Technical changes also made it possible to link DOIs via the shorter 0 instead of 1. Together, this let Crossref [change](https://www.crossref.org/blog/new-crossref-doi-display-guidelines-are-on-the-way/) its recommended format to 2.

### How can I import from Citavi? {#how-can-i-import-from-citavi}

The best format for importing records from Citavi into Zotero is Citavi XML. This format will import the item bibliographic metadata, as well as quotations, tasks, attachments, etc. Beginning in Zotero 6, Zotero will also import PDF annotations. Zotero can import files exported from Citavi 5 and Citavi 6.

#### Citavi XML format (recommended)#
##### Special case: Cloud project#

If your project is a cloud project, you need to make a local copy of it, as preparation for the export:

1.  In Citavi click on 0 -> 1 -> 2
2.  In the following dialog, click on 0, put a meaningful name (e.g. *myproject\_local*) in the 1 field and click on 2.
3.  In the next dialog, click on 0 and click on 1.
4.  As soon as the local copy is finished, the local project will open in a new Citavi window. From there, resume with the following procedure.

##### Local project#

1.  In Citavi click on 0 -> 1 -> 2
2.  The backup is saved in a file with the file ending with 0 or 1 and lies normally in your [home directory](https://en.wikipedia.org/wiki/Home_directory) under 2 or 3 (i.e., follow this format 4). Alternatively, you can look up the backup folder under Tools / Options / Folders and there click on “Open folder with Windows Explorer”, see 5.
3.  The 0 or 1 file is a ZIP file, which cannot be processed directly — you must first unzip it and continue with the resulting unzipped file. To unzip it, either change the extension to .zip or use software dedicated to archiving/unarchiving files, such as 7zip. Note that file extensions are often hidden by the operating system, so to change the extension you might need to enable extension visibility.
4.  To import attached files (e.g., PDFs) into Zotero, you have to make sure that they are in the same folder as the 0 or 1 file you are importing. Citavi will save all attachments in the project folder (e.g., 2). Copy them from there.
5.  Import the 0 or 1 file in Zotero.

#### Limitations#

-   Nested collections in Citavi will not be nested anymore in Zotero. However, the numbers for the collections names should represent the nesting and adjusting this then manually should be easy.
-   To import PDF annotations, you must be running Zotero 6.

#### RIS format (alternative)#

As an alternative to the above procedure, you can export your Citavi library in the RIS format. This format will only retain item bibliographic metadata. Any additional item elements, such as citations from the “knowledge management” functionalities in Citavi, will not be included.

### How can I quickly switch between Zotero and my browser, PDF viewer, and/or word processor? {#how-can-i-quickly-switch-between-zotero-and-my-browser-pdf-viewer-andor-word-pro}
#### How can I quickly switch between Zotero and my browser, PDF viewer, and/or word processor?#

While you may be tempted to minimize windows or move them around on your screen to view the windows behind them, that’s almost never the fastest option.

To quickly move between open programs, **leave the foremost window where it is** and try one of the following methods:

#### Mac#

**Dock:** Click the icon of the program you want to switch to in the Dock.

**Cmd-Tab:** Press and hold Cmd, press Tab one or more times to highlight the program you want to switch to, and then let go of both keys. If you press Tab too many times while still holding Cmd, you can press \` (tilde, the key above Tab), Shift-Tab, or the left-arrow key to move in the other direction. While the application switcher is showing (i.e., after pressing Tab once with Cmd held down), you can also select the desired program with the mouse. If you know the program you want to switch to was the most recently used program, you can press Cmd-Tab and immediately release both keys to switch to the other program without waiting for the application switcher to show, and then repeat this to switch back.

If you’d prefer not to have a program’s window visible in the background, you can press Cmd-H to hide the program while still being able to switch to it with Cmd-Tab. For some programs, like Zotero, you can also close the main window without quitting the program and reopen it by using Cmd-Tab followed by Cmd-0 or by clicking the program’s Dock icon. (If you minimize a window, it’s no longer accessible via Cmd-Tab, though you can still open it via the Dock icon.)

**Mission Control:** If you only have a few windows open and are using a Mac with a trackpad, you can swipe up on the trackpad with four fingers to activate Mission Control, which displays all open windows, and then click on the window you want to switch to. (If this isn’t working for you, check Trackpad settings in System Preferences.)

#### Windows/Linux#

**Taskbar:** Click the icon of the program you want to switch to in the taskbar.

**Alt-Tab:** Press and hold Alt, press Tab one or more times to highlight the program you want to switch to, and then let go of both keys. If you press Tab too many times while still holding Alt, you can press Shift-Tab to move in the other direction. If you know the program you want to switch to was the most recently used program, you can press Alt-Tab and immediately release both keys to switch to the other program, and then repeat this to switch back.

#### Linux#

**Launcher:** Click the icon of the program you want to switch to in the launcher.

**Alt-Tab:** Press and hold Alt, press Tab one or more times to highlight the program you want to switch to, and then let go of both keys. If you press Tab too many times while still holding Alt, you can press Shift-Tab to move in the other direction. If you know the program you want to switch to was the most recently used program, you can press Alt-Tab and immediately release both keys to switch to the other program without waiting for the application switcher to show, and then repeat this to switch back.

### How do I import a Mendeley library into Zotero? {#how-do-i-import-a-mendeley-library-into-zotero}

Zotero can directly import all data, including the full folder structure, from an online Mendeley library.

To import your Mendeley library, follow these steps:

1.  Make sure that all data and files have been synced to Mendeley servers.
    -   If you use **Mendeley Reference Manager**, your data and files are already all online.
    -   If you use **Mendeley Desktop**, check your sync settings to make sure that data and files are being synced, and confirm that you can open PDFs in your online Mendeley library.
2.  Make sure you’re running the latest version of Zotero (Help → Check for Updates…).
3.  Go to File → Import within Zotero and choose the “Mendeley Reference Manager (online import)” option.

You’ll be asked to log in to Mendeley to allow Zotero to perform the import. Your Mendeley password is never seen or stored.

#### Alternative Method#

If for some reason you’re not able to perform a direct online import, it’s possible to import from a local Mendeley database by installing an old version of Mendeley Desktop from before Mendeley began encrypting the local database. All your data and files will still need to be synced to Mendeley servers. See the local import instructions for more information.

#### Using Mendeley Citations#

Zotero’s Word and LibreOffice plugins can read citations created by Mendeley Desktop and automatically relink them to imported Mendeley items in your Zotero library, so you can continue using the same documents with Zotero.

Citations created with Mendeley Cite are not readable by Zotero.

Note: Prior to Zotero 6.0.19 (released in December 2022), Zotero could read and use Mendeley citations, but they wouldn’t be linked to imported items in your Zotero library. If you imported your Mendeley library using Zotero 6.0.18 or earlier, you’ll need to repeat the import process once using Zotero 6.0.19 or later — just select “Relink Mendeley Desktop citations” when starting the importer. (If you’ve already imported citations using 6.0.19 or later, this option will no longer appear.)

#### Known Issues#

There are a few issues to be aware of when importing from Mendeley.

-   If your Mendeley account requires institutional credentials to log in, you may need to create a separate Mendeley account and connect it to the institutional account, and then use the personal account to log in. See [How do I use institutional (Shibboleth) credentials with Mendeley?](https://service.elsevier.com/app/answers/detail/a_id/33535/supporthub/mendeley/p/16075/) on the Mendeley support site for more information.
-   It’s not possible to directly import group libraries. To import items in group libraries, copy the group items to a collection in your Mendeley library before importing. You can then create a Zotero group and drag imported collections or items to that group.
-   Mendeley allows any field to be added to any type. When importing into Zotero, if a field isn’t valid for a given item type, the field is placed into the Extra field. When possible, those will be used automatically in citations (e.g., Original Date), and future versions of Zotero will automatically convert those to any real fields that become available.

#### Troubleshooting#

Make sure you’re running the latest version of Zotero available via Help → “Check for Updates…”.

If you’re running the latest version and something doesn’t come through how you expect or you run into any trouble, let us know in the Zotero Forums.

#### Mendeley Database Encryption#

The importer described above imports data directly from an online Mendeley library, which requires all data and files to be uploaded to Elsevier servers in order to be imported into Zotero.

Zotero originally announced work on a fully local importer in early 2018, but a few months later, Elsevier began encrypting the local Mendeley database, making it unreadable by Zotero and other standard database tools. This change came despite Mendeley having long touted the openness of their database format as a guarantee against lock-in and explaining in documentation that the database could be accessed using standard tools. Mendeley Desktop itself had imported data from Zotero’s own open database since 2009.

The [Mendeley 1.19 release notes](https://www.mendeley.com/release-notes/v1_19) claimed that the encryption was for “improved security” on shared machines, yet applications rarely encrypt their local data files, as file protections are generally handled by the operating system with account permissions and full-disk encryption, and someone using the same operating system account or an admin account can already install a keylogger to capture passwords. Mendeley later [switched to claiming](https://twitter.com/mendeley_com/status/1006915998841221120) that the change was required by new European privacy regulations — a bizarre claim, given that those regulations are designed to give people control over their data and guarantee data portability, not the opposite — and continued to assert, falsely, that full local export was still possible, while [repeatedly](https://twitter.com/mendeley_com/status/1006919608471818240) [dismissing](https://web.archive.org/web/20211120213432/0) reports of the change as “#fakenews”.

Direct access to the Mendeley database is the only fully local way to export the full contents of one’s own research. The export formats supported by Mendeley don’t contain folders, various metadata fields (date added, favorite, and others), or PDF annotations. While Mendeley offers a web-based API, it contains only uploaded data, so relying on it means that anyone wanting to export their own data first needs to upload all their data and files to Elsevier’s servers. The API is under Elsevier’s control and can be [changed](https://service.elsevier.com/app/answers/detail/a_id/31598/supporthub/mendeley/p/16075/) or [discontinued](https://mendeleyblog.wordpress.com/2021/03/11/mendeley-refocusing-announcement-mobile-app-retirement//) at any time.

Since making this change, Elsevier has replaced Mendeley Desktop with Mendeley Reference Manager, which is essentially a wrapper around the website and doesn’t contain a real local database at all.

### How do I import BibTeX or other standardized formats? {#how-do-i-import-bibtex-or-other-standardized-formats}

Zotero can import bibliographic data stored in a variety of standardized formats used by databases and other reference management tools. The most popular formats are RIS, Bib(La)Tex, and MODS.

If you have a database stored in one of these formats, such as a BibTeX database you’ve compiled or a RIS database you’ve exported from another reference manager, you can import them into Zotero by clicking File → “Import…” and choosing “A file”.

See also Moving to Zotero.

#### Zotero can import the following bibliographic formats#

-   Zotero RDF
-   CSL JSON
-   BibTeX
-   BibLaTeX
-   RIS
    -   Can be convenient for quick edits between export & import because of its simple structure
-   Bibliontology RDF
-   MODS (Metadata Object Description Schema)
-   Endnote XML
    -   Best format for exporting from Endnote
-   Citavi XML
    -   Best format for exporting from Citavi
-   MAB2
-   MARC
-   MARCXML
-   MEDLINE/nbib
-   OVID Tagged
-   PubMed XML
-   RefWorks Tagged
    -   Best format for exporting from RefWorks
-   Web of Science Tagged
-   Refer/BibIX
    -   Generally avoid if any other option is available
-   XML ContextObject
-   Unqualified Dublin Core RDF

### How do I import from EndNote? {#how-do-i-import-from-endnote}
#### Exporting Your Library from EndNote#

Zotero can’t directly import “.enl” EndNote libraries, so the first step is exporting your library from EndNote. The best export format for this is XML.

With older EndNote libraries, it may be necessary to convert figures to attachments before you export. This is done by going to the References menu → Figure and selecting “Convert Figures to File Attachments…”.

1.  If you wish to export a subset of your EndNote library, select the entries you wish to export.
2.  Go to the File menu → Export. A dialog box will pop up asking you where to save the export file.
3.  Navigate to your EndNote data directory (typically, My Documents\endnote.Data). This directory contains a ‘PDF’ folder, but you should be sure to select the data directory rather than any subfolder.
    -   **This is important!** Zotero will look for file attachments in a directory relative to the location of the exported XML file. If you save this file in the wrong spot, file attachments won’t be included when you import into Zotero.
4.  For “Save as type:”, choose “XML”.
5.  If you only want to export a subset of your library, check the “Export Selected References” box. Otherwise, make sure it is unchecked.
6.  Click “Save”.
7.  Close EndNote.

#### Importing into Zotero#

If you are not importing into an empty library, we **highly recommend** making a backup of you Zotero data directory. This can avoid frustration if you do not like the way your library has transferred. In that event, simply restore your library from the backup.

You should also temporarily disable automatic sync in Zotero’s Sync preferences. After you have imported your library and checked to be sure you are satisfied with the imported data, you can re-enable automatic sync.

In Zotero, click “Import…” in the File menu. A dialog box will appear asking you to select the file to import. Navigate to the location where you exported your EndNote library (if you followed the above instructions, this should be My Documents\endnote.Data) and select the .xml file. Click Open.

Note that, if Zotero encounters any fields in the EndNote XML data that it does not support (e.g., custom fields, author address, author affiliation), it will add these data to a note attached to the imported item. These notes will be tagged with “\_EndnoteXML import”. If the import adds many of these notes, Zotero’s performance can be negatively impacted. You should review each of these notes to determine if the data needs to be retained and delete any unnecessary notes. Additionally, you should check these notes to determine if any data could be migrated to proper Zotero fields (which is particularly important if you were using EndNote fields in non-standard ways).

You can quickly display all of the notes generated during import by clicking on the “\_EndnoteXML import” tag in the tag selector in the lower-left corner of the Zotero window. You can quickly delete all of these notes by selecting the tag in the tag selector, clicking in the items list and typing Cmd+A (Mac) or Ctrl+A (Windows/Linux) to select all matching items, and then right-clicking on a selected item and choosing “Move Items to Trash…”.

#### Getting Further Help#

If you have any issues related to importing and exporting references, feel free to ask for help on the Zotero Forums.

#### RIS format (alternative)#

Instead of EndNote XML, it is also possible to export items from EndNote using RIS. The only benefit of RIS over XML is that the EndNote database ID can be retained for each item. If this is required, then, in Zotero, open the Advanced pane of Zotero preferences and click the “Config Editor” button on the “General” tab. Search for “RIS.import.keepID” and double-click to set it to **true**. If you want data for unknown fields to be retained in notes (as described above), also search for “RIS.import.ignoreUnknown” and set this option to **false**.

In EndNote, export your library as above, setting “Save as type:” to “Text file (.txt)”. Set “Output style:” to “RefMan (RIS)”. Navigate to your EndNote data folder (as described above), save the export file, then import into Zotero as described above.

With this setup, RIS will retain EndNote database IDs, but any italics, bold, or other formatting set in fields will be lost. XML should always be used unless EndNote database IDs are needed for a specific reason.

### How do I import references into Zotero? {#how-do-i-import-references-into-zotero}
#### How do I import references into Zotero?#

If you are moving to Zotero from another reference manager, see Moving to Zotero.

For general use, the best method for adding items to Zotero is to use the Zotero Connector button in your web browser. See For information on using Zotero’s import functions to add items to your library, see the Getting Stuff into Your Library for more information on this and other methods.

For troubleshooting import and export issues, you can try the relevant steps in the translator troubleshooting procedure and direct questions to the Zotero forums.

### How do I turn off automatic case changes/capitalization during item import? {#how-do-i-turn-off-automatic-case-changescapitalization-during-item-import}
#### How do I turn off automatic case changes/capitalization during item import?#

By default, Zotero will try to clean up casing/capitalization of item titles when items are imported (e.g., titles in ALL CAPS will be converted Title Case). To disable this function, follow these steps:

1\) In the Advanced pane of Zotero preferences, on the “General” tab, click the “Config Editor” button.

2\) Find “extensions.zotero.capitalizeTitles” in the list. To find it quickly, you can type “capitalize” in the Filter box.

3\) Double-click “extensions.zotero.capitalizeTitles” to toggle between the true and false setting.

### How does the Import from Clipboard feature work? {#how-does-the-import-from-clipboard-feature-work}
#### How does the Import from Clipboard feature work?#

Import from Clipboard allows you to import items from the raw code of any supported format (RIS, BibTeX, CSL JSON, etc.).

Whenever you’re viewing raw bibliographic data, you can import it into Zotero simply by copying it to the clipboard and then choosing “Import from Clipboard” from the File menu in Zotero or by using the keyboard shortcut (Ctrl-Alt-Shift-I (Windows/Linux) or Cmd-Option-Shift-I (Mac)).

See Importing Standardized Formats for a list of supported formats.

Note that, on most websites, you should use the Zotero Connector’s “Save to Zotero” button to save data and files to Zotero rather than importing metadata directly. See Adding Items to Zotero for more information.

### I have bibliographies in Microsoft Word documents, PDFs, and other text files. Can I import them into my Zotero library? {#i-have-bibliographies-in-microsoft-word-documents-pdfs-and-other-text-files-can-}
#### I have bibliographies in Microsoft Word documents, PDFs, and other text files. Can I import them into my Zotero library?#
##### Citations inserted using Zotero or Mendeley Desktop#

Zotero can read existing citations created by the Zotero and Mendeley Desktop (not Mendeley Cite) word processor plugins, allowing you to continue using those citations in the same document even if the items don’t exist in your Zotero library. Simply click Add/Edit Citation, search for an existing citation, and select it from the Cited section of the search results. (This applies to the default citation dialog only, not the “classic” dialog.)

If a document contains Zotero or Mendeley Desktop citations not in your library and you need to make changes to the metadata or include them in other documents, you’ll need to extract the citations into your library. For Word .docx documents, you can use [Reference Extractor](https://rintze.zelle.me/ref-extractor/). Note that to continue using the same document, you’ll want to replace all instances of the original citation with the new item from your library, being sure to select from the library section of the citation dialog’s search results rather than the Cited section. (In an upcoming version of Zotero, it will be possible to relink orphaned citations without needing to reinsert them.)

If you still have the references in the reference manager, you can import them into Zotero:

-   Zotero has a built-in Mendeley importer that can import all data and automatically relink citations in existing documents.
-   For other programs, export your data in a format such as RIS or BibTeX and then import the file into Zotero. You’ll need to replace all existing citations in any document for which you want Zotero to generate a correct bibliography.

##### Citations inserted using Microsoft Word’s built-in citation feature#

If you used Word’s built-in citation feature, you can follow these steps to format the bibliography as BibTeX, which Zotero can import:

1.  Download this [Word bibliography stylesheet](https://gist.githubusercontent.com/JaimeChavarriaga/40166befb14f2fe5dac390688d9eaf03/raw/faf4aa3f72e553095f81f1440c3dce744c2755a2/bibtex.xsl).
2.  Save the stylesheet to Word’s bibliography styles folder:
    -   *Word 2016/2019/Office 365 for Windows:* 0
    -   *Word 2010 for Windows:* 0 or 1
    -   *Mac:* Go to the Applications folder. Right-click on Microsoft Word and choose “Show Package Contents”. Navigate to: 0
3.  In Word, change your bibliography style to “BibTeX export” and copy the bibliography to the clipboard.
4.  Use Zotero’s Import from Clipboard function.

To continue using the same document, you’ll want to replace all instances of the original citations so that Zotero can generate a correct bibliography.

##### Plain-text citations and bibliographies#

If the references have ISBNs, DOIs, or PubMed IDs, you can use the Add Item by Identifier function in Zotero to quickly add these items to your Zotero library.

If you have many references, you can use [AnyStyle](http://anystyle.io), an online bibliography parser written by a Zotero developer. Export parsed citations as BibTeX or CSL-JSON and import them into Zotero. Some people have also had success asking ChatGPT or other AI systems to parse bibliographies and generate BibTeX or CSL-JSON, though if you try this, you should check the output carefully for errors.

Otherwise, your best option is to find the items online in a repository that Zotero supports — most likely just by searching the web for the citation — and saving to Zotero with the Zotero Connector. This will help ensure that you get high-quality data that includes all fields that might be necessary in the citation style you’re using.

As a last resort, you can manually enter the references in Zotero.

In all cases, you’ll need to replace all existing citations in any document for which you want Zotero to generate a correct bibliography.

### Importing a Mendeley Library Into Zotero (Alternative Local Method) {#importing-a-mendeley-library-into-zotero-alternative-local-method}

Zotero can import from an online Mendeley library, and we recommend using that method if possible.

If for some reason you’re unable to perform an online import, it’s possible to import from a local Mendeley Desktop database, but you’ll first need to install an old version of Mendeley Desktop. Later versions of Mendeley Desktop began encrypting the local database, preventing you from getting your own data out of the app.

To perform a local import, follow these steps:

1.  Make sure you’ve synced all data *and* files to Mendeley servers in your current version of Mendeley. The only way to get existing data into a pre-encryption version of the app is by syncing it from Mendeley servers, and Mendeley sync doesn’t make any information about attached files available unless you actually sync the files themselves.
2.  If you’re already using Mendeley Desktop, make a [backup of your Mendeley database](https://service.elsevier.com/app/answers/detail/a_id/18153/).
3.  Close Mendeley Desktop
4.  Download and install Mendeley Desktop 1.18 using the links below.
5.  Open Mendeley Desktop 1.18 and perform a fresh sync to pull down your Mendeley data and files from the Mendeley servers. (If Mendeley doesn’t open, you may need to go to your Mendeley data directory and move the file ending with 0 out of the way.)
6.  Verify that all your data and files are available locally. If some files are unavailable, it may help to add all items in the library to a folder and restart Mendeley Desktop.
7.  Make sure you’re running the latest version of Zotero.
8.  Start the import in Zotero by going to File → “Import…”, choosing the file option, navigating to your Mendeley Desktop data directory, and selecting the file with the filename 0.

By default, the Mendeley data directory can be found at the following locations:

-   Windows: 0
-   macOS: 0
-   Linux: 0

#### Mendeley 1.18 installers#

*Mendeley 1.18 installers are no longer available from Elsevier.*

### Importing from Other Reference Managers {#importing-from-other-reference-managers}

It’s easy to migrate your data from other reference management tools to Zotero. Instructions for popular tools are linked below.

-   Mendeley
-   EndNote
-   Citavi
-   Microsoft Word Bibliography XML
-   Plain text reference lists
-   Bib(La)TeX
-   JabRef

You can also import from other tools, such as Reference Manager, RefWorks, Papers, Google Scholar Library, ReadCube, etc., by exporting to a standardized reference format, such as RIS, BibTeX, or CSL JSON, and then importing into Zotero by clicking File → “Import…” and choose “A file”.

If you already have data in a Zotero library on another computer, follow these instructions to transfer your library to your new computer.

### Is the Zotero web library the same as the Zotero desktop app? {#is-the-zotero-web-library-the-same-as-the-zotero-desktop-app}

The Zotero desktop app is the primary way of using Zotero and offers complete functionality. When people talk about “Zotero”, they’re almost always referring to the desktop app.

The Zotero web library is a complementary tool for accessing your Zotero data when you’re away from your main computer or using a platform that can’t run Zotero (locked-down institutional computers, some Chromebooks, etc.). The web library also allows you to share a Zotero library publicly with people who might not use Zotero.

A desktop app allows for vastly more functionality than a website. Here’s an incomplete list of features that only the Zotero desktop app provides:

-   A local database fully under your control, with optional syncing
-   Fast, offline access to all your data — whereas the web library can only load a subset of your data at a time, the desktop app allows you to access and modify everything much faster
-   Word, LibreOffice, and Google Docs plugins that let you insert citations and bibliographies directly from Zotero and keep them updated automatically
-   A better experience when saving items from the Zotero Connector (+ other features)
-   The ability to work with a large number of references — certain operations, like export, are limited to 100 items in the web library for technical reasons
-   More comprehensive export output with options to include files and notes in some formats
-   Real windows that you can interact with via Cmd-Tab/Alt-Tab, resize, etc. — not just an interface within a single browser tab
-   Local filesystem access — e.g., adding PDFs directly from your documents or creating links to files on the local disk
-   Find Available PDF to locate open-access versions of files
-   Importing of references from other reference managers
-   PDF full-text indexing (coming soon to web library)
-   Advanced search and saved searches
-   Retracted item notifications
-   My Publications collection to share your work publicly
-   CSL style editor
-   Custom citation styles
-   Unfiled items collection (coming soon to web library)
-   Duplicate items collection and item merging
-   An unrestricted plugin system that allows for many more features and further customization

### Known Translator Issues {#known-translator-issues}

Zotero uses “translators” to allow you to save items from the websites you visit. When you get the “An error occurred while saving this item” message while trying to save an item, first check the list of known translator issues below. If the translator you’ve been using is already listed, you don’t need to take any action (unless you have some programming experience and want to help fix the translator). Otherwise, try going through the steps for troubleshooting translator issues.

You can see which translator Zotero tries to use by hovering your cursor over the “Save to Zotero” icon in the address bar of your browser. The tooltip shows the translator name between parentheses.

#### Translators with Major Issues#

-   ASCE
    -   Currently completely broken.
-   DOI
    -   The DOI translator is a translator that scans webpages for [DOIs](http://www.doi.org/), and collects the item metadata by using Crossref’s [DOI lookup service](http://www.crossref.org/guestquery/). If the DOI translator gives an error, either the translator mistook something on the webpage for a DOI, or Crossref doesn’t (yet) have any item metadata available for the DOI.
-   EBSCO
    -   Doesn’t work for items in “My Folder”.
-   JSTOR
    -   Only saves PDFs after manually clicking “OK” to the terms and conditions once in the session. Workaround: Manually download one PDF during each session; all subsequent ones should work fine.
-   Old Baileys
    -   Frequently fails silently, due to an issue with google ads. Works reliably with an ad-blocker add-on enabled. See this thread for developments
-   Google Scholar \* Google Scholar will lock you out to protect its data against automated downloads when you use its service a lot (which may be as quickly as saving three pages of results). See \[\[ adding\_items\_to\_zotero#large-scale\_imports\_from\_databases\|here for a workaround\]\] for large downloads

#### Translators with Minor Issues#

-   Proquest
    -   When downloading large amounts of articles from search results, ProQuest starts to require users to input a captcha to continue to articles. Once that point is reached, grabbing items from search results is no longer possible with Zotero. Individual items will continue to work.

### Proxies {#proxies}

We’re in the process of updating the documentation for Zotero 5.0. Some documentation may be outdated in the meantime. Thanks for your understanding.

*This feature is not available in Safari.*

This page describes the Proxies tab of Zotero Connector preferences. Zotero users should be able to make complete use of the proxies feature without ever looking at this tab. By default, Zotero will prompt you to store the proxy and then route you through the proxy automatically and without further input.

The Proxies preferences allow you to adjust the following options:

#### Enable proxy redirection#

Default: checked - unchecking this will disable Zotero’s proxy redirection feature. You can do this temporarily and your proxy settings will remain saved. Do not use this option if you no longer have access to the proxies saved in Zotero. In that case, delete those settings by selecting them in the “Configured Proxies” box and pressing the minus (-) button below it.

#### Automatically detect new proxies#

Default: checked - unchecking this will prevent Zotero from automatically detecting and storing proxies it detects.

#### Disable proxy redirection when domain name contains…#

Default: unchecked - typically you won’t need to use a proxy when you are connected to the internet through your institution’s network. This option automatically disables Zotero’s proxy redirection when the domain of your internet provider contains the given string. In the United States, “.edu” (the default setting) will usually work, in other countries you will have to find out your institution’s domain name.

#### Configured Proxies#

Proxies can be added manually by clicking on the “+” button. This section allows to specify the URL of the database being accessed under hostname and the URL scheme of the proxy. Existing proxies can be edited by selecting them in the configured proxies box. You can remove proxies by clicking the “-” button.

#### Troubleshooting#

Having trouble accessing a site due to Zotero’s proxying functionality? See Proxy Troubleshooting.

### Retrieve PDF Metadata {#retrieve-pdf-metadata}

Users new to Zotero may find the prospect of importing all their data somewhat daunting. Many researchers already have a large collection of PDFs that they’ve previously organized manually. Zotero makes it easy to import these PDFs and retrieve full bibliographic metadata (for searching, citing, indexing, and organizing), taking much of the pain out of switching.

To use this feature, simply drag your existing PDFs into your Zotero library or use the “Store Copy of File” or “Link to File” options from the add new item menu (green plus sign). By default, Zotero will automatically retrieve metadata for each PDF, create an appropriate parent item, and rename the associated file based on the metadata. (You can disable these automatic functions in the General pane of Zotero preferences.)

If Zotero can find a match for the PDF, it will create a full Zotero item with the available data and attach the PDF. If it can’t, it will leave the PDF as a standalone attachment, allowing you to add a parent item another way — either by saving an item from the web and dragging the PDF on top of it or by right-clicking on the PDF, choosing Create Parent Item, and entering an identifier such as a DOI or ISBN. If all else fails, you can click Manual Entry after selecting Create Parent Item and manually enter metadata for the item.

If you’re not happy with the metadata saved for the PDF, you can right-click on the new parent item and choose Undo Retrieve Metadata to leave the PDF as a standalone attachment for further manual processing.

Zotero should retrieve high-quality metadata for most academic PDFs. While it can sometimes extract basic information (title, author) from other documents, you shouldn’t expect that — anything can be distributed as a PDF, but that doesn’t mean there’s any standard metadata available for it.

**Note:** While this feature can greatly facilitate importing large existing libraries of PDFs, it **is not** the best way to add items to your library in general. Items and PDFs can be imported faster by using the Zotero Connector plugin in your browser from the article pages (not the PDFs) of publisher websites or most scholarly databases. This saves several steps versus downloading the PDF manually and adding it to Zotero. The item metadata will also often be higher quality. See Adding Items to Zotero for more info.

#### How It Works#

The Retrieve Metadata feature uses a Zotero web service to find item metadata. The Zotero client sends the first few pages of text from the PDF to the web service, which uses a variety of extraction algorithms and known metadata from Crossref, paired with DOI and ISBN lookups, to build a parent item for the PDF. The Zotero lookup service doesn’t require a Zotero account and doesn’t log any data about the content or results of searches.

### Troubleshooting Problems Saving to Zotero {#troubleshooting-problems-saving-to-zotero}

July 2025: If you're a Firefox user experiencing problems saving on many sites, make sure you're running Zotero Connector 5.0.169 or later. If you still have 5.0.151, click the application menu in the top-right corner of Firefox and accept the updated permission or reinstall the Zotero Connector from the download page.

When you’re unable to save high-quality metadata from a particular website, most likely either the page isn’t supported by an existing Zotero translator or the page layout recently changed, breaking Zotero’s ability to recognize data. On some sites, most notably Google Scholar, you may also be running into site access limits.

If the problem occurs across multiple sites, there may be a problem with your installation, and you should try these steps:

1.  If you don’t see a “Save to Zotero” button in your browser toolbar at all, **make sure you’ve installed the Zotero Connector** in your browser. The Zotero Connector should be listed in the browser’s Extensions pane. You can install the appropriate Zotero Connector from the Zotero download page. If the extension is installed but not visible, you may need to pin it to your toolbar.
2.  If you only see a gray webpage icon, **make sure you’re looking at a supported site**. If you’re not sure, try a Wikipedia article or an Amazon book page.
3.  If you’re looking at a supported site, **try reloading the page** and **make sure the page has fully loaded**. Pressing your browser’s stop button or pressing Esc on your keyboard can help if something on the page has stalled.
4.  If the gray webpage icon appears on all sites, **try restarting your browser**.
5.  If you don’t have the Zotero program open, **try opening Zotero before saving**. (If you haven’t installed Zotero, you can do so from the download page.) While the Connector can save directly to your online library, you’ll usually get better results saving to Zotero directly. If the Zotero Connector reports that Zotero is unavailable and tries to save to zotero.org, see Zotero Unavailable.
6.  **Check for Zotero Connector updates** from the Extensions pane of your browser, in case a new version of the Connector is available that you haven’t yet received.
7.  Make sure you have an **up-to-date version of your browser**. See System Requirements for supported browser versions.
8.  If you’re saving with Zotero open, **check your Zotero version** to make sure you have the latest version available from zotero.org. We can only troubleshoot translators that fail for the current 7.0.x version of Zotero.
9.  Check **Known Translator Issues** to see if problems have been noted on the site you’re trying to save from.
10. **Hover your cursor over the “Save to Zotero” button.** The tooltip shows which translator, if any, Zotero will use for the page you are viewing. If the displayed translator is incorrect, post the site name, a URL, and the name of the incorrectly detected translator to the Zotero Forums. Note that “Web Page”, “Embedded Metadata,” “DOI,” and “COinS” are generic translators, and Zotero may not be able to save full metadata or PDFs from all sites on which they appear. A problem with a primary translator may cause a generic translator to be used instead.
11. Make sure you’ve given the Zotero Connector **permission to access all websites**. In Chrome or Edge, right-click on the Save to Zotero button, select “This Can Read and Change Site Data”, and make sure it’s set to “On All Sites”. In Safari, go to the Websites tab of the Safari settings, select Zotero Connector under Extensions, and make sure “For other websites” is set to “Allow”. The Zotero privacy policy explains why this is necessary.
12. If you are having chronic problems getting the Zotero Connector to work across multiple sites, you may have an extension conflict. Try **uninstalling and reinstalling the Zotero Connector**. If that doesn’t help, try **disabling all extensions** except the Zotero Connector. If this solves the problem, re-enable the extensions one-by-one until you find the conflict, and then post the name of the extension that was causing the issue to the forums.
13. In rare cases where you’re seeing a webpage icon on all sites or saving is failing on all sites, there may be a problem with your translators. Disable all third-party plugins, restart Zotero, and select “**Reset Translators**” from the Advanced pane of the Zotero preferences (if applicable), and then do the same for the Zotero Connector (right-click on the save button → Options/Preferences → Advanced tab). This isn’t necessary as a regular troubleshooting step.
14. **If none of these solutions solve your problem**, we’ll need additional information to further debug your issue. Please create a new thread in the Zotero Forums — or use your existing thread if you’ve already created one — and provide the following:
    -   The **exact URL of a page that isn’t working** (even if none are)
    -   A **Debug ID** from the Zotero Connector for reloading the page and attempt to save or, if you’re only getting a webpage icon, for reloading the page.
    -   What it says in the popup when you try to save (e.g., “Saving to My Library” or “Saving to zotero.org”)
    -   The **name of the translator** from step 10, if applicable

### What are these DOIs doing in my bibliography? {#what-are-these-dois-doing-in-my-bibliography}
#### What are these DOIs doing in my bibliography?#

The Digital Object Identifier ([DOI](http://www.doi.org/)) is a unique identifier used to provide a permanent link to digital objects. DOIs are now issued to most (new) journal articles, as well as many books, book chapters, and other items.

Several professional association style guides now specify the use of the DOI when citing journal articles. For example, [APA](http://owl.english.purdue.edu/owl/resource/560/10/) recommends the use of the DOI over the URL. Many publishers also request that the DOI be included in submitted manuscripts, even if they are not included in the final published paper, as they use them for formatting and to link references with electronic databases.

In general, if a CSL/Zotero style includes the DOI, it is a good idea to leave it in. If you have a good reason to remove the DOI from the official style, it is possible to edit the CSL style. See this discussion for more details.

### Why can’t I access a proxied site when the Zotero Connector is enabled? {#why-cant-i-access-a-proxied-site-when-the-zotero-connector-is-enabled}
#### Why can’t I access a proxied site when the Zotero Connector is enabled?#

When you attempt to access a site that you’ve previously accessed through a proxy server, the Zotero Connector can automatically redirect you through your proxy.

If a site is inaccessible through your browser when the Zotero Connector is enabled but works in other browsers or when the Connector is disabled, there may be a problem with a proxy setting that the Connector has stored.

To help ensure this doesn’t happen in the future, please perform the following steps:

1.  Generate a Debug ID from the Connector for an attempt to load the page.
2.  Open the Proxies pane of the Zotero Connector preferences and look for a relevant proxy entry. Copy down the Hostname and Scheme, and then click the entry and copy down the settings from the section below.

Post the Debug ID and proxy details to a new forum thread so that developers can investigate.

You can then delete the proxy entry from the Proxies pane, which should fix the problem temporarily.

### Why do attachments have names like “PDF” or “Accepted Version” instead of their filenames in the items list? {#why-do-attachments-have-names-like-pdf-or-accepted-version-instead-of-their-file}

Attachments have two separate names: the attachment title shown in the items list and the filename of the file on disk.

Zotero automatically renames files on disk based on parent item metadata such as the title and authors. Since the parent item row in the items list already displays that metadata, Zotero doesn’t show the filename directly in the items list. Instead, it uses simpler attachment titles such as “PDF” or “Ebook” for the first file of a given type or includes additional information about the source of the file (e.g., “ScienceDirect Full Text PDF” for a file saved from ScienceDirect, or “Accepted Version” or “Submitted Version” for open-access files). These separate titles avoid cluttering the items list with redundant metadata and prevent parent items from being unnecessarily expanded when searching for titles or creators.

Subsequent files added to an item from the filesystem will still get titles named after the filename (without the file extension), since those are likely to be supplementary files and the filename may be informative.

You can view and change the title and filename by clicking on the attachment and looking in the item pane.

While we recommend the default behavior in order to avoid redundant information in the items list, if you really prefer to view filenames instead of titles, you can enable “Show attachment filenames in the items list” option in the General pane of the settings.

#### Changes in Zotero 7#

Zotero has always automatically renamed files on disk, and it has always used separate, simpler titles such as “ScienceDirect Full Text PDF” when saving attachments from the web, for the reasons explained above.

Zotero 7 changed attachment-title handling in a couple particular cases:

1.  Prior to Zotero 7, if you **manually** ran Rename File from Parent Metadata, the attachment title was changed to match the new filename. This was a bug that led many people to believe that files weren’t being automatically renamed and that it was necessary to run Rename File from Parent Metadata on every new attachment. In Zotero 7, the title is no longer changed, and titles remain as “PDF”, “ScienceDirect Full Text PDF”, or whatever they were set to originally. Files are still renamed as always, as you can see if you click on the attachment item and look in the item pane.
2.  When dragging a file from the filesystem or creating a parent item, Zotero now sets the title of the first attachment of a given type to “PDF”, “EPUB”, etc., instead of setting the title based on the filename, in order to match titles of attachments saved from the web.

People who were unnecessarily running Rename File from Parent Metadata on every attachment, predominantly adding local files rather than saving from the web, or using the ZotFile plugin (which also set the title to the filename) might be used to seeing filenames in the items list, but we’d encourage people to give the new behavior a try.

#### Updating Titles Changed Before Zotero 7#

Zotero 8 provides an option to convert attachments previously changed to match the filename to use simpler “PDF” titles instead. Simply select the attachments or their parent items and go to Tools → Management Attachments → Normalize Attachment Titles.

### Why do I see highlighted text twice in the PDF reader or in notes created from annotations? {#why-do-i-see-highlighted-text-twice-in-the-pdf-reader-or-in-notes-created-from-a}
#### Why do I see highlighted text twice in the PDF reader or in notes created from annotations?#

Some external PDF readers copy highlighted text into the annotation comment field to make it editable. Zotero parses and displays highlighted text itself and uses annotation comments as intended, for editable comments. If you then view annotations created externally within Zotero, you’ll see the highlighted text twice. Often, the version of the highlighted text in the comment field will include extra line breaks, since external PDF readers don’t always remove them automatically the way Zotero does.

If you’re using Zotero to view annotations or add them to notes, you should disable the setting in your external PDF reader that copies highlighted text into comments

In a future version, Zotero will attempt to automatically remove comments with text that duplicates the highlighted text.

### Why does Zotero store PDF annotations in its database instead of in the PDF file? {#why-does-zotero-store-pdf-annotations-in-its-database-instead-of-in-the-pdf-file}

The new PDF reader in Zotero 6 makes PDF reading and annotation a first-class part of the Zotero experience.

To enable this tight integration, Zotero stores annotations in the Zotero database, not in the PDF file. This allows for fast, conflict-free syncing, including in groups, and enables advanced functionality that wouldn’t be possible otherwise.

Zotero annotations can be exported to PDFs with embedded annotations at any time and will never be locked in Zotero.

#### Benefits of Zotero Annotations#

With annotations stored in the database, Zotero is able to quickly sync just the details of each new or updated annotation. By contrast, with standard PDF annotations, the entire PDF file needs to be transferred after every change, so if two people in a group added an annotation at the same time, or even if a PDF was just left open on one computer, it would create an unresolvable file conflict, forcing the user to choose one side or the other. This happened regularly in earlier versions of Zotero, both in personal and group libraries, and we expect PDF annotation to get far more usage as part of the app.

Storing annotations in the database also enables advanced functionality, such as being able to tag annotations and filter for them throughout the Zotero interface. We plan to add other extended features like this going forward.

There are major performance benefits as well. For syncing, as discussed above, saving annotations back to the file requires Zotero to transfer the entire file — which could be many megabytes — after every change, whereas transferring just an individual annotation is instantaneous. It’s also harder for Zotero to track changes to external files, so if you annotate something externally, there may be a delay before you can search for those annotations in Zotero or before the updated file syncs — you might need to wait for Zotero to notice the file modification or manually trigger reprocessing and syncing.

We’ll always try to support external workflows as efficiently as possible, but it will never match the seamless experience we’re able to provide when everything is done within the app.

#### Interacting with Embedded Annotations#

While Zotero saves its own annotations to its database, it’s possible to interact with annotations embedded in a PDF file in much the same way, as well as to export PDFs with embedded annotations.

Embedded annotations show up in the Zotero PDF reader, and you can add them to Zotero notes in exactly the same ways as Zotero-created annotations. External annotations are read-only by default — indicated by a lock icon — but you can transfer them into Zotero by selecting File → “Import Annotations…” from within the PDF reader, after which they’ll be fully editable. The annotations are removed from the PDF to avoid conflicts and duplicates. (Early versions of Zotero 6 included a “Store Annotations in File…” option as well, but it could result in file conflicts and lost data, and it was removed.)

You can export a copy of the PDF with annotations embedded by using File → “Export PDF…” from the library view or “Save As…” from the PDF reader. (To export the original file, drag the attachment item from the items list to your filesystem or use right-click → Show File and copy the file from there.)

When exporting metadata (e.g., BibTeX or RIS) from your library, there’s an “Include Annotations” option under “Export Files” that will embed annotations in all exported PDFs. We plan to support other ways to export annotations in future updates.

#### Using an External PDF Reader#

While we’ve tried to create a new, better PDF experience within Zotero, which requires annotations to be stored within the database, you can always choose to use a different PDF reader if you decide one works better for you. You can set the default PDF reader from the General pane of the Zotero preferences.

If you’d like to keep the Zotero PDF reader as the default but occasionally open a PDF externally, simply right-click on the PDF in Zotero and choose Show File, and then double-click the PDF in your filesystem. (Annotations of course won’t show up, and some changes, such as moving or deleting pages, may cause problems with existing annotations in Zotero, which is why we don’t expose this option directly. Rotating and deleting individual pages can be done safely from the thumbnails tab of the Zotero PDF reader sidebar.)

#### Data Portability#

Data portability is one of Zotero’s founding principles, and we’ve gone to great lengths to ensure that, if you choose to use the built-in PDF reader, your annotations will never be locked within Zotero.

If you choose to stop using Zotero in the future, you can trivially export your entire library with annotations embedded in your PDFs.

In addition to the methods detailed above for exporting annotations, annotations are stored locally in Zotero’s open SQLite database and are extractable using standard open-source tools. They are also accessible to plugins within Zotero and to external tools via the Zotero web API.

### Why doesn’t the Zotero Connector offer to save complete data from a webpage? {#why-doesnt-the-zotero-connector-offer-to-save-complete-data-from-a-webpage}
#### Why doesn’t the Zotero Connector offer to save complete data from a webpage?#

The Zotero Connector senses information on webpages through *site translators*, and the save button in the browser toolbar will feature an icon for a given type (book, newspaper article, etc.) if a site translator recognizes the page being viewed:

On all other webpages, the Zotero Connector will display a gray page icon:

You can hover over the button to see which translator, if any, Zotero will use to save the page.

Zotero’s translators should work with most library catalogs, popular websites such as Amazon and the New York Times, and many gated databases. See Zotero Translators for more examples. If a site isn’t currently supported or a translator isn’t working, you can still click the gray icon to save basic information from the page to your Zotero library, but you may need to fill in some details that Zotero couldn’t automatically detect.

If you see only a webpage icon on Amazon product pages or NY Times articles (not the home pages), see Troubleshooting Translator Issues.

If you see an item type icon on Amazon product pages and NY Times articles but see a webpage icon on another supported site, you can report it in the forums. Be sure to provide an example URL.

### Why is my browser saying the Zotero Connector needs access to my data on all websites? {#why-is-my-browser-saying-the-zotero-connector-needs-access-to-my-data-on-all-web}
#### Why is my browser saying the Zotero Connector needs access to my data on all websites?#

When you install the Zotero Connector in your browser, your browser will tell you the permissions required by the extension:

-   Chrome: “Read and change all your data on all websites”
-   Firefox: “Access your data for all websites”
-   Safari: “Can read sensitive information from webpages, including passwords, phone numbers, and credit cards. Can alter the appearance and behavior of webpages. This applies on all webpages.”

This is the standard permission an extension needs in order to interact with the content on webpages as you browse the web. The Zotero Connector needs it to detect content on the page and update the save icon with the detected item type (e.g., a book icon for “Save to Zotero (Amazon)”, a PDF icon for “Save to Zotero (PDF)”), as well as to provide advanced features such as automatic proxy redirection and automatic RIS/BibTeX import.

No data about the pages you visit is logged on your computer or sent to Zotero servers unless you click the save button, at which point the Connector will save data and files either to the Zotero app on your computer (if it’s open) or your online library on zotero.org (if you’ve given the Connector permission to do so and logged in).

If saving fails after you click the save button and “Report broken site translators to zotero.org” is enabled in the Zotero Connector preferences, the Zotero Connector will send an an anonymous error report, including the URL, browser, and version information, to Zotero servers, so that we can more quickly fix site compatibility issues. We store this information for up to one week. No additional personally identifying information (e.g., username or IP address) is stored, and reports are generally only viewed in aggregate.

For more information on the data Zotero collects, see our privacy policy.

### Zotero Connector {#zotero-connector}

Zotero Connector is a browser extension that helps you create a bibliographic library with items rich in metadata. It adds a button to your browser which allows to save items with a single click. Zotero Connector is available for Firefox, Chrome and Safari.

#### Adding Items#

After installing the Connector a Zotero button will be added to your browser. Clicking the button will save the currently open website into your Zotero library. The saving workflow has three steps:

##### Select a collection in Zotero#
##### Click the Connector button#
##### The item is saved into your library#

For more detailed instructions read the Adding Items to your Zotero Library guide.

#### Other Features#

Aside from item saving, the Connector has other features that may help your workflow.

##### Institutional Proxy Detection#

Many institutions provide a way to access electronic resources while you are off-campus by signing in to a web-based proxy system. The Connector makes this more convenient by automatically detecting your institutional proxy. Once you’ve accessed a site through the proxy, the connector will automatically redirect future requests to that site through the proxy (e.g., if you open a link to jstor.org, you’ll be automatically redirected to jstor.org.proxy.my-university.edu).

Proxy detection does not require manual configuration. You can disable or customize it from the connector preferences.

##### Bibliographic File Importing#

Many online databases offer users the option to export citation data directly to RIS or Refer format. Zotero Connector will automatically prompt you to add the references directly into your library. If you choose ‘Cancel’, you can download the file normally.

##### Citation Style Installation#

The Connector makes it easy to install citation styles from Zotero Style Repository. Clicking on a style will prompt you to install it into Zotero. If you choose ‘Cancel’, you can download the style normally.

##### Saving to Online Library#

Zotero Connector allow saving items directly to your zotero.org online library without Zotero open. The preferred way to manage and save to your library should still be by running the Zotero client, but saving to online library can be used as an alternative, if the client is not available.

When using the Connector to save a page without Zotero running, you will be prompted to authorize with your zotero.org account. After successful authorization the Connector will save items directly to your online library.

You can see your online library by clicking My Library at the top of this page. To access the saved items from Zotero, you should set up Syncing.

#### Preferences#
##### Locating the preferences page#
##### Chrome#

You can find the Connector Preferences by either right-clicking on Zotero Connector extension icon and selecting Options or by navigating to // and clicking the Options link under Zotero Connector extension.

##### Firefox#

Connector Preferences page is found by right-clicking on a page, under Zotero Connector submenu or by navigating to // and clicking the Preferences button under Zotero Connector extension.

##### Safari#

You can locate the Connector Preferences page by long-pressing on the Connector extension button and selecting Zotero Preferences or by right-clicking on a page and selecting Zotero Preferences.

##### Sections#

The preference page has the following sections:

-   General: check Zotero client status and authorize saving to zotero.org
-   Proxies (not available on Safari): set up and change institutional proxy settings
-   Advanced: options for troubleshooting, reporting errors and debugging

### Zotero Connector and Safari {#zotero-connector-and-safari}
#### Installation#

The Zotero Connector for Safari is bundled with the Zotero desktop app. (Unlike other browsers, Safari does not allow direct installation of browser extensions.) After opening Zotero for the first time, you can enable the Zotero Connector from the Extensions pane of the Safari settings (“Safari” menu → “Settings” → “Extensions”, **not** “Safari” menu → “Safari Extensions…”).

The Zotero Connector for Safari requires macOS 11 Big Sur or later.

Using an iPhone or iPad? You can save to the [Zotero iOS app](https://apps.apple.com/us/app/zotero/id1513554812) using the Share sheet in Safari and other browsers.

#### Extension not showing up? Save button missing or flickering?#

A macOS bug can cause the Zotero Connector to disappear from Safari or stop working after the Zotero app is updated.

If you find that the extension isn’t listed in the Safari settings or the toolbar button isn’t appearing or working properly, you can likely fix it using one of the options below. Your data will not be affected.

Currently, the only known way to avoid this problem altogether is to switch to a browser with a more reliable extension framework such as Firefox, Chrome, or Edge.

##### Fix 1: Delete and Reinstall Zotero#

1.  Delete the Zotero app from Applications
2.  Redownload it from the download page
3.  Start Zotero

It may also help to restart your computer in between deleting the app and restoring it, though this isn’t usually necessary.

##### Fix 2: Compress and Uncompress Zotero#

This option avoids redownloading the app.

1.  In Finder, go to the Applications folder and compress the Zotero app (right-click → Compress “Zotero”)
2.  Delete the app
3.  Double-click the ZIP file
4.  Delete Zotero.zip
5.  Start Zotero

##### Fix 3: Force Extension Reloading (advanced)#

1.  Go to Safari → Settings → Advanced
2.  Enable “Show features for web developers” at the bottom
3.  Click the new Developer tab that appears and toggle “Allow Unsigned Extensions” on and off. The Zotero Connector is signed, so it is not necessary to leave “Allow Unsigned Extensions” enabled.

#### Limitations#

Due to technical limitations of the Safari extension framework, some features available in Firefox, Chrome, and Edge aren’t available in Safari:

-   Automatic proxy redirection
-   Automatic RIS/BibTeX import
-   Automatic CSL installation

Other differences:

-   Gated PDFs may not be saved on some sites (e.g., ScienceDirect)
-   It’s not possible to right-click on the toolbar button to access secondary translators. Instead, right-click on the page itself.

### Zotero Connector Preferences {#zotero-connector-preferences}

The Zotero Connector browser extensions allow you to add items to your Zotero library with the click of a button in Firefox, Chrome, or Safari. This page describes the preferences for the Zotero Connectors.

#### Accessing the Zotero Connector Preferences#

-   **Firefox:** Right-click on the Zotero save button and choose Preferences
-   **Chrome:** Right-click on the Zotero save button and choose Options
-   **Safari:** Right-click on the page background and choose “Zotero Preferences”

#### General#

-   **Zotero Status:**
    -   Whether the Zotero Connector can connect to the Zotero desktop client. If the Connector reports that Zotero is unavailable, see Zotero Unavailable.
-   **Save to Zotero.org:**
    -   When the Zotero desktop client is closed, the Zotero Connector will save directly to the zotero.org servers. These settings let you reauthorize your broswer to save to your zotero.org account or clear your account credentials. You can also control whether PDF attachments and web page snapshots are automatically saved when importing to zotero.org.
-   **Automatic File Importing:**
    -   By default, the Zotero Connector will offer to import RIS, BibTeX, and Refer/BibIX bibliographic files when you open them in your browser. You can disable this feature or manage the sites from which data is imported here.

#### Proxies#

Many institutions require you to sign-in to a proxy system to access electronic resources while you are off-campus. The Zotero Connector can make this more convenient. When it detects that you are using an institutional proxy to access a particular site, it will ask if you want to remember it in the future. If you agree, Zotero will automatically use the proxy for matching URLs in the future. You should be routed through the proxy login site if you’re not already logged in, then you can access the database as you normally would.

Zotero users can use the proxies feature without ever looking at this preference tab. By default, Zotero will prompt you to store the proxy and then route you through the proxy automatically and without further input.

Zotero proxy redirection is not available in Safari.

The Proxies preferences allow you to adjust the following options:

-   **Enable proxy redirection**
    -   Zotero’s proxy redirection is enabled by default. Uncheck this option to disable proxy redirection. You can do this temporarily and your proxy settings will remain saved. *Do not* use this option if you no longer have access to the saved proxies. In that case, delete those settings by selecting them in the “Configured Proxies” box and pressing the minus (-) button below it.
-   **Show a notification when redirecting through a proxy**
    -   By default, Zotero will show a temporary banner at the top of your browser when it redirects through a saved proxy. Uncheck this box to disable this notification.
-   **Automatically detect new proxies**
    -   By default, Zotero will automatically detect when you visit a page through an institutional proxy and offer to remember the proxy the next time you visit the website. Uncheck this box to prevent Zotero from prompting you to store proxies it detects.
-   **Disable proxy redirection when domain name contains**
    -   Typically you won’t need to use a proxy when you are connected to the internet through your institution’s network. This option automatically disables Zotero’s proxy re-direction when the domain of your internet provider contains the given string. In the United States, “.edu” (the default setting) will usually work. In other countries you will have to find out your institution’s domain name.
    -   This option is disabled by default. This option is only available when the Zotero desktop client is open.

##### Configured Proxies#

When Zotero automatically detects and saves institutional proxies, they will be stored here. You can remove stored proxies by clicking the minus (-) button below the list. If you are having issues with a proxy, try to remove it from the list and re-add it by visiting the site and letting Zotero automatically detect the proxy settings again.

You can manually add proxies by by clicking on the plus (+) button. From there, you can specify the URL of the database being accessed under hostname and the URL scheme of the proxy. You can add/remove additional URLs to redirect through a single proxy by clicking on the plus (+) and minus (-) buttons below the Hostnames list. You can also enable/disable automatic association of new hostname URLs with a proxy server.

Some proxy servers require hyphens in proxied hostname URLs to be converted to dots. Check the box for this option if this is the case for your proxy server.

If you are having trouble accessing a site due to Zotero proxy redirection functionality, see Proxy Troubleshooting.

#### Advanced#

These preferences are used for reporting errors and troubleshooting information to the Zotero developers.

-   **Report Errors:** If you are having a problem using the Zotero Connector, use this button to submit an Error Report ID to Zotero, then post to the Zotero forums. See Reporting Problems for instructions on how to submit helpful error reports.
-   **Debug Output Logging:** To help diagnose a problem, the Zotero developers may ask you to submit a Debug Log ID. This is different from an Error Report ID above. To submit a debug log, check “Enable Logging”, then complete the sequence of steps neeeded to produce your error. Then, click “Submit Debug Report” and post the Debug ID number to the Zotero forums. Try to avoid performing unrelated actions when making a debug log.
-   **Translators:** Zotero will automatically check for and install updated translators. You can manually check for updates here. By default, the Zotero Connector will report broken site translators to zotero.org. This helps Zotero to keep the Zotero import process working smoothly on sites across the Web.
-   **Advanced Configuration:** The Config Editor options are not useful for general troubleshooting and should only be used if instructed by the Zotero developers.

### Zotero Connector: “Is Zotero Running?” {#zotero-connector-is-zotero-running}

When you click the Save to Zotero button in your browser or try to use Zotero with Google Docs, you may receive the following message:

*“The Zotero Connector was unable to communicate with the Zotero desktop application.”*

The Zotero Connector needs to connect to the Zotero desktop app in order to save data or insert citations into Google Docs. (It can also save pages directly to zotero.org, but saving to the Zotero application provides the best experience.)

First, make sure that Zotero is installed and open on your computer. If you don’t yet have the Zotero application, you can install it from the downloads page.

Next, restart your browser and try again.

After restarting, if Zotero is open but the Zotero Connector still reports that Zotero is unavailable, something on your computer is preventing the Connector from talking to Zotero. You can determine whether the problem is within the browser or a system-wide problem by loading the URL 0 in one or more browsers while Zotero is open. If Zotero is running, it should display “Zotero is running” or “Zotero Connector Server is Available”.

-   If you can load that URL but the Zotero Connector still shows Zotero as unavailable, try these steps until the problem is resolved:
    1.  Make sure you’ve given the Zotero Connector permission to access all websites. In Chrome or Edge, right-click on the Save to Zotero button, select “This Can Read and Change Site Data”, and make sure it’s set to “On All Sites”. In Safari, go to the Websites tab of the Safari settings, select Zotero Connector under Extensions, and make sure “For other websites” is set to “Allow”. The Zotero privacy policy explains why this is necessary.
    2.  Uninstall and reinstall the Zotero Connector.
    3.  Temporarily disable any browser extensions that may block network requests, such as AdBlock, uBlock Origin, NoScript, EFF Privacy Badger, or Request Policy. If the Connector stops saying that Zotero is offline, reenable each extension one at a time and, if the problem recurs, whitelist 0 port 1 in the extension’s settings.
    4.  Temporarily disable any other installed extension.
    5.  Try in another browser. If the problem occurs there as well, security software on your computer is likely blocking extension requests across multiple browsers.
    6.  If the other browser works, create a new profile in your original browser. If this fixes the problem, something is wrong with your original profile, and you’ll need to either identify the problem or transfer your bookmarks, history and other data to the new profile. If the problem occurs in a new profile, security software on your computer may be interfering with extension requests only in this particular browser.
-   If you can’t load the URL in one browser but you can in another browser, try these steps in the original browser until the URL works:
    1.  If your computer is connecting via a proxy server, ensure that the host 0 (an alias for your computer itself) is excluded from proxying in either the browser or system proxy settings.
    2.  Temporarily disable any browser extensions that may block network requests, such as AdBlock, uBlock Origin, NoScript, EFF Privacy Badger, or Request Policy. If the URL starts working, reenable each extension one at a time and, if the problem recurs, whitelist 0 port 1 in the extension’s settings.
    3.  Temporarily disable any other installed extension.
-   If you can’t load the URL in any browser, try these steps until the URL works:
    1.  If your computer is connecting via a proxy server, ensure that the host 0 (an alias for your computer itself) is excluded from proxying.
    2.  Restart your computer.
    3.  Temporarily disable any security software running on your system. If the URL starts working, reenable each piece of software one at a time and, if the problem recurs, whitelist 0 port 1 in the software’s settings.
    4.  Restart Zotero with debug logging enabled via Help → Debug Output Logging → “Restart with Logging Enabled…”, and then go to Help → Debug Output Logging → View Output. Copy the output to a text file and search for “HTTP server”. You should see one of three messages:
        -   0 — Zotero is listening successfully and something else is blocking the connection.
        -   0 — Zotero wasn’t able to listen on port 23119, possibly because some other program was already doing so. This should be preceded by an error that may provide more information.
        -   0 — Zotero couldn’t detect a network connection at all.
    5.  It may help to try in a new OS account. (You don’t need to set up syncing in Zotero — an empty library is fine for testing.) If the URL starts working, you’ll need to figure out what in your original account is blocking the connection, based on the message you see in Zotero debug output. If it doesn’t work in a new account either, some system-level software on your computer is blocking the connection, and you’ll need to troubleshoot this on your own system.

### Zotero PDF Reader and Note Editor {#zotero-pdf-reader-and-note-editor}
#### Creating Annotations#
##### Highlights and Underlines#

The reader supports two highlighting/underlining modes. You can use whichever mode works best for your workflow.

In the default, **unlocked mode**, with nothing selected in the annotation toolbar, whenever you select text in the document, a popup will appear with a selection of colors, and you can click a color to create a highlight or underline on the selected text.

To quickly make many highlights or underlines, you can turn on **locked mode**: click the highlight or underline tool in the toolbar, and then choose a color using the color-picker button. Then, every time you select text in the document, a highlight or underline will be automatically created in the color you’ve picked.

#### Adding Annotations to Notes#

You can add annotations to notes in various ways. Annotations added to notes will automatically include links back to the PDF page as well as citations that you can later add to a Word, LibreOffice, or Google Docs document with one of the word processor plugins.

##### In the PDF reader#

You can easily add annotations to notes right from the PDF reader.

First, use the Notes button in the top-right corner to open the Notes pane, where you can create a new note or open an existing note.

To create a new note from all annotations in the current PDF, click one of the “+” buttons and select Add Item Note from Annotations or Add Standalone Note from Annotations.

If you already have a note open in the Notes pane, you can drag individual annotations from the PDF or from Annotations tab in the left-hand sidebar as you type your note. Alternatively, you can select one or more annotations in PDF or in the the Annotations tab of the left-hand sidebar, right-click one of the annotations, and select Add to Note.

You can also drag annotations from the PDF reader to a note that’s opened in a separate window.

If you’re sure you won’t use a quote more than once, it’s also possible to add quotes to Zotero notes without creating an annotation first. Simply select text in the PDF and drag it to an open Zotero note.

##### In the items list#

You can create a child note from all annotations in a PDF by right-clicking on the parent item in the items list and choosing Add Note from Annotations.

You can also create a standalone note with annotations from multiple items by selecting the parent items, right-clicking, and choosing Create Note from Annotations.

#### Customizing Annotations in Notes#

You can customize how annotations are added to notes by using note templates.

#### Working with Annotations in Notes#
##### Viewing an Annotation in Context#

When you add an annotation to a note, it will add both an annotation and a citation by default.

To view the annotation in context, click the annotation and click “Show on Page” in the popup that appears. This will open the original PDF in either the built-in PDF reader or an external PDF reader if you’ve configured one in the General pane of the preferences. When possible, Zotero will open the PDF to the page where the annotation was made — this works for the built-in PDF reader and most, but not all, popular external PDF readers (e.g., Acrobat on Windows, Preview on macOS, evince on Linux). If Zotero isn’t opening to appropriate page PDF in your external PDF reader, let us know in the Zotero Forums.

To view the associated item in your library, click the citation and select “Show Item”.

##### Displaying Annotation Colors#

To show annotation colors in your note, click the “…” button in the top-right corner of the note editor and select “Show Annotation Colors”. You can remove colors later with “Hide Annotation Colors”.

##### Hiding or Showing Citations#

To hide or show all citations in a note, click the “…” button in the top-right corner of the note editor and select “Hide Annotation Citations” or “Show Annotation Citations”.

You can hide an individual citation by clicking on it in the editor and selecting “Hide Citation” from the popup, or simply by deleting the citation completely. You can add the citation back at any time by clicking on the highlighted text and selecting “Add Citation” from the popup or by using “Show Annotation Citations” in the “…” menu.

#### Using an External PDF Reader#

If you’d prefer to open PDFs in an external PDF reader, you can choose one from the General pane of the Zotero preferences.

To open a single PDF in an external reader, right-click on the item and choose Show File, and then open the PDF from your OS file manager. Some changes you make in external PDF readers will cause problems in the built-in reader — e.g., adding, deleting, reordering, or rotating pages will cause annotations to appear on the wrong page or in the wrong position. (Pages can be deleted or rotated from the Annotations tab in the Zotero reader sidebar.)

Note that annotations created in the built-in PDF reader are stored in the Zotero database, so they won’t be visible in external PDF readers unless you export a PDF with embedded annotations. See Annotations in Database for more info.

### Zotero Translators {#zotero-translators}

Zotero automatically detects journal articles, library records, news items, and other objects you might like to save to your Zotero library. Zotero uses so-called “translators” to detect and import data from websites. There are currently more than 600 different translators, facilitating data import from countless sites.

You can see the translator Zotero is using on a given page by hovering over the Zotero save button in your browser.

#### Types of Sites Supported#
#### Library Catalogs#

Zotero imports records from many library cataloging systems, providing seamless import from hundreds of academic and non-academic libraries. Supported library catalogue systems include: Aleph, Amicus, BiblioCommons, Dynix, Encore, Mango, InnoPAC, Primo, SirsiDynix, TLC/YouSeeMore, Voyager, and WorldCat.

#### Databases#

Zotero imports data, and in many cases full-text PDFs, from the most popular electronic databases, including EBSCO, IEEEXplore, JSTOR, Google Scholar, ProQuest, PubMed, and many more. It also works with most major journal publishers, including Cambridge University Press, Oxford University Press, Project MUSE, ScienceDirect (Elsevier), SpringerLink, Taylor and Francis, and many more.

#### Individual Site Translators#

Zotero has dedicated translators for hundreds of websites, ranging from Amazon (in various countries), The New York Times, and The Economist, to Mainichi Daily News (Japan), Kommersant (Russia), Spiegel Online (Germany), and many more from around the world.

#### Metadata Import#

Zotero detects and imports metadata embedded by an increasing number of websites and databases in open formats such as COinS, Embedded RDF, Google/HighWire meta tags, and unAPI.

#### Full List of Translators#

You can see a list of the 600+ Zotero translators, along with their code, at the [Zotero Translator GitHub repository](https://github.com/zotero/translators/).

#### Translator Troubleshooting#

If Zotero fails to import high-quality data from a site that you think should be supported, first check whether you’re experiencing a general problem by trying to save an article from Wikipedia or a book from Amazon. If that doesn’t work either, see Troubleshooting Translator Issues. You’re having trouble with a specific site, report it in the Zotero Forums along with the exact URL you’re trying.

## Getting started {#getting-started}
### The Basics {#the-basics}

**Zotero \[zoh-TAIR-oh\] is a free, easy-to-use tool to help you collect, organize, cite, and share your research sources.**

Read on for an overview of Zotero’s features and capabilities.

#### How do I install Zotero?#

See the installation instructions.

#### How do I open Zotero?#

Zotero can be opened from your operating system’s dock or file manager like any other program.

#### What does Zotero do?#

Zotero is, at the most basic level, a reference manager. It is designed to store, manage, and cite bibliographic references, such as books and articles. In Zotero, each of these references constitutes an item. More broadly, Zotero is a powerful tool for collecting and organizing research information and sources.

#### What kind of items are there?#

Every item contains different metadata, depending on what type it is. Items can be everything from books, articles, and reports to web pages, artwork, films, letters, manuscripts, sound recordings, bills, cases, or statutes, among many others.

#### What can I do with items?#

Items appear in Zotero’s center pane. The metadata for that item is shown in the right pane. This includes titles, creators, publishers, dates, page numbers, and any other data needed to cite the item.

#### Organize#
##### Collections#

The left pane includes My Library, which contains all the items in your library. Right-click on My Library or click on the New Collection button () above the left pane to create a new collection, a folder into which items relating to a specific project or topic can be placed. Think of collections like playlists in a music player: items in collections are aliases (or “links”) to a single copy of the item in your library. The same item can belong to many collections at one time.

##### Tags#

Items can be assigned tags. Tags are named by the user. An item can be assigned as many tags as is needed. Tags are added or removed with the tag selector at the bottom of the left pane or through the Tags tab of any item in the right-hand pane. Up to 6 tags can be assigned **colors**. Colored tags are readily visible in the item list and can be quickly added or removed using the number keys on your keyboard.

##### Searches#

Quick searches show items whose metadata, tags, or fulltext content match the search terms and are performed from the Zotero toolbar. Clicking the spyglass icon to the left of the search box opens the Advanced Search window, allowing for more complex or narrow searches.

##### Saved Searches#

Advanced searches can be saved in the left pane They are similar to collections, but will update with new matching items automatically.

#### Collect#
##### Attachments#

Items can have notes, files, and links attached to them. These attachments appear in the middle pane underneath their parent item. Attachments can be shown or hidden by clicking the arrow next to their parent item.

##### Notes#

Rich-text notes can be attached to any item through the Notes tab in the right-hand pane. They can be edited in the right-hand pane or in their own window. Click the New Note button () in the toolbar to create a note without attaching it to an item.

##### Files#

Any type of file can be attached to an item. Attach files with the Add Attachment (paperclip) button in the Zotero toolbar, by right-clicking on an existing item, or by drag-and-dropping. Files do not need to be attached to existing items. They can simply be added to your library. Files can also be downloaded automatically when you import items using the Zotero Connector in your browser.

##### Links & Snapshots#

Web pages can be attached to any item as a link or a snapshot. A link simply opens the website online. Zotero can also save a snapshot of a web page. A snapshot is a locally stored copy of a web page in the same state as it was when it was saved. Snapshots are available without an internet connection.

##### Capturing Items#

With the Zotero Connector for Chrome, Firefox, or Safari, it’s simple to create new items from information available on the internet. With the click of a button, Zotero can automatically create an item of the appropriate type and populate the metadata fields, download a full-text PDF if available, and attach useful links (e.g., to the PubMed entry) or Supplemental Data files.

##### Single or Multiple Captures#

If the save icon is a book, article, image, or other single item, clicking on it will add the item to the current collection in Zotero. If the save icon is a folder, the webpage contains multiple items. Clicking it will open a dialog box from which items can be selected and saved to Zotero.

##### Translators#

Zotero uses bits of code called translators to recognize information on webpages. There are generic translators which work with many sites and translators written for individual sites. If a site you’re using does not have a translator, feel free to request one on the Zotero Forums.

##### Saving a Web Page#

If the Zotero Connector does not recognize data on the page, you can still click the save button in the browser toolbar to save the page as a Web Page item with an attached snapshot. While this will save basic metadata (title, URL, access date), you may need to fill in additional metadata from the page by hand.

##### Add Item by Identifier#

Zotero can add items automatically using their an ISBN number, Digital Object Identifier (DOI), or PubMed ID. This is done by clicking the Add Item by Identifier button () in the Zotero toolbar, typing in the ID number, and clicking OK. You can even paste or enter (press Shift+Enter for a larger box) a list of such identifiers at once.

##### Feeds#

Subscribe to RSS feeds from your favorite journals or websites to keep up to date with the latest research. Go to the article web page or save items to your library with the click of a button.

##### Manually Adding Items#

Items can be added manually by clicking the green New Item () button in the Zotero toolbar and selecting the appropriate item type. Metadata can then be added by hand in the right-hand pane. While you should generally not add items manually, it can be useful for adding primary documents that aren’t available online.

#### Cite#
##### Citing Items#

Zotero uses Citation Style Language (CSL) to properly format citations in many different bibliographic styles. Zotero supports all the major styles (Chicago, MLA, APA, Vancouver, etc.) as well as the specific styles for over 8,000 journals and publishers.

##### Word Processor Integration#

Zotero’s Word, LibreOffice, and Google Docs plugins allow users to insert citations directly from their word processing software. This makes citing multiple pages or sources or otherwise customizing citations a breeze. In-text citations, footnotes and endnotes are all supported. With community-developed plugins, Zotero can also be used with LaTeX, Scrivener, and numerous other writing programs.

##### Automatic Bibliographies#

Using the word processor plugins makes it possible to automatically generate a bibliography from the items cited and to switch citation styles for the entire document with the click of a button.

##### Manual Bibliographies#

Zotero can also insert citations and bibliographies into any text field or program. Simply drag-and-drop items, use Quick Copy to send citations to the clipboard, or export them directly to a file.

#### Collaborate#
##### Syncing#

Use Zotero on multiple computers with Zotero syncing. Library items and notes are synced through the Zotero servers (unlimited storage), while attachment syncing can use the Zotero servers or your own WebDAV service to sync files such as PDFs, images, or audio/video.

##### Zotero Servers#

Items synced to the Zotero servers can be accessed online through your zotero.org account. Share your library with others or create a custom C.V. from selected items.

Make copies of your research readily available on zotero.org for readers, the public, and other researchers using My Publications.

##### Groups#

Zotero users can create collaborative or interest groups. Shared group libraries make it possible to collaboratively manage research sources and materials, both online and through the Zotero client. Zotero.org can be the hub of all your project group’s research, communication and organization.

### Can I still use Zotero if I can’t install programs on my computer? {#can-i-still-use-zotero-if-i-cant-install-programs-on-my-computer}

On Windows, you can cancel the admin prompt when you install Zotero, and it will perform a local installation that you will have permission to update as a local user.

We also provide a ZIP version of Zotero for Windows that you can extract and run from any folder. The ZIP version doesn’t automatically register Zotero as a handler for various file types (e.g., RIS files), so we recommend using the installer if possible.

On macOS and Linux, you can run the application from any folder (e.g., an Applications folder inside your home folder on macOS instead of /Applications).

### Does Zotero offer installable packages of Zotero for specific Linux distributions? {#does-zotero-offer-installable-packages-of-zotero-for-specific-linux-distribution}
#### Does Zotero offer installable packages of Zotero for specific Linux distributions?#

No. Zotero is distributed as a tarball that can be extracted and run manually or packaged by others according to distribution-specific standards. Some that have been packaged are

1.  debian/ubuntu/some chromebooks
    -   0 (replaces the retorquere PPA)
        -   Includes packages for Zotero and Juris-M
2.  Snap packages, usable on multiple distributions: 0
3.  Flatpack packages, usable on multiple distributions: 0
4.  debian
    -   ~~[sid/unstable](http://packages.debian.org/sid/zotero-standalone)~~ no longer available
5.  ubuntu
    -   ~~There is a [package](https://launchpad.net/ubuntu/+source/zotero-standalone-build), but it may be defective (see discussion here.~~ no longer available
    -   ~~There is also a PPA: 0 no longer maintained; last version 5.0.60
    -   ~~And another PPA: 0 ; includes package for both Zotero and Juris-M~~ migrated (see above)

If your distro is not listed above, check if a package is already available. If so, please list it here; if not, please check if a packaging request has already been made (and consider making one if not), follow your distribution’s guidelines for requesting a package.

### How Do I Install Zotero on a Chromebook? {#how-do-i-install-zotero-on-a-chromebook}
#### Step 1: Set up Linux on Chrome OS#

Following the steps from Google to [Set up Linux on your Chromebook](https://support.google.com/chromebook/answer/9145439?hl=en).

#### Step 2: Open Terminal#

1.  After Linux is installed, you will notice a new app in your overflow menu (where all your app icons live) called Terminal.
2.  Wait for the Terminal app to open. This might take a few minutes the first time.

#### Step 3: Install Zotero#

Enter these commands in Terminal to install a packaged version of Zotero [maintained by a community member](https://github.com/retorquere/zotero-deb):

    curl -sL https://raw.githubusercontent.com/retorquere/zotero-deb/master/install.sh | sudo bash
    sudo apt update
    sudo apt install zotero

(If you prefer, you can install the official tarball, but you will have to perform some setup steps manually.)

Once these finish, you can close the Terminal and go back to your overflow (apps) menus. You will now see an icon for Zotero, and clicking on it will open the app. You can then pin the app to your Chrome Launcher.

#### Step 4: Set up the Zotero Connector#

*This step may no longer be required on current versions of ChromeOS. If you can save to the Zotero app from the Zotero Connector, you can skip this step.*

To use the Zotero Connector, which allows you to save from Chrome to Zotero and use Zotero in Google Docs, you may need to install a port-forwarding app such as Connection Forwarder.

To set up forwarding using Connection Forwarder, close Zotero and any other Linux apps, and then set a forwarding rule as follows:

-   Protocol: TCP
-   Source: 127.0.0.1 (Localhost) port 23119 (Source Port, i.e., the connection port from Chrome)
-   Destination: 127.0.0.1 (Localhost) port 8080 (Destination Port, i.e., the target port in Linux)

Within the Zotero app, go to Edit → Preferences → Advanced → Config Editor, set 0 to 1, and then restart Zotero.

Once you have done all these things, you should be good to go.

### How do I uninstall the Zotero word processor plugins? {#how-do-i-uninstall-the-zotero-word-processor-plugins}
#### Word#

You can uninstall the Word plugin by deleting Zotero.dotm from your Word Startup folder and restarting Word.

#### LibreOffice#

You can uninstall the LibreOffice plugin by going to Tools → “Extension Manager…”, selecting Zotero LibreOffice Integration, and clicking Remove.

#### Google Docs#

The Google Docs plugin is part of the Zotero Connector. If you’re no longer using Zotero, you can uninstall the Zotero Connector from your browser’s Extensions pane. If you want to continue using the Zotero Connector but don’t want the Zotero menu or toolbar button in Google Docs, you can disable Google Docs integration from the Advanced pane of the Zotero Connector preferences.

### How do I uninstall Zotero? {#how-do-i-uninstall-zotero}

You can uninstall Zotero like any program on your computer. The Zotero data directory is not removed when you uninstall Zotero.

You can uninstall the Zotero Connector from your browser’s Extensions pane.

To uninstall the Zotero word processor plugins, see these instructions.

### Installation Instructions {#installation-instructions}
#### Installation Instructions#
##### Where do I download Zotero?#

You can download Zotero on the Zotero download page. Be sure to also install the Zotero Connector for your browser.

##### How do I install Zotero?#
##### Mac#

Open the .dmg you downloaded and drag Zotero to the Applications folder. You can then run Zotero from Spotlight, Launchpad, or the Applications folder and add it to your Dock like any other program.

After installing Zotero, you can eject and delete the .dmg file.

##### Windows#

Run the setup program you downloaded.

##### Linux#
##### Official Tarball#

Download the tarball, extract the contents, and run 0 from that directory to start Zotero.

For Ubuntu and other distros that support .desktop files, the tarball includes a .desktop file that can be used to add Zotero to the launcher:

1.  Move the extracted directory to a location of your choice (e.g., 0).
2.  Run the 0 script from a terminal to update the .desktop file for that location. .desktop files require absolute paths for icons, so 1 replaces the icon path with the current location of the icon based on where you’ve placed the directory.
3.  Symlink 0 into 1 (e.g., 2)

Zotero should then appear either in your launcher or in the applications list when you click the grid icon (“Show Applications”), from which you can drag it to the launcher.

You may need to re-run 0 after certain Zotero updates. If something isn’t working, it may help to remove the current symlink (1), wait a few seconds for Zotero to disappear from the launcher, and recreate it.

##### Debian/Ubuntu-based Distros#

A longtime community member maintains [zotero-deb](https://github.com/retorquere/zotero-deb), a lightweight wrapper for the official tarball that uses 0 for installation and updates. If you’re not comfortable installing tarballs or are having trouble following the above instructions, this is what we recommend using.

##### Other Packages#

Unofficial packages are also available for various distros, but note that such packages are built by third parties, and we can only provide support for the official tarball and zotero-deb. In particular, third-party packages may be sandboxed, breaking various functionality.

##### Chromebook#

To set up Zotero on a Chromebook, see Installing on a Chromebook.

##### How do I upgrade to a new version?#

Zotero should update itself automatically by default, or you can go to the Help menu and select “Check for Updates…” to check for updates manually. You can also always manually install a new version of Zotero over your existing version without losing any data.

##### Troubleshooting#

If you’re running older software, check the system requirements to be sure Zotero is compatible with your system.

If something still isn’t working, let us know in the Zotero Forums.

### Installing the Zotero Word Processor Plugins {#installing-the-zotero-word-processor-plugins}

The word processor plugins are bundled with Zotero and should be installed automatically for each supported word processor on your computer when you first start Zotero.

You can reinstall the plugins later from the Cite → Word Processor Plugins section of the Zotero settings. If you’re having trouble, see the troubleshooting instructions.

If you previously installed the Firefox versions of the word processor plugins into Zotero 5.0 or Zotero Standalone 4.0, you should uninstall them from Tools -> Add-ons.

### Manually Installing the Zotero Word Processor Plugin {#manually-installing-the-zotero-word-processor-plugin}

The Zotero Word plugins will be installed automatically into Word for most users. If you don’t see a Zotero toolbar in Word, you should attempt to reinstall the plugin from the Cite → Word Processors pane of the Zotero preferences. If you receive an error or still don’t see the plugin after trying to reinstall from the preferences, you can try the manual installation instructions below.

Note that, if you rely on manual installation, you may run into problems later due to the plugin in Word becoming outdated, so it’s better to figure out why automatic installation isn’t working (e.g., security software blocking the installation or an incorrect Word Startup folder location) and fix the underlying problem.

#### Word for Windows#

1.  Open the Zotero installation folder (0).
2.  In the installation folder, open 0, where you can find a copy of the Zotero.dotm file.
    -   If the folder is empty, the file was somehow deleted — possibly by security software — and you should reinstall Zotero.
    -   If the folder is empty immediately after reinstalling Zotero, you can download [Zotero.dotm](https://github.com/zotero/zotero-word-for-windows-integration/raw/main/install/Zotero.dotm), but your security software may delete the downloaded file as well, and you’ll need to configure it not to do so.
    -   If you see two “Zotero” files without file extensions, your computer is set not to display file extensions, and you can determine which one is Zotero.dotm by right-clicking on each file and selecting Properties. One will say “Microsoft Word 97-2003 Template (.dot)” and one will say “Microsoft Word Template (.dotm)”.
3.  Find your Word startup folder and copy the path to the clipboard:
    1.  In the Word ribbon, click the File tab, click Options, and then click Advanced.
    2.  Under General, click File Locations. The current Startup folder should be listed.
        -   In most cases, the Startup folder path should be the default location of 0. The path should not include “Zotero” in any way, and if it does you previously configured it incorrectly. If that’s the case, you should reset the path to the default location.
    3.  Select the Startup folder path and click Modify, click in the whitespace to the right of the path in the location bar at the top of the window, copy the complete path to the clipboard with Ctrl-C, and then **click Cancel** to close the dialog without making changes.
4.  Open a new File Explorer window and paste the Startup folder path into the address bar. You should now have two folders open: the “install” folder containing Zotero.dotm and the Word startup folder.
5.  Copy the Zotero.dotm file from “install” to your Word Startup folder. (Be sure to copy the file rather than moving it. If dragging, hold down Ctrl.)
6.  Restart Word to begin using the plugin.

#### Word for Mac#

1.  In Finder, press Cmd-Shift-G and navigate to 0, where you can find a copy of the Zotero.dotm file. If the folder is empty, the file was somehow deleted — possibly by security software — and you should reinstall Zotero.
2.  Find your Word startup folder by following the instructions below. You should now have two folders open: the Word startup folder and the “install” folder containing Zotero.dotm.
3.  Copy the Zotero.dotm file to your Word Startup folder. (Be sure to copy the file rather than moving it.)
4.  Start (or restart) Microsoft Word to begin using the plugin.

#### LibreOffice#

1.  Navigate to the Zotero application files:
    -   Mac: In Finder, press Cmd-Shift-G and paste in or 0
    -   Windows: Open the folder 0
    -   Linux: Go to the directory where Zotero is installed and open 0
2.  Double-click the Zotero\_OpenOffice\_Integration.oxt file to install it. Alternatively, go to Tools → Extension Manager in LibreOffice, click Add, and select the .oxt from the above folder.

If you get an error, there’s a problem with your LibreOffice installation, and you should follow the troubleshooting steps.

#### Locating your Word Startup folder#

Note: On non-English systems or in certain custom setups, these locations may be different.

##### Word 2007 or later for Windows#

The default location of the Startup folder is 0.

If changes you make to the Startup folder aren’t taking effect, you can confirm that Word isn’t set to a different location. In the Word ribbon, click the File tab, click Options, and click Advanced. Under General, click File Locations. The Startup folder should be listed there. Select it and click Modify. In the window that opens, click the whitespace to the right of the path in the location bar at the top and copy the complete path to the clipboard by pressing Ctrl-C. **Click Cancel** to close the dialog without making changes. You can then open a new File Explorer dialog and paste the path into the address bar to open the Startup folder.

Note that the path should not include “Zotero” in any way, and if it does you previously configured it incorrectly. If that’s the case, you should reset the path to the default location.

##### Word 2016 and 2019 for Mac#

The default location of the Startup folder is 0. (1 refers to the Library folder within your home directory.) You can open it from the Finder by pressing Cmd-Shift-G and copying in the path. Alternatively, to navigate to it in Finder, hold down Option on your keyboard, click the Go menu, and select Library (which is hidden by default), and then follow the rest of the path.

If changes you make to the Startup folder aren’t taking effect, you can confirm that Word isn’t set to a different location. In Word, open the “Word” menu in the top-left of the screen and select “Preferences”. Click on “File Locations” under “Personal Settings” and click on “Startup” at the bottom of the list.

Generally, no location should be listed, causing Word to use the default location. If another location is listed (e.g., 0, from an earlier version of Word), clearing the setting and letting Word use the default location may fix installation problems and allow Zotero to install the plugin automatically going forward.

Note that the path should not include “Zotero” in any way, and if it does you previously configured it incorrectly. If that’s the case, you should reset the path so that it is blank and the default location is used.

### Troubleshooting Errors with Word Processor Plugin Installation {#troubleshooting-errors-with-word-processor-plugin-installation}

Follow the steps below for Word for Windows, Word for Mac, or LibreOffice if you receive an error trying to install the word processor plugin.

For general troubleshooting see Word Processor Plugin Troubleshooting.

#### Word for Windows#

Zotero will automatically try to install the Word plugin and keep it up to date. You can force automatic installation at any time by going to the Zotero settings → Cite → Word Processors and pressing “(Re)install Word Add-in”. If the automatic installation procedure fails, follow the steps below.

**Try to install the plugin via the Zotero settings after each step.**

1.  Close Word before attempting to install the plugin. The plugin will fail to install if Word is open.
2.  Temporarily disable any proprietary security software.
3.  Reset your Word Startup folder. In Word go to File → Options → Advanced. At the bottom of the dialog under the “General” section, press “File Locations…”. In the dialog under “Startup”, the “Location” field should either be empty or end with 0. If it does not, select the Startup entry and click “Modify”. In the file dialog, click in the address bar, paste 1, press Enter/Return, and then click OK to confirm the new location.
4.  Make sure your user account has permissions to write to the Word Startup folder.
    1.  Open the Word Startup folder in Explorer 0
        -   Windows 11: press the three dots (…) menu button and select Properties.
        -   Windows 10: Select the Home tab and press Properties.
    2.  Open the Security tab.
    3.  In the top list, click on your Windows Account username, and confirm in the bottom list that for the entry “Full control” there is a checkmark under Allow.
    4.  Click the Advanced button and confirm that under “Owner” your username is listed.
5.  If you are working on a computer provided by your institution, contact your IT department.
6.  Install the plugin manually.

If the plugin doesn’t appear in Word after a manual installation, see Zotero toolbar doesn’t appear.

Note that you’ll need to repeat a manual installation every time the plugin is updated, so it’s much better to fix automatic installation. To troubleshoot the automatic installation, please create a new thread in the Zotero Forums so we can try to help. Be sure to include a Report ID from Zotero, your operating system and Word versions (open a Word document, choose File from the top-left corner, click Account in the left navigation bar, and see the section under About Word, on the right side of the window), and the steps you’ve taken to try to fix the problem.

#### Word for Mac#

Zotero will automatically try to install the Word plugin and keep it up to date. You can force automatic installation at any time by going to the Zotero menu → Settings → Cite → Word Processors and pressing “(Re)install Word Add-in”. If the automatic installation procedure fails, follow the steps below.

**Try to install the plugin via the Zotero settings after each step.**

1.  Close Word before attempting to install the plugin.
2.  If you’re running macOS 15 Sequoia, when Zotero attempts to install the Word plugin, you’ll be prompted to allow Zotero to “access data from other apps”. You must click **“Allow”** to give Zotero permission to copy the plugin to the Word Startup folder. If you clicked “Don’t Allow”, restart Zotero and reinstall the plugin from the Cite pane of the Zotero settings, and choose “Allow” when prompted. Zotero uses this permission solely to install the plugin, and the permission lasts only until Zotero is next closed. You won’t be prompted again until the next time Zotero needs to update the plugin.
3.  Temporarily disable any proprietary security software.
4.  Reset your Word Startup folder. In the Word menu, select “Settings…”. Under “Personal Settings”, select “File Locations”. Select the entry for “Start-up” and click Reset.
5.  Make sure your user account has permissions to write to the Word Startup folder. Open Finder and select Go -> “Go to Folder…” from the menu bar. Paste 0 into the dialog and press Return. In the menu bar, select File → Get Info. At the bottom of the dialog, expand the “Sharing & Permissions” section and make sure that your user has the “Read & Write” privilege on the folder.
6.  If you are working on a Mac provided by your institution, contact your IT department.
7.  Install the plugin manually.

If the plugin doesn’t appear in Word after a manual installation, see Zotero toolbar doesn’t appear.

Note that you’ll need to repeat a manual installation every time the plugin is updated, so it’s much better to fix automatic installation. To troubleshoot the automatic installation, please create a new thread in the Zotero Forums so we can try to help. Be sure to include a Report ID from Zotero, your operating system and Word versions (in Word, go to the Word menu → About Microsoft Word to find the full version number), and the steps you’ve taken to try to fix the problem.

#### LibreOffice#

Zotero will automatically try to install the LibreOffice plugin and keep it up to date. You can force automatic installation at any time by going to the Zotero settings → Cite → Word Processors and pressing “(Re)install LibreOffice Add-in”. If the automatic installation procedure fails, follow the steps below.

**Try to install the plugin via the Zotero settings after each step.**

1.  Check that LibreOffice is up to date.
2.  Open the LibreOffice preferences by going to Tools → Options (Windows/Linux) or LibreOffice → Settings… (Mac). In the dialog, click LibreOffice → Advanced. Ensure that “Use a Java runtime environment” is checked and that a JRE is selected in the list below.
    -   If no JRE appears in the list, install the current [Java JDK](https://www.oracle.com/java/technologies/downloads/). (On macOS and Windows, choose the “Installer” for the easiest installation. On an Apple Silicon Mac, choose the ARM64 installer, and make sure you’re running the Apple Silicon version of LibreOffice.)
3.  If you believe your Java configuration is correct and you’re still getting an error for a manual installation attempt, you can try deleting some or all of your [LibreOffice profile folder](https://wiki.documentfoundation.org/UserProfile#Default_locations).
4.  If you are working on a computer provided by your institution, contact your IT department.
5.  Install the plugin manually.

If the plugin doesn’t appear in Word after a manual installation, see Zotero toolbar doesn’t appear.

Note that you’ll need to repeat a manual installation every time the plugin is updated, so it’s much better to fix automatic installation. To troubleshoot the automatic installation, please create a new thread in the Zotero Forums so we can try to help. Be sure to include a Report ID from Zotero, your operating system and LibreOffice version, and the steps you’ve taken to try to fix the problem.

### Why am I getting a “disk I/O error” at Zotero startup? {#why-am-i-getting-a-disk-io-error-at-zotero-startup}

A disk I/O error is a general, unspecified problem with Zotero’s ability to access your disk. If you have your Zotero data directory on a network drive or external disk, you should move it back to the local disk to avoid these kinds of problems.

If you’re getting this error with your data directory on your local disk, there may be a genuine problem with your filesystem or physical disk, or security software on your system may be interfering with Zotero’s ability to access the database.

## Groups and collaboration {#groups-and-collaboration}
### Sharing data directory between Zotero Standalone and Zotero for Firefox {#sharing-data-directory-between-zotero-standalone-and-zotero-for-firefox}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

When you install both Zotero Standalone and Zotero for Firefox, you will be asked if you want to share your data directory. The general recommendation is to respond “Yes”, because this will usually do what you expect Zotero to do. Here’s some more detail on what’s going on.

#### Why share the data directory#

When you add new items to your Zotero library, the metadata for those items is stored inside Zotero’s database (zotero.sqlite) and file attachments are placed inside a special “storage” folder both of which are located inside your Zotero data directory. By default, the data directory is located inside the profile\_directory. When Zotero for Firefox and Zotero Standalone are using distinct data directories, they are not “aware” of each other’s presence and data between the two does not transfer as smoothly (or at all).

If Zotero Standalone and Zotero for Firefox share the same data directory, you will observe the following benefits:

1\. Since items are saved to the same database, your data is always in-sync. Furthermore, if you open Zotero for Firefox and Zotero Standalone at the same time, Zotero for Firefox switches into a light-weight connector mode (very similar to Chrome/Opera/Safari connectors), where all the heavy lifting involving the database is performed inside Zotero Standalone. This may help speed up Firefox performance, but it also means that you *cannot* open up Zotero for Firefox pane while Zotero Standalone is open (if you try, this will simply focus the Zotero Standalone window).

If the two were not sharing the same data directory, you would notice that adding an item to, say, Zotero for Firefox, does not immediately cause the item to appear in Zotero Standalone. If you have sync enabled in both Zotero Standalone and Zotero for Firefox, items do transfer eventually. This is because Zotero for Firefox ends up uploading the new item to zotero.org and Zotero Standalone, in turn, downloads it back onto your computer. As you can imagine, this is a slow process that results in useless utilization of your internet bandwidth.

2\. In addition to “instant sync”, sharing the data directory ensures that your data is not needlessly duplicated on your computer. When directories are shared, both metadata and file attachments are stored in one central location. Without sharing the data directory, both Zotero for Firefox and Zotero Standalone keep individual copies of each. Depending on the size of your library, this could result in gigabytes of duplicated data.

3\. Under rare circumstances, because the data is not instantaneously synced and you are able to edit your library in both Zotero for Firefox and Zotero Standalone, *not* sharing the data directory could result in sync conflicts. Sharing the data directory eliminates this possibility.

#### Sharing the data directory manually#

Normally sharing the data directory happens automatically. When you install both Zotero for Firefox and Zotero Standalone, whichever you open up second asks if you want Zotero to share the data directory. Click “Yes” (the default) and everything is configured automatically.

However, if the prompt never appears or the data directories end up separate for some other reason, you can follow these steps to share the data directory between Zotero Standalone and Zotero for Firefox.

1\. Determine which data directory you want to use. Usually you can just pick either Zotero Standalone or Zotero for Firefox at random. If, however, the two databases are not the same (e.g. you did not have sync set up and you added a bunch of metadata to Zotero for Firefox, but it did not show up in Zotero Standalone), you should pick the more complete database.

If neither database is complete, first merge them by syncing both applications (make sure that the sync completes without errors or you may end up losing some data!). Once the databases are in sync, you can pick either one of them to be the main database.

2\. Open up Preferences for the application that you chose to keep the data directory. Under Advanced -> Files and Folders, click Show Data Directory. Note the path of the directory that is displayed.

3\. Open the other application (either Zotero for Firefox or Zotero Standalone) and navigate to Preferences -> Advanced -> Files and Folders. In the “Data Directory Location” section, select Custom and choose the directory that you noted in step 2 (this should contain a zotero.sqlite file). The application will tell you that it needs to be restarted. Do so.

The data directories between Zotero Standalone and Zotero for Firefox should now be shared. See below for ways to confirm this.

#### How do I know if my data directories are shared?#

The simplest way to confirm this is by opening both Zotero Stadalone and Firefox and then clicking the Zotero icon in Firefox. If the data directories are shared, this should bring Zotero Standalone window into focus. Otherwise, Zotero will open up inside Firefox.

#### Why \*not\* share the data directory?#

You may want to opt against sharing the data directory if you require two entirely separate Zotero libraries. One way to accomplish this is to have separate data directories for Zotero Standalone and Zotero for Firefox. You can then use them as separate libraries with completely different sets of items. (An alternative to such an approach is using [Firefox/Zotero Standalone profiles](https://support.mozilla.org/en-US/kb/profile-manager-create-and-remove-firefox-profiles)).

#### Preferences are not shared#

While all your data is stored inside the data directory, your preferences are stored inside your profile directory. Sharing your data directory does *not* share your profile directory, so preferences are not shared between Zotero for Firefox and Zotero Standalone. This means that if sometimes you use Zotero Standalone and sometimes you’re using Zotero for Firefox on its own, you need to set up preferences (most importantly, sync under Preferences -> Sync) for both of them separately.

## Organising your library {#organising-your-library}
### Annotation {#annotation}

Zotero’s annotation and highlighting functionality is no longer supported.

Third-party tools can be used to annotate and highlight web pages before saving them to Zotero.

For PDFs, PDF annotation tools that save annotations directly to files will allow those changes to sync properly when using Zotero file syncing.

### Collections and Tags {#collections-and-tags}

Items in Zotero libraries can be organized with **collections** and **tags**.

**Collections** allow hierarchical organization of items into groups and subgroups. The same item can belong to multiple collections and subcollections in your library at the same time. Collections are useful for filing items in meaningful groups (e.g., items for a particular project, from a specific source, on a specific topic, or for a particular course). You can import items directly to a specific collection or add them to collections after they are already in your library.

**Tags** (often called “keywords” in other contexts) allow for detailed characterization of an item. You can tag items based on their topics, methods, status, ratings, or even based on your own workflow (e.g., “to-read”). Items can have as many tags as you like, and you can filter your library (or a specific collection) to show items having a specific set of one or more tags.

Tags are portable, but collections are not. Copying items between Zotero libraries (My Library and group libraries) will transfer their tags, but not their collection placements. Both organizational methods have unique advantages and features. Experiment with both to see what works best for your own workflow.

#### Collections#
##### The Zotero Collections Model#

It is important to understand that items can belong to multiple collections and subcollections. Adding an item to multiple collections **does not duplicate** the item. In this sense, collections are more like music playlists than folders in your computer filesystem. Just as a single song can be added to more than one playlist, a single item in a Zotero library can be added to multiple collections.

The library root — “My Library” for personal libraries or the group name for group libraries — always shows all items in the library, and items are duplicates only if they appear multiple times in that view.

##### Creating Collections#

Click the “New Collection…” button () above the left pane in Zotero to add a new collection. You can also right-click on “My Library” or the name of a Group library and choose “New Collection…” to add a new collection. The new collection will appear as a folder under “My Library” or the selected Group library.

Subcollections can be created by dragging and dropping an existing collection onto another collection or by right-clicking a collection and choosing “New Subcollection…”. You can convert a subcollection into a top-level collection by dragging it to the library root (e.g., “My Library”).

You can drag collections between Zotero libraries in the left pane. If you have editing privileges for the destination library, the collection (and all of its items and subcollections) will be added to the destination library.

##### Adding Items to Collections#

New items are automatically saved into the collection selected at the time. When saving from the browser, you can choose a different collection from the popup after clicking the save button.

To add existing items in your library to a collection, select them in the center pane and drag them onto the collection in the left pane. By default, the items will be added to the new collection but won’t be removed from their original location.

To *move* items between collections, hold down Cmd (Mac) or Shift (Windows/Linux) when dragging items to the new collection. Remember that the “My Library” view always shows all items in your library, so you cannot “move” items out of this view. To see only items that do not belong to any collection in your library, open the Unfiled Items special collection.

##### Renaming Collections#

Right-click on the collection and choose “Rename Collection…” to give a collection a new name. Collections are sorted alphabetically in your library. You can use punctuation marks to sort collections to the top of the list.

##### Deleting Collections#

Right-click on the collection and choose “Delete Collection…” to remove a collection from your library. Deleting a collection **does not delete** the items in the collection. Items are still accessible by clicking on My Library or the Group library name.

Deleting a collection will also delete its subcollections (but not the items in them).

To delete both the collection and its items, right-click on the collection and choose “Delete Collection and Items…”.

##### Removing Items from a Collection#

Select a collection in the left Zotero pane, then select the items in the center pane that you want to remove from the collection. Press the Delete key to remove the items from the collection. You can also right-click the selected items and choose “Remove Item(s) from Collection…”. This only removes the items from the selected collection, not from “My Library” or any other collections.

To *delete* an item in a collection, select the items in the collection, right-click on them, and choose “Move Item(s) to Trash…”. When “My Library” or a Group library name is selected in the left Zotero pane, pressing the Delete key will also move items to the trash. When a collection or subcollection is selected, press Cmd/Ctrl-Delete to move the items directly to the trash.

##### Identifying Collections an Item is In#

See How can I see what collections my item is in?

##### Saved Searches#

Saved Searches are like smart collections. They automatically update to include all the items in your library that meet the search criteria. For example, you can create a “To Read” Saved Search with the criteria 0 1 2. Then, opening the Saved Search will show you all of the items in your library that you haven’t read yet (i.e., the ones without a “read” tag). To mark an item as read, add the “read” tag to it.

##### Special Collections#
##### My Publications#

The My Publications special collection allows you to share your own research work (including items, notes, and attachment files) publicly with the world. Drag your publications to this collection to launch the My Publication wizard, which will allow you to select which notes, links, and files are shared.

##### Duplicate Items#

The Duplicate Items special collection shows items that Zotero has identified as potential duplicates. You can use this collection to review and merge duplicate items. For details on how duplicate detection works and how to merge duplicates, see here.

##### Unfiled Items#

Items that are not in any collection can be found in the “Unfiled Items” special collection at the bottom of the collections list in the left Zotero pane. If the “Unfiled Items” collection is not visible, right-click on “My Library” or the Group library name and choose “Show Unfiled Items”.

##### Trash#

When you delete an item by pressing Delete (Cmd/Ctrl-Delete in a collection) or by right-clicking and choosing “Move Item(s) to Trash”, they are moved to the Trash special collection. By default, items remain in the Trash for 30 days (you can adjust this period in the General pane of Zotero preferences), after which they are deleted permanently. You can restore an item from the Trash to your library by selecting it in the Trash and clicking the “Restore to Library” button (or by right-clicking and choosing “Restore to Library”). You can delete an item permanently by clicking the “Delete Permanently…” button (or by right-clicking and choosing “Delete Item…”.

##### Show Items from Subcollections#

By default, items added to a subcollection do not automatically appear in parent collections. This can be changed by toggling “Show Items from Subcollections” in the View menu. This setting is not currently available in the online library.

#### Tags#
##### The Tag Selector#

The tag selector is located at the bottom of the left Zotero pane. It shows all the tags that have been assigned to the items currently shown in the center pane (i.e., items in the currently selected collection that match the current search). To show all of the tags present in the library, click the multi-color button in the lower-right corner of the tag selection and choose “Display All Tags in This Library”. Tags not assigned to currently shown items are greyed out. You can filter items by their tags by clicking on one more tags in the tag selector. Only items that have all of the selected tags will be shown in the center pane. As you apply filters, the list of tags in the tag selector will be updated to show only the tags assigned to currently visible items. Clicking on a tag a again will deselect it. To deselect all tags at once, click the multi-color button and choose “Deselect All”.

The filter box at the bottom of the tag selector can be used to search for tags. Type in the search box to see all tags that match. To go back to viewing all the tags in the collection, press the Escape key or click the “X” button on the right.

##### Adding Tags to Items#

To add a tag to an item, select the item in the center Zotero pane, and then click the Tags tab in the right-hand pane. Click the Add button, type the tag name, and type Enter/Return. You can rename a tag by clicking on it and typing the new name. Once you have added the tag it will also appear in the tag selector in the bottom left.

As you type, you will be shown a list of matching existing tags. You can choose a suggested tag using the up and down arrow keys and insert it by pressing Tab or Enter/Return.

You can also drag items to a tag in the tag selector in the bottom left to quickly add the tag to all of those items.

##### Removing Tags from Items#

To remove a tag from an item, select the item in the items list, click the Tags tab in the right-hand pane, and click the “-” (minus) button next to the tag you want to remove.

You can also hold Cmd (Mac) or Shift (Windows/Linux) while dragging items to a tag in the tag selector in the bottom left to remove the tag from all of those items.

If you’ve assigned a color to a tag, you can also remove the tag from all selected items by pressing the number key associated with that tag on your keyboard.

##### Bulk Editing Tags#

You can use the tag selector to globally rename and delete tags. To rename a tag across all items it is assigned to, right-click the tag in the tag selector and choose “Rename Tag…”. To delete a tag from all items, right-click the tag and choose “Delete Tag…”. You can merge tags by renaming one to have the same name as the other.

##### Emoji Tags#

If you use an emoji in a tag, it will be displayed directly in the items list.

##### Colored Tags#

Colored tags appear as small colored squares next to items’ titles in the center pane. Colored tags make it easy to quickly scan your library for tags that have a certain tag. Colored tags are shown at the top of the tag selector and are always visible, even if not assigned to any visible items.

Many people use colored tags for “to read” or “favorite” items.

You can assign a color to a tag by right-clicking on it and choosing “Assign Color…” In the popup window, select a color from the dropdown box and click “Set Color”. You can remove a color from a tag by right-clicking, choosing “Assign Color…”, and clicking “Remove Color”.

Each colored tag is also assigned a number, corresponding to its position at the top of the tag selector. You can quickly add or remove a colored tag from selected items by typing the corresponding number key on your keyboard. You can change a colored tag’s number by right-clicking on it, choosing “Assign Color…”, and choosing a new position.

If you use an emoji in a colored tag, the emoji will be displayed directly in place of the colored square. The tag will otherwise behave like a colored tag, showing at the top of the tag selector and being assignable with a number key.

Up to 9 tags can be assigned colors and numbers. For further documentation see [Colored Tags](https://zotero-manual.github.io/tags/#colored-tags).

##### Automatic Tags#

When items are saved to a Zotero library from the web, tags are sometimes automatically added to items. For example, OPAC library catalogs provide subject headings for their records, which are saved as Zotero tags. Automatic tags behave the same as manually added tags but are marked by a red icon in the “Tags” tab of the right-hand Zotero pane (versus the blue icon for regular tags).

Automatic tags can be hidden from the tag selector by clicking the multi-color button in the lower-right corner of the tag selector and unchecking “Show Automatic”. You can delete all automatic tags from a library by clicking “Delete Automatic Tags in This Library”. To prevent Zotero from adding automatic tags, uncheck “Automatically tag items with keywords and subject headings” in the General pane of Zotero preferences.

### Do annotations sync? {#do-annotations-sync}
#### Do annotations sync?#
#### PDFs#

If you annotate your PDF files using third-party software (such as Adobe Reader, FoxIt, PDF-XChange, or Mac OS X’s Preview) that writes annotations to the file, the synced file will include your annotations.

#### HTML snapshots#

Zotero syncing does not include the annotations and highlighting you have made on HTML snapshots using Zotero’s (no longer maintained) annotation feature. Annotations from other tools will sync if they are saved to the HTML snapshots

### Duplicate Detection {#duplicate-detection}

As you build your Zotero library, you might introduce a few duplicated items. For example, you could have saved the same item twice from a webpage, or imported items already in your library. Fortunately, Zotero can help you identify possible duplicates and allow you to merge them.

#### Finding Duplicates#

Clicking on the “Duplicate Items” collection in your library or right-clicking the library in the left pane and selecting “Show Duplicates” will show the items Zotero thinks are duplicates in the center pane.

Zotero currently uses the the title, DOI, and ISBN fields to determine duplicates. If these fields match (or are absent), Zotero also compares the years of publication (if they are within a year of each other) and author/creator lists (if at least one author last name plus first initial matches) to determine duplicates. The algorithm will be improved in the future to incorporate other fields.

At this time, it is not possible to mark false positive matches as non-duplicates. This functionality will be added in the future.

Note that duplicate detection only works *within a library*. Items in different group libraries are separate items. They won’t show up in the “Duplicate Items” collection of any one of the libraries.

#### Merging Duplicates#

You should always resolve duplicate items by merging them, rather than deleting one of the duplicates. Merges will retain all of the collections and tags of the merged items; deleting one item will lose these data. Merges are also automatically recognized by the word processor plugins and don’t affect your automatically generated citations and bibliographies.

To merge items in the “Duplicate Items” collection, select an item in the center pane. Zotero will automatically co-select the other items that it thinks are duplicates. Click the “Merge <number> Items” button in the right pane to merge the items. If the item fields don’t match completely, you can select one item to be the “master” from the list at the top of the right pane, then select alternative versions of mismatched fields using the icons to the right of each field.

It may be easier to see which items are selected if you sort the items by Title. You can select a single item in the “Duplicate Items” view by holding down Alt/Option while clicking. You can de-select an item from a set of duplicates by holding down Ctrl (Windows/Linux) or Cmd (Mac) while clicking.

You can also select a group of two or more items *of the same item type* anywhere in your Zotero library, right-click, and select “Merge Items…” from the context menu.

### How can I access my library from multiple computers? Can I store my Zotero library and associated files on an external drive? {#how-can-i-access-my-library-from-multiple-computers-can-i-store-my-zotero-librar}
#### How can I access my library from multiple computers? Can I store my Zotero library and associated files on an external drive?#

The best way to access your Zotero library from multiple computers is to use Zotero syncing. Zotero syncing will automatically sync your library data using the Zotero servers. Attachment files can be synced using the Zotero servers or using a third-party WebDAV servers.

You absolutely should not store your Zotero database in a cloud storage folder (e.g., Dropbox, Google Drive, OneDrive), which will lead to data corruption. You can, however, configure Zotero to sync your attachment files using these services while using Zotero syncing to sync your library data.

If, for whatever reason, you cannot use Zotero’s built-in sync features, you can also store your Zotero data directory on an external hard drive and use this to move your Zotero data between computers. You can set your Zotero installation on each computer to point to the external drive in the Advanced pane of Zotero preferences.

Another option is to run a copy of Zotero directly from a portable drive.

### How can I move my Zotero library to a different computer? {#how-can-i-move-my-zotero-library-to-a-different-computer}
#### Option A: Use Zotero Sync#

The easiest way to transfer your library between between computers is by using Zotero Sync. Set up syncing from the Sync pane of the Zotero preferences, making sure you use the same username on all computers.

You’ll need enough online storage space to fit all files in your library. Zotero will warn you if you hit your quota, in which case you may need to delete some files, add a storage plan, or transfer your library using Option B below.

#### Option B: Copy the data folder#

If you’re comfortable moving files between devices, you can transfer your library by copying the Zotero data folder from your first computer to your new computer — e.g., by using an external hard drive or a local network connection. If switching to a new computer, your OS may provide a way to automatically transfer all data (e.g., Migration Assistant on a Mac).

To move your data manually, first locate your data by opening the Zotero preferences, going to Advanced → Files and Folders, and clicking “Show Data Directory”. See here for the default locations of the data folder.

Be sure to close Zotero on both machines before copying the Zotero files. If you’ve already opened Zotero on the new computer, there will already be a Zotero data folder with an empty database, and you should delete the whole data folder before copying the new folder to the same location.

#### Don’t Use Export/Import#

Exporting and importing your library (for instance via Zotero RDF) is not a recommended option. None of the available export formats allow for a complete transfer of your library data, reimporting will break connections with any existing word processor documents, and if you use syncing later you will end up with duplicates of any imported items.

### How can I see how many items I have in my Zotero library? {#how-can-i-see-how-many-items-i-have-in-my-zotero-library}

When no items are selected, Zotero’s right-hand pane displays the number of items in the current view, so you can see the number of items in a given library or collection simply by making a selection in the left pane. If items are currently selected, click to another collection and back again to view the item count.

To determine total items, including child attachments and notes, click the items list, press the + (plus) key to expand all parent items, and then either deselect the selected item (Cmd-click (Mac) or Ctrl-click (Windows/Linux)) or use Select All (Cmd-A (Mac) or Ctrl-A (Windows/Linux)). You can press “-” (minus) afterward to collapse all items.

To determine how many top-level items match a search that causes child items to be expanded, click the items list, press the “-” (minus) key to collapse all items, and deselect the selected item with Cmd-click (Mac) or Ctrl-click (Windows/Linux).

### How can I see what collections my item is in? {#how-can-i-see-what-collections-my-item-is-in}

Zotero 7 includes a section in the item pane showing a list of collections an item is in.

You can also select the item and then hold down the Option key (macOS), Control key (Windows), or Alt (Zotero 6)/Ctrl (Zotero 7) key (Linux). This will highlight all collections that contain the selected item.

### How do I export my Zotero library? {#how-do-i-export-my-zotero-library}
#### How do I export my Zotero library?#

To export an entire library, right-click on it in the Zotero collections pane and choose “Export Library…”, or select “Export Library…” from the “File” menu. To export an individual collection, right-click on it and choose “Export Collection…”. To export specific items, select them in the items list, right-click, and choose “Export Items…”.

When sharing items with another Zotero user, select Zotero RDF with files and notes for the most complete transfer.

**Please note:** An export is not a proper backup, not recommended for transferring entire Zotero libraries between computers, and not a replacement for ongoing collaboration using Zotero groups. Reimported items will have new Date Added/Modified times, will no longer be linked to citations in existing word processor documents, and may have minor changes in a small number of fields or, for some formats, even may be missing some fields altogether.

### How do I get my Zotero collection to work with an Exhibit or Citeline presentation? {#how-do-i-get-my-zotero-collection-to-work-with-an-exhibit-or-citeline-presentati}

**These directions have not been updated in some time and may no longer be up to date.**

#### How do I get my Zotero collection to work with an Exhibit or Citeline presentation?#

The [Simile Project](http://simile.mit.edu/), sponsored by MIT, has developed several JavaScript applications that help a Zotero user integrate a bibliography into a website.

[Citeline](http://citeline.mit.edu/), under Simile’s purview, has a Firefox plugin called Zotz to streamline the conversion process. This plug-in, however, works only with Firefox 2.0. Further, when importing the BibTeX file, the Citeline filter incorrectly handles Unicode characters beyond the Basic Latin set. Boxes or question marks are frequently encountered. At this time Citeline is not able to provide support for non-Latin characters.

[Exhibit](http://www.simile-widgets.org/exhibit/), now supported independent of Simile, is very similar to Citeline, but is more powerful and customizable. Exhibit has a page explaining how to [generate a publications exhibit](http://www.simile-widgets.org/wiki/How_to_make_a_publications_exhibit). But a different kind of problem is encountered here: one must convert a Zotero collection to Exhibit’s JSON format. And Zotero does not directly export to a JSON format.

The handiest conversion tool is Simile’s [Babel converter](http://service.simile-widgets.org/babel/) ([alternate link](http://simile.mit.edu/babel/)). But in using Babel one encounters conversion problems similar to those found in the Citeline conversion. Exhibit calls for Unicode characters rendered as UTF-32, but Babel renders them UTF-8. Further, the characters < and > are converted to {\\textless} and {\\textgreater}. And if you have started with a Zotero-generated BibTeX file, you may find curled braces {} in fields where they do not belong (an idiosyncrasy in Zotero’s BibTeX converter). So some cleanup is necessary before the JSON file exported by Babel is ready for interaction with Exhibit.

Here is a detailed explanation of the problems in the Babel conversion process: when Babel encounters a character outside the Basic Latin table (but less than value U+0800), Babel replaces it with two new character sequences: one a multiplicand (e.g., \u00C3 and upward, depending upon how large the target Unicode number is) and the other the Unicode value of the character point in the Basic Latin plane that is an exact multiple of 64 less than the intended Unicode value. For example, é (Unicode value U+00E9) is translated by Babel, not as \u00E9 (as Exhibit calls for) but into \u00C3\u00A9, which displays as Ã©.

### How do I merge tags? {#how-do-i-merge-tags}
#### How do I merge tags?#

Equivalent tags (e.g. “webpage” and “website”) can be merged by renaming one tag to the name of the other. Right-click one of the tags in the tag selector (bottom left of the Zotero pane), select “Rename Tag…” and enter the name of the matching tag.

### How do I print my Zotero references and notes? {#how-do-i-print-my-zotero-references-and-notes}
#### How do I print my Zotero references and notes?#

Select the collection or items you would like to print out in your report with your cursor and then right click on them (ctrl-click on a Mac). Select “Generate report from selected item.” A report featuring all of your Zotero information should pop up in your browser. If there are any notes associated with individual items or collections, those will be included with the report.

### How do I sort my notes by page number? {#how-do-i-sort-my-notes-by-page-number}
#### How do I sort my notes by page number?#

Notes are sorted alphabetically by default, so you can start each note with the page number, adding leading zeros, i.e. \[002\], \[045\], \[367\]. For multiple notes from the same page, enter \[367a\], \[367b\], etc.

*Note:* Adding a large number of notes to many items can negatively impact Zotero’s performance. We recommend storing large numbers of comments about individual items in a single attached child note.

### Library Item Overview {#library-item-overview}

A Zotero library is made up of “items”. Items can take several forms: regular items, attachments, and notes.

All items can have tags applied to them and can be linked to related items.

#### Regular Items#

The foundational elements in most Zotero libraries, *regular items* take the form of reference types—books, journal articles, manuscripts—and have associated bibliographic metadata (Title, Author, Publisher, etc.). Regular items can be created manually by using the New Item drop-down menu or automatically by clicking the Zotero address bar icon to save from a supported website.

You generally always want to work with regular Zotero items, rather than bare attachment files (e.g., PDFs, word processor documents), as regular items have item metadata that is required to interact with most of Zotero’s features.

Regular items can only be top-level items. Attachments and notes can be added to regular items as child items.

#### Attachments#

Attachments are files and web links without full bibliographic metadata. There are four kinds of attachment items:

##### Web Links and Other URIs#

Web links are essentially bookmarks to websites. When you save a link, Zotero stores only the page title, URL, and access date, and you need to return to the site to view the page content. You can also attach links to the URIs for other programs, such as or evernote:// links.

##### Snapshots#

Snapshots contain the same information as web links, but Zotero also saves a copy of the page as it currently exists so that you can view it later, even if the original webpage has changed or disappeared.

Snapshots can be single files, such as PDFs, or consist of multiple files, as is the case with an HTML page and its associated images.

##### Linked Files#

Linked files are links to files stored outside of the Zotero data directory on your computer.

##### Imported Files#

Imported files are files stored within the Zotero data directory. When you import a file (either using “Store Copy of File” or by dragging in a file), Zotero copies the file to its data directory, leaving the original untouched. After importing a file, you may wish to delete the external copy to avoid confusion.

#### \#

Each attachment has a single embedded note field.

Attachments can either be *child attachments* (attached to regular items) or *standalone attachments* (top-level items not attached to regular items). Attachments cannot have child items attached to them.

As of Zotero 2.0b3, web links and snapshots can only be created as child items (with the temporary exception of PDF snapshots, which can still be created as top-level items to allow for use of the Retrieve PDF Metadata feature).

#### Notes#

Notes are pieces of text without full bibliographic metadata.

Notes can either be *child notes* (attached to regular items) or *standalone notes* (top-level items not attached to regular items). Notes cannot have child items attached to them.

While the Zotero program will allow you to insert embedded images into notes, they will currently not sync (and make prevent your Zotero library from syncing altogether). Improved support for embedding images in notes is planned.

### Note Templates {#note-templates}
#### Annotations#

You can use note templates to customize how PDF annotations are added to notes. To view the available templates, go to the Advanced pane of the Zotero preferences, open the Config Editor, and search for 0. There are currently three available templates: for highlight annotations, for note annotations, and for the title of notes created from all of an item’s annotation.

Templates support basic HTML, with variables within curly brackets. Here’s the default template for highlights:

    {{:highlight}} {{:citation}} {{:comment}}

You can see that, by default, highlight annotations are added as a single paragraph, with the highlighted text followed by the citation and any comment. Quotation marks are automatically added around the highlight.

If you prefer to have the highlight text in a blockquote, it’s a simple change:

    <blockquote>{{:highlight}}</blockquote>{{:citation}} {{:comment}}

##### Conditionals#

Templates also support conditionals. Rather than combining the citation and comment in a single paragraph as in the previous example, you might want to create a separate paragraph for the comment, but only if a comment actually exists. You can test whether a variable is set with a simple 0:

    <blockquote>{{:highlight}}</blockquote>{{:citation}}{{:if comment}}{{:comment}}{{:endif}}

Conditionals can also be used to test for specific values. Here, text highlighted in red becomes a header without quotation marks, text highlighted in blue becomes a blockquote, and all other highlights use a single paragraph:

    {{:if color == '#ff6666'}}
        {{:highlight quotes='false'}}
    {{:elseif color == '#2ea8e5'}}
        {{:if comment}}{{:comment}}:{{:endif}}<blockquote>{{:highlight}}</blockquote>{{:citation}}
    {{:else}}
        {{:highlight}} {{:citation}} {{:comment}}{{:if tags}} #{{:tags join=' #'}}{{:endif}}
    {{:endif}}

##### Variables#

Here are the available variables and their supported parameters:

-   0
    -   0
        -   omitted: Include quotation marks around text unless the highlight is placed within a blockquote
        -   “true”: Always include quotation marks
        -   “false”: Never include quotation marks. The highlight must be placed within a blockquote to remain as an active annotation.
-   0
-   0
-   0 — yellow: ‘#ffd400’, red: ‘#ff6666’, green: ‘#5fb236’, blue: ‘#2ea8e5’, purple: ‘#a28ae5’, magenta: ‘#e56eee’, orange: ‘#f19837’, gray: ‘#aaaaaa’
-   0
    -   0 — string to use to join tags

(Note that 0 is primarily for use in conditionals. Annotation colors can be toggled on and off from the note editor menu. An upcoming version will add a preference for controlling whether colors are shown by default.)

### Notes {#notes}

In addition to items and file attachments, you can also store notes in your Zotero library: child notes, which belong to a specific item, and standalone notes. Notes are synced along with item metadata; they don’t count against your Zotero file storage quota.

#### Child Notes#

To create a child note, select an item in the center pane. Then, either click the “New Note” button at the top of the center pane () and select “Add Child Note”, or go to the “Notes” tab in the right-hand pane and click the “Add” button. You can also right-click an item and select “Add Note”.

A note will be created as an attachment to the item (it will also show up under the “Notes” tab), and a note editor will appear in the right-hand pane. You can create a dedicated window for the note editor by clicking the “Edit in a separate window” button at the bottom of the editor. Text in notes is saved as you type.

#### Standalone Notes#

Standalone notes work the same as child notes, but are not directly related to any item in your library. Standalone notes will appear alongside any other items in your library. To create a standalone note, click the “New Note” () button and select “New Standalone Note”.

#### Tagging and Relating#

As with any other item in Zotero, notes can be tagged and related to other items. Both features can be accessed via the bottom of the note editor.

#### Searching Notes#

Notes can be searched via the general Zotero search. You can search text within a note using the note editor’s “Find and Replace” button or Ctrl/Cmd-F.

#### Images in Notes#

You can embed images into Zotero notes by dragging them from your filesystem or, in some cases, by pasting them in. Embedded images sync between devices and count against your Zotero Storage quota.

### Related Items {#related-items}

In addition to collections and tags, a third way to express relationships between items is by setting up “relations”. Relations can set up between any pair of items in a library (it is not possible to relate items from different libraries).

To create a relation, select an item in the center pane and go to the “Related” tab of the right pane. Click the “Add” button, and select one or more items from the same library in the pop-up window (hold down Ctrl or Shift \[Windows/Linux\] or Cmd or Shift \[Mac\] to select multiple items) and click “OK”. The selected items will now show up as related items in the “Related” tab, and clicking an item will take you straight to that item.

With the [Zutilo plugin](https://github.com/willsALMANJ/Zutilo), you can also select multiple items in the center pane, then right-click and choose “Related selected items” to related all selected items to each other in one click.

Note that when you relate item A to B, B will be automatically related to A. But relations are not [transitive](http://en.wikipedia.org/wiki/Transitive_relation): relating A to B, and B to C, will not automatically relate A to C.

Some suggestions of how you could use this feature:

-   connect book chapters to their parent volume
-   connect book reviews to the book reviewed
-   connect different versions of a work (e.g., connecting a conference presentation that eventually became an article that eventually became a book)
-   link associated items from different collections
-   link items that form parts of a single work (e.g., articles in a series)
-   connect standalone notes to the items they discuss
-   link one item to another discussed in the Abstract or Notes fields
-   link items that have similar comments in the Abstract or Notes fields

### Search {#search}

For information on how to use Zotero’s search features, see Searching. This page describes the settings in the Search pane of the Zotero preferences.

The Search preferences pane is used to configure and manage Zotero’s PDF/EPUB/HTML full-text search feature.

#### Full-Text Cache#

Zotero creates an index to allow the full text contents of PDF and plain-text attachments in your library to be searched with Quick Search (“Everything” option) and Advanced Search (via “Attachment Content”).

**Note:** At this time, PDF, EPUB, and HTML full text content (and plain text files) can be indexed by Zotero. Other document types (e.g., .docx, .odt) cannot be indexed by Zotero.

This section includes these options to manage your full-text index:

-   **Rebuild Index…:** Re-create the full-text index for all items from scratch. This option may be helpful if you use OCR text recognition on a large number of attachment files or you find that searching using “Everything” or “Attachment Content” does not return the correct results.
    -   Before using this option, verify that the item has searchable text and that the text is properly stored in the PDF/EPUB/HTML (e.g., try to copy text out of the document and ensure that it is high quality).
    -   You can re-index individual PDF/EPUB/HTML files in your library by right-clicking on the relevant attachment item in your library and choosing “Reindex Item”.
    -   If you are still having issues after re-indexing an item, please ask a question on the Zotero Forums.
-   **Clear Index…:** Delete the full-text index. Use this option if you intend to disable full-text indexing and wish to reduce the size of your Zotero database (note that the full-text index typically occupies a relatively small amount of storage space on your computer).
-   **Maximum characters to index per file:** The number of characters to index in PDF/EPUB/HTML and plain-text attachments (default: 500,000, approximately 100,000 words or 180–200 pages of content). Set this value to 0 to disable full-text indexing.

#### Index Statistics#

This section provides details about the size of the full-text index in your Zotero database. Reported statistics are:

-   **Indexed:** The total number of files that are completely indexed.
-   **Partial:** The number of files that are partially indexed.
-   **Unindexed:** The number of files that are not yet indexed.
-   **Words:** The total number of words included in your Zotero library’s full-text index.

### Searching {#searching}
#### Quick Search#

Quick searches provide a fast way to find items in a library or collection.

##### Running a Quick Search#

To begin searching, click inside the search box at the top-right of the center pane (or type Ctrl/Cmd-F) and start typing your search terms. As you type, only those items in the center column that match the search terms will remain.

##### Quick Search Options#

Quick search can be used in three different modes:

-   “Title, Year, Creator” - matches against these three fields, as well as publication titles
-   “All Fields & Tags” - matches against all fields, as well as tags and text in notes
-   “Everything” - matches against all fields, tags, text in notes, and indexed text in PDFs

##### Speeding Up Quick Searches#

By default, the quick search searches automatically as you type. In very large libraries, this can result in a slow search experience.

To speed things up, type a quotation mark 0 mark at the beginning of the search field. This causes the search to start only when you type 1/2.

#### Advanced Search#

Advanced searches offer more and finer control than quick searches, and allow you to make \#saved searches.

##### Running an Advanced Search#

To open the Advanced Search window, click on the magnifying glass icon () at the top of the center pane.

In this window, you can filter items by the content of specific fields or by other properties, like item type or the collection an item belongs to. Multiple filters can be set up by clicking the plus button.

Set the library to search in by using the “Search in libary:” option at the top of the window. It is not currently possible to search in multiple libraries at one time.

By default, items only show up in the search results if they satisfy all search criteria. To change the search so that all items matching at least one criterion are returned, change the “Match” option to “any”.

You can filter items by the collections or saved searches they belong to by searching by “Collection”. To include items in subcollections of matching collections in the search results, check “Search subcollections”.

To hide non-matching parent items with child items that do match the search criteria, and to collapse matching parent items with matching child items, check “Show only top-level items”.

To match search criteria against both parent items and their children, check “Include parent and child items of matching items”. If this option is selected and “Match” is set to “all”, parent/child items will still show up if just part of the criteria is met by the parent item and the other part by a child item.

##### Wild Cards#

The % percent sign character acts as a wild card in advanced searches, substituting for zero or more characters. For example, the search term “W% Shakespeare” will match “W Shakespeare”, “W. Shakespeare” as well as “William Shakespeare”.

To search for items where a field contains any content at all, search for the % percent sign character alone.

##### Saved Searches#

When you save an Advanced Search, it appears as a collection in your library (but with a Saved Search icon, , instead of the regular collection icon). Saved Searches are continuously updated. For example, if you set up a Saved Search for “Date Added” “is in the last” “7” “days”, the saved search will always show the items that have been added in the last 7 days. Saved searches only store the search criteria, not the search results.

To save a search, click the “Save Search” button in the Advanced Search window and provide a name for the search. Saved searches can be edited or deleted by right-clicking the Saved Search and selecting “Edit Saved Search…” or “Delete Saved Search…”, respectively.

You can also create a saved search in a library by right-clicking on the library name and choosing “New Saved Search…”.

##### Complex Search Criteria#

It is possible to run complex Boolean searches by using multiple Saved Searches. For example, to run the search 0, first make a Saved Search called “Condition1” for 1, then make a Saved Search called “Condition2” for 2. Finally, run a third Advanced Search and search for 3.

#### Full-Text Indexing#

Full-text indexing allows embedded text within PDFs, EPUBs, HTML, and text files to be searched with Quick Search (via the “Everything” option) and Advanced Search (via “Attachment Content”). Indexing happens automatically in the background when Zotero is idle.

You can control how much text in a PDF/EPUB/HTML/text file is indexed in the Search pane of Zotero preferences (default: 500000 characters, 100 pages). You can remove indexed text with the “Clear Index…” button or re-create the index from scratch using the “Rebuild Index…” button. You can check the index status of any PDF/EPUB/HTML attachment by selecting the attachment item in the Zotero library and looking at the “Indexed:” field in the right pane.

If an item isn’t being indexed (e.g., if it is not showing up in an ‘Everything’ Quick Search), verify that the item has searchable text and that the text is properly stored in the PDF (e.g., try to copy text out of the document and ensure that it is high quality). If the PDF has valid text, rebuild the item’s index by right-clicking on it and choosing “Reindex Item”. If you are still having issues, please ask a question on the Zotero forums.

**Note:** At this time, only PDF/EPUB/HTML full text content (and plain text files) can be indexed by Zotero. Other document types (e.g., .docx, .odt) cannot be indexed by Zotero.

### Sorting {#sorting}

Items in the center pane can be sorted by various fields, such as their title, creators, or the date they were added to your library.

To change the way items are sorted, click on any of the column headers at the top of the center pane. For example, if you click on “Title”, all your items will be sorted alphabetically by title. Clicking a header multiple times toggles between ascending and descending sorts. (The header will show an upward and downward arrow, respectively.)

By default, Zotero will show columns for the Title, Creators, and Attachments properties in the center pane. You can change which fields are shown by right-clicking on the column headers and selecting fields from the drop-down menu.

For each column, you can also choose the Secondary Sort field (the field that’s used to break ties when sorting).

By default, fields in the center pane are arranged from left to right in the order in which they are shown in the drop-down menu. You can rearrange them by dragging and dropping the headers. To reset the order, select “Restore Column Order” in the drop-down menu.

### What ways can I organize and manage my collections? {#what-ways-can-i-organize-and-manage-my-collections}
#### What ways can I organize and manage my collections?#

“My Library” holds every item in your library. To organize these items, you can create collections and subcollections on the left side of the Zotero window. Note that items can belong to multiple collections; it is best to think of collections as more like “playlists” than like folders on your computer. To add a new collection, click on the “New Collection…” icon at the top-left of the Zotero window or right-click on the left Zotero pane. To make a subcollection, right-click on a collection and choose “New Subcollection…” or drag an existing collection into another collection. To rename a collection, right-click on it (control-click on Macs) and select “Rename Collection…” To add items to a collection, drag them from the middle pane.

You can also organize your collection through tags. Click on an item in the middle pane. Now select the “Tags” tab in the right pane. Click “Add”, type the tag name, and type Enter/Return. To add multiple tags at once, type Shift+Enter/Return. Enter each tag on a new line, then type Shift+Enter/Return again to add all of the tags at once. To filter your library or a collection on a tag, select the tag from the tag selector pane in the lower-left of the Zotero window. You can also drag items from the center pane to a tag in the tag selector to add the tag to all selected items.

See Collections and Tags for more information.

### Where did the Extract Annotations button go? {#where-did-the-extract-annotations-button-go}

“Extract Annotations” was a feature of the third-party ZotFile plugin. In Zotero 6, that feature has been replaced by Zotero’s own advanced PDF functionality.

You can create a note from all of an item’s annotations by right-clicking an item in the items list and selecting “Add Note from Annotations”, or you can open the PDF in the new built-in PDF reader and add annotations selectively. See Adding Annotations to Notes for more on the various ways of using annotations in notes in Zotero 6.

For more details on the new functionality, see the Zotero 6 announcement.

## Other help pages {#other-help-pages}
### About Zotero {#about-zotero}

Zotero is a project of [Digital Scholar](http://digitalscholar.org/). It was created at the [Roy Rosenzweig Center for History and New Media](https://rrchnm.org/) at [George Mason University](https://www.gmu.edu).

#### Credits#

**Director**

-   Sean Takats

**Lead Developer**

-   Dan Stillman

**Senior Developer**

-   Faolan Cheslack-Postava

**Developers**

-   Bogdan Abaev
-   Martynas Bagdonas
-   Abe Jellinek
-   Tom Najdek
-   Dima Petrov
-   Michal Rentka
-   Miltiadis Vasilakis
-   Adomas Venčkauskas
-   Xiangyu Wang

**Design**

-   Yexing Sha

**Alumni**

-   Roy Rosenzweig (Executive Producer)
-   Johannes Krtek (Designer, 2017–2022; new app icon, 2024)
-   Zoë Ma (Translator Development, 2023)
-   Brendan O’Connell (Outreach Coordinator, 2022–2023)
-   Fletcher Hazlehurst (Developer, 2020–2021)
-   Simon Kornblith (Senior Developer, 2006–2016)
-   Kim Nguyen (Site Design, 2010-2012)
-   Ben Schneider (Graduate Research Assistant, 2013-2015)
-   Debbie Maron (Community Outreach, 2011)
-   Dan Cohen (Co-Director, 2006-2010)
-   Jeremy Boggs (Site Design, 2006-2011)
-   Trevor Owens (Technology Evangelist, 2007-2010)
-   Connie Sehat (Co-Director, 2007-2009)
-   Fred Gibbs (Developer, 2007-2009)
-   Matt Burton (Developer, 2008-2009)
-   Jon Lesser (Developer, 2008-2009)
-   Andrew Howard (Technical Support, 2009)
-   Elena Razlogova (Project Guidance, 2006-2008)
-   Shekhar Krishnan (Evangelist, 2007-2008)
-   Michael Berkowitz (Developer, 2007-2008)
-   Asa Kusuma (Intern Developer, 2007-2008)
-   Ben Parr (Intern Developer, 2007-2008)
-   Ramesh Srigiriraju (Intern Developer, 2007)
-   Kari Kraus (Technology Evangelist, 2006-2007)
-   Josh Greenberg (Co-Director, 2006-2007)
-   David Norton (Developer, 2006-2007)

##### Community Contributors#

Zotero benefits from a strong user community. Exceptional contributions have been made by:

-   Frank Bennett ([citeproc-js](https://github.com/Juris-M/citeproc-js), [Citation Style Language](http://citationstyles.org/), user support, [Juris-M](https://juris-m.github.io/))
-   Bruce D’Arcus ([Citation Style Language](http://citationstyles.org/), user support)
-   Emiliano Heyns ([Better BibTeX](https://retorque.re/zotero-better-bibtex/), client improvements)
-   Sebastian Karcher (user support, wiki documentation, CSL styles, web translators, translations)
-   Sylvester Keil ([anystyle.io](https://anystyle.io/), [arkivo](https://github.com/inukshuk/arkivo), web services integration)
-   Joscha Legewie ([ZotFile](http://zotfile.com/), client improvements)
-   Avram Lyon (user support, [Zandy](http://www.gimranov.com/avram/w/zandy-user-guide), web translators, wiki documentation)
-   [POBrien333](https://github.com/POBrien333) (CSL styles)
-   Julian Onions (CSL styles)
-   Jason Puckett (third-party documentation)
-   Mikko Rönkkö ([ZotPad](http://www.zotpad.com/), user support)
-   Will Shanks (plugin development, client improvements)
-   Michael Steele (Zotero Standalone installer improvements)
-   Aurimas Vinckevicius (web translators, client improvements, user support)
-   Brenton M. Wiernik (user support, wiki documentation, CSL styles)
-   Rintze Zelle ([Citation Style Language](http://citationstyles.org/), user support, wiki documentation, CSL styles, translators, translations)
-   Philipp Zumstein (translators, Scaffold, user support)

##### Localization#

For the list of contributors to each locale, see [the Zotero project on Transifex](https://explore.transifex.com/zotero/zotero).

Spell-checking in Zotero relies on dictionaries from [Mozilla contributors](https://addons.mozilla.org/en-US/firefox/language-tools/).

#### Third-Party Software#

Zotero is built on [Firefox](https://www.mozilla.org/firefox/) and depends on many other exceptional open-source projects:

-   [ACE Editor](https://ace.c9.io/)
-   [Citation Style Language](https://citationstyles.org/)
-   [citeproc-js (Frank Bennett)](https://github.com/Juris-M/citeproc-js)
-   [Dark Reader](https://darkreader.org/)
-   [epub.js](https://github.com/futurepress/epub.js)
-   [KaTeX](https://katex.org/)
-   [Monaco Editor](https://microsoft.github.io/monaco-editor/)
-   [Pastel SVG icons (Michael Buckley)](https://codefisher.org/pastel-svg/)
-   [pdf.js](https://mozilla.github.io/pdf.js/)
-   [ProseMirror](https://prosemirror.net/)
-   [proseMirror-math](https://github.com/benrbray/prosemirror-math)
-   [React](https://reactjs.org/)
-   [RNV](http://www.davidashen.net/rnv.html)
-   [SingleFile (Gildas Lormeau)](https://github.com/gildas-lormeau/SingleFile)
-   [Tabulator RDF parser](http://www.w3.org/2005/ajar/tab)

#### Funding Acknowledgments#

Zotero was generously funded by the United States [Institute of Museum and Library Services](http://www.imls.gov), [The Andrew W. Mellon Foundation](http://www.mellon.org), and the [Alfred P. Sloan Foundation](http://www.sloan.org).

#### Initial Advisory Board (2006-2008)#

-   Raymond Yee (University of California, Berkeley)
-   Dan Chudnov (Yale University)
-   Abby Smith (Council on Library and Information Resources)
-   Clifford Lynch (Coalition for Networked Information)
-   David Seaman (Digital Library Federation)
-   Martha Anderson (Library of Congress)
-   Matthew MacArthur (Smithsonian National Museum of American History)
-   Wally Grotophorst (George Mason University Library)
-   Kathy Perry (Virtual Library of Virginia)

### Can I open snapshots in a new tab or window? (Zotero for Firefox) {#can-i-open-snapshots-in-a-new-tab-or-window-zotero-for-firefox}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### Can I open snapshots in a new tab or window? (Zotero for Firefox)#

Yes. When you double click a snapshot in your Zotero library, or right-click the snapshot and select “View Snapshot”, the snapshot will open in the currently active Firefox tab. To open the snapshot in a new tab and keep focus on the current tab, select “View Snapshot” while holding the Control key (Command key on OS X). To open the snapshot in a new tab and shift focus to the new tab, select “View Snapshot” while holding the Control and Shift keys (Command and Shift keys on OS X). To open the snapshot in a new window, selecting “View Snapshot” while holding the Shift key.

### Contact Zotero {#contact-zotero}

**In most cases, the best way to talk to someone from Zotero quickly is to post to the Zotero Forums. See Getting Help.**

-   If you’re unable to log into your account and are unable to reset your password, email support@zotero.org. For all other technical support issues, see Getting Help. We do not provide technical support via email.
-   If you have a billing question about your Zotero storage subscription or a question about institutional storage plans, email storage@zotero.org.
-   If you’re a developer interested in contributing code, using the Zotero APIs, or otherwise extending Zotero, post to the [zotero-dev mailing list](https://groups.google.com/g/zotero-dev). Note that the dev list is for technical questions and is **not** for user support or feature requests.
-   Security issues with Zotero or zotero.org that can’t be reported publicly can be reported to security@zotero.org.

For all other issues, please use the Zotero Forums, where you’ll generally receive a response quickly from a Zotero developer or expert community member.

### Does Zotero support non-Western characters? {#does-zotero-support-non-western-characters}
#### Does Zotero support non-Western characters?#

Yes, Zotero is Unicode-compliant and can process non-Western characters.

### Feeds {#feeds}

Feeds are a great way to discover new research. With feeds, you can subscribe to updates from a journal, website, publisher, institution, research group, or other source and quickly find new articles or works. If you find an item in a publication’s feed that you want to save and read further, you can add it to your Zotero library with the click of a button.

#### Subscribing to Feeds#

To subscribe to a feed, click the Add Library button above the left pane in the Zotero window. You can add feeds in two ways. First, you can add a feed using the URL provided on the journal’s (or other source’s) website. For journals, these are usually available from the journal homepage. Look for the RSS icon () or search for the journal name and “RSS feed” in a search engine. Many publishers place the RSS icon next to social media buttons or email alert links.

RSS Links on ScienceDirect, RSS Button for Sage Journals


Once you’ve found the RSS feed URL for your source, click the Add Feed menu, then choose “From URL…”. The Add Feed window will open. Paste the RSS feed URL to the URL text field. If the URL is for a valid feed, you will be able to enter a title for the feed and click “Save”. You can click “Advanced Options” to adjust how frequently a feed updates and to set how long read and unread items are kept in a feed before being removed.

You can also import a set of feeds using an OPML file (e.g., from a list of subscriptions exported from Feedly or other RSS Reader services). To import an OPML file, choose “From OPML…” from the Add Feed menu.

#### Reading Feeds#

Feeds you have subscribed to will appear at the bottom of the left pane of the Zotero window, below My Library and your Group Libraries and above the tag selector. Click on a feed to view currently available items in the feed. Right-click on a feed to manually refresh it, mark all items in the feed as read, modify feed settings, or unsubscribe from the feed.

When viewing a feed, unread items will be shown in bold text, while read items are shown in regular text. When you select an item in the feed, you can mark it as read or unread by clicking the “Mark As Unread/Read” button or pressing the 0 1` key. Save the item to your Zotero library by clicking on the “Save to My Library” button or pressing Ctrl/Cmd-Shift-S. You can save a feed item to a group library by clicking the dropdown arrow on the right side of the “Save to My Library” button and choosing the group library.

*Note, the read/unread state of items is not synced across computers when using Zotero sync.*

### File Handling Issues {#file-handling-issues}
#### Unexpected behavior (wrong action, gibberish (“%PDF…”), etc.) when opening files from Zotero#

Deleting handlers.json from your profile directory and restarting Zotero may help fix file handling issues. handlers.json stores file handling associations and will be recreated automatically the next time you restart Zotero.

#### PDFs opening in wrong application on Linux systems#

While Zotero be configured to use the system-default PDF reader, file handling on Linux can be affected by many settings. If you find that PDFs open in, say, Gimp even though you’ve set the file handler to be Okular in Dolphin/Nautilus/etc., the easiest solution is to choose a specific PDF reader from the General pane of the Zotero preferences.

If you’d like to keep the setting as “System Default”, you can try the following:

Open the file

    ~/.local/share/applications/defaults.list

in a text editor and look for the line “application/pdf=…”. Change it to

    application/pdf=kde4-okularApplication_pdf.desktop

If the file does not exist, create it and enter the following content:

    [Default Applications]
    application/pdf=kde4-okularApplication_pdf.desktop

Note: On some systems, you might need to use

    application/pdf=kde-okularApplication_pdf.desktop

instead.

For other PDF readers you might use one of the following:

    application/pdf=evince.desktop

    application/pdf=acroread.desktop

    application/pdf=Foxit-Reader.desktop

### File Renaming {#file-renaming}

Zotero automatically renames PDFs and other files saved to your library based on the bibliographic details (title, author, etc.) of the parent item, freeing you from having to sort through piles of randomly named files or manually rename each new file to your preferred format.

Zotero will always rename files saved from the web (via the Zotero Connector, Add Item by Identifier, or Find Available PDF).

By default, it will also rename the first stored PDFs and EPUBs you add to items, as well as files for which it successfully retrieves metadata. You can disable this by unchecking “Automatically rename attachment files using parent metadata” (“Automatically rename files” in Zotero 8) in the General pane of the Zotero settings. If an item already has an attachment, additional files will not be automatically renamed, to avoid changing the filenames of supplementary materials.

**Starting in Zotero 8,** Zotero will keep attachment filenames in sync as you make changes to parent item metadata (e.g., changing the title). In previous versions, you would need to right-click on the attachment and select Rename File from Parent Metadata to update a filename after editing metadata.

Linked files are not automatically renamed by default, but you can enable the “Rename linked files” setting in to apply renaming to those as well.

You can use “Rename files of these types” to adjust which file types are automatically renamed.

#### Attachment Title vs. Filename#

The attachment filename being renamed is separate from the attachment title shown in the items list. See Attachment Title vs. Filename for more information.

#### Customizing the Filename Format#

By default, Zotero names files after the parent item’s creator (1–2 authors or editors), year, and title:

0

While Zotero has always renamed files automatically, Zotero 7 introduces a new, powerful syntax for customizing filenames. The default format can be customized from the General pane of the Zotero settings.

This is the default template string:

0

The following variables and parameters are supported:

##### Variables#

- Variable: 0; Description: Parent item’s principal creators (depending on the item type these are authors or artists but not editors or other contributors).
- Variable: 0; Description: The number of the parent item’s principal creators. *(Zotero 7.0.16)*
- Variable: 0; Description: Parent item’s editors.
- Variable: 0; Description: The number of parent item’s editors. *(Zotero 7.0.16)*
- Variable: 0; Description: All parent item’s creators.
- Variable: 0; Description: The total number of all parent item’s creators. *(Zotero 7.0.16)*
- Variable: 0; Description: Parent item’s creator (1–2 authors or editors), same as the value of the “Creator” column.
- Variable: 0; Description: Parent item type. Complete list of recognized item types can be found here
- Variable: 0; Description: The title of the attachment that is being renamed or created
- Variable: 0; Description: Year, extracted from parent item’s date field.
- Variable: 0; Description: Access date in UTC, or in a specified time zone if the 1 parameter is present. *(Zotero 8)*
- Variable: Any item field; Description: Complete list of fields can be found at the bottom of this page.


If a variable value starts or ends with a space, which is likely to happen when used in conjunction with the 0 parameter, these spaces are removed from the filename.

##### Parameters#

- Parameter: 0; Variables: All; Description: Truncates the variable value from the beginning. For example, 1 will be replaced with the value of the parent item’s title, omitting the first 5 characters. It can be combined with 2.
- Parameter: 0; Variables: All; Description: Truncates variable value at fixed number of characters, e.g, 1 will be replaced with the first 20 characters of parent item’s title. Truncation happens after every other parameter has been applied, except for 2, 3 and 4.
- Parameter: 0; Variables: All; Description: Prepends variable with given character(s), e.g., 1 will be replaced by the word “title” followed by parent item’s title. If variable is empty (e.g., item’s parent has empty title), entire statement, including prefix, is ignored.
- Parameter: 0; Variables: All; Description: Appends given character(s) at the end of a variable with, e.g., 1 will be replaced by parent item’s title followed by an exclamation mark. If variable is empty (e.g., item’s parent has empty title), entire statement, including suffix, is ignored.
- Parameter: 0; Variables: All; Description: Converts case of a variable, following values are accepted: 1, 2, 3, 4, 5, 6, 7 and, added in Zotero 7.0.16, 8. E.g., 9 will result in 10 in the file name.
- Parameter: 0; Variables: All; Description: Use a regular expression to replace a matching string in the variable with a value specified by the 1 parameter. For example, 2 is substituted with the parent item’s title where the first occurrence of the word “problem” is replaced with the word “solution”. It can be further configured with 3 parameter.
- Parameter: 0; Variables: All; Description: Defines a replacement value to use when matching using regular expressions (see 1). It is possible to specify capture groups defined in 2. For example, to prefix all occurrences of the words “dog” and “cat” with the word “super-” in the item’s title, the following template can be used: 3.
- Parameter: 0; Variables: All; Default Value: ‘i’; Description: Defines [flags](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_expressions#advanced_searching_with_flags) to use when matching using regular expressions (see 1). For example, 2 is substituted with the parent item’s title with all white space removed (without 3, only the first white space would be removed).
- Parameter: 0; Variables: All; Description: Use a regular expression to test for a matching string in the variable. This parameter is useful in conditions and it cannot be used with any other parameters, except for 1. For example, the following template will only return the parent item’s URL if the URL’s domain name is zotero.org: 2.
- Parameter: 0; Variables: 1, 2, 3; Description: Limits number of creators to use, e.g., 4 will be replaced with first editor of parent’s item.
- Parameter: 0; Variables: 1, 2, 3; Default Value: 4; Description: Customizes how creator name appears in the filename, with the following accepted options: 5 will use full name of the creator, beginning with family (last) name of the creator, 6 also uses full name but inverts the order, and options 7 and 8 will only use part of the parent item’s creator’s name.
- Parameter: 0; Variables: 1, 2, 3; Default Value: (single space character); Description: Defines what characters to use to separate given and family name, especially useful when combined with 4.
- Parameter: 0; Variables: 1, 2, 3; Default Value: 4; Description: Defines what characters to use to separate consecutive creators.
- Parameter: 0; Variables: 1, 2, 3; Description: Enables use of initials for part or whole of creators name, with the following accepted options: 4 will use initials for the entire name, 5 and 6 will only use initials for that part of the name. Order of name parts is controlled by 7 parameter and only parts included by 8 parameter can be converted to initials. E.g. 9 will be replaced by a comma-separated list of authors, where each author’s given name is replaced with an initial, followed by a dot and a space (e.g. 10).
- Parameter: 0; Variables: 1, 2, 3; Default Value: 4; Description: Controls what character is appended to the initial, if name part was initialized.
- Parameter: 0; Variables: 1; Description: Whether to use localized value of the item type variable, e.g., 2 will be replaced by parent item’s type spelled in the language Zotero is using.
- Parameter: 0; Variables: 1; Description: Converts the date to the specified time zone, e.g. 2 will be replaced by the parent item’s Access Date, converted to local time in New York. [List of time zones](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones). *(Zotero 8)*


##### Examples#

A year of publication, followed by a hyphen-separated list of authors, followed by a title truncated at 30 characters:

Template:

    {{ year suffix="-" }}
    {{ authors name="family-given" initialize="given" initialize="-" join="-" suffix="-" case="hyphen" }}
    {{ title truncate="30" case="hyphen" }}

Filename: 0

Anything not included inside a 0 bracket is copied to the filename literally:

Template:

    {{ itemType localize="true" }} from {{ year }} by {{ authors max="1" name="given-family" initialize="given" }}

Filename: 0

Templates also support conditionals. Certain part of the template can be included or excluded using a combination of 0, 1, and 2. The condition must end with 3. The template below will use the DOI for journal articles and preprints, the ISBN for books, and the title for any other item type:

    {{ if itemType == "book" }}
    {{ISBN}}
    {{ elseif itemType == "preprint" }}
    {{ DOI }}
    {{ elseif itemType == "journalArticle" }}
    {{ DOI }}
    {{ else }}
    {{ title }}
    {{ endif }}

As of Zotero 7.0.16, it’s possible to compare numeric values like 0 using relational operators such as 1, 2, 3, and 4. For example, the following template checks the number of authors: if there are two or more, it uses the first author’s name followed by et al.; if there are one or two authors, it includes all their names in the filename.

    {{ if {{ authorsCount > 2 }} }}
    {{ authors max="1" suffix=" et al" }}
    {{ else }}
    {{ authors join=" & " }}
    {{ endif }}

It’s possible to use regular expressions to match values and change the behavior of the template. For example, the following template preserves common attachment names (such as “Full Text”), but for attachments with non-matching titles, it uses the standard Zotero filename template:

    {{ if {{ attachmentTitle match="^(full.*|submitted.*|accepted.*)$" }} }}
    {{ attachmentTitle }}
    {{ else }}
    {{ firstCreator suffix=" - " }}{{ year suffix=" - " }}{{ title truncate="100" }}
    {{ endif }}

As of Zotero 8, it’s possible to include the access date and time of an item, converted to a local time in a specified time zone. Since “:” is not allowed in file names, we should replace it to avoid it being removed. Example below uses “-” as a replacement character.

    {{ accessDate timeZone="Europe/Berlin" replaceFrom=":" replaceTo="-" regexOpts="g" }}-{{ title truncate="100" }}

##### Complete List of Fields#

-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0
-   0

### Forum Guidelines {#forum-guidelines}

The Zotero Forums receive tens of thousands of posts per year. Zotero developers, and a number of Zotero experts who volunteer their time, read every post. To allow them to respond quickly, efficiently, and accurately, and to maintain a friendly and helpful atmosphere, please follow the guidelines outlined below.

#### Existing thread or new thread?#

In general, please start a new thread for your issue. It is much better to start a new thread rather than posting to a thread with a potentially unrelated issue. The Zotero developers or volunteers on the forums can point you to a relevant thread if one exists.

If you like, you can take a quick glance at the most recent posts in the forums and do a quick search to see if someone has posted about the same issue. Don’t spend more than a couple minutes searching — the forums receive a very high volume of posts, and people will be happy to point you to a relevant thread if one exists.

If you do find an existing thread, read the entire discussion — not just the thread title — to make sure it’s relevant, that a solution or workaround hasn’t already been provided, and that additional feedback is still necessary. You should also check the date of the last comment to make sure it’s not referring to a very old issue that’s no longer applicable.

If you’re not sure if your issue is exactly the same, it’s better to start a new thread, and someone will point you to an existing discussion if appropriate. (You can include links to other possibly relevant threads in your post.)

**Syncing issues and word processor plugin issues generally require individual troubleshooting and should almost always go in new threads**.

#### Don’t create redundant threads#

Don’t post on the same issue to more than one thread. Developers and many community members read every post, and posting in more than one place makes things harder both for the people who respond to you and for other people trying to find information on the same issue.

This also applies if you have an existing thread on an issue: don’t create a new one for the same problem, even if you’ve been waiting for a response. If you have additional related comments, just post them to the same thread.

If you want to change something in your original post, you can simply edit it.

#### Provide enough information#

See Reporting Problems for instructions for what to include in your post.

#### Etiquette#

We try our best to keep the atmosphere in the forums helpful and friendly. You can help us by observing the following advice:

-   **Post when you’re calm**. In many cases, the reason for posting to the forums is that you experienced some kind of frustration: something isn’t working, you are unable to figure out how to do something, or maybe you believe you’ve lost your work. Take some time to calm down before posting. The forums are not a place to vent frustration or anger.
-   **Respect others**. Don’t be rude or demand immediate attention. Zotero developers have many people to help, and Zotero community members are volunteering their time and expertise to help others.
-   **Don’t tag specific people for general questions.** Zotero developers read all threads, and community volunteers choose the questions they want to spend time on. Multiple people are able to answer most questions.
-   **AVOID ALL-CAPS**. Don’t write in all capital letters. It is interpreted by many as yelling and makes your post less readable. The same goes for using lots of bold or italics, or using multiple exclamation marks to stress your points.
-   **Keep an open mind.** Different people use Zotero in different ways. Not everybody shares your priorities, sometimes you’re the only person experiencing a particular problem, or maybe you just didn’t discover the best way to perform a task. Avoid generalizing statements like *“it’s obvious that without this feature, Zotero is useless for anybody”*.

Most importantly, remember that you and whoever responds to you typically share the same goal: resolving your issue as quickly as possible.

#### Be patient#

Zotero developers and many dedicated community members read every post.

While many forum posts are answered within minutes, others may not be answered for hours or days. In other cases — particularly for feature requests — developers may wait to see if other people chime in in support of an idea before responding, or there may just not be anything useful to add at this time.

If you don’t receive a response right away, give it time. If you’re having a serious issue and think your post may have slipped through the cracks, post a polite follow-up to the same thread asking if anyone has had a chance to look into your issue. Don’t create a new thread for the same issue, as it will likely just be deleted.

#### Be responsive#

After posting to the forums, be sure to follow up. You’ll automatically receive an email notification the first time someone replies to a thread you’ve posted to, but you won’t receive additional notifications until you visit the thread again, so it’s a good idea to click through to the thread to ensure that you’re notified the next time someone responds. (You can also subscribe to the forums using a feed reader.) If a suggested solution solves your issue, post a comment to the thread to let others know. If it doesn’t, confirm that you’ve tried what was suggested, and explain what happened when you did so.

Note that you may receive a short reply to your question, such as a simple “yes” or “no”, or a link to a previous relevant discussion, even one in which the issue hasn’t been resolved. Don’t interpret this as a dismissal of your question. The people lending user support deal with a large volume of posts and often don’t have time to write extensive replies. If you aren’t satisfied with an answer, ask for further elaboration. If you were given a link to another thread, respond in the new thread or follow it for updates.

If you resolve a problem on your own, post an update to the thread, and say what you did to fix it. Other people may be having the same issue and may benefit from your experience, or you may be able to provide the developers with additional information to fix the problem.

Go to the Zotero Forums

### Frequently Asked Questions {#frequently-asked-questions}

How do I back up my Zotero library? Where does Zotero store my references, notes and files?

How can I transfer my library to another computer?

How can I access my library from multiple computers?

How do I add an edited volume or a book chapter?

How do I see what collections an item is in?

Can I import existing bibliographies in Microsoft Word documents, PDFs, and other text files, into Zotero?

For additional topics, see the knowledge base.

***
If you can’t find the answer to your question here, see Getting Help.

### Getting Help {#getting-help}

Need help with Zotero? To resolve your issue as quickly as possible, please follow the steps below.

#### Step 1: Update Zotero#

Upgrade Zotero if you aren’t running the latest version, as your issue may have already been resolved in the most recent release.

#### Step 2: Check the troubleshooting pages#

If you have a general question about Zotero, check the Frequently Asked Questions and Knowledge Base.

If you’re experiencing a problem, the following dedicated troubleshooting pages may be helpful:

-   Issues installing Zotero
-   Issues with Zotero data (e.g., library missing, restoring from a backup)
-   Issues saving from websites
-   Issues with word processor plugins
-   Issues with data syncing
-   Issues with file syncing
-   Issues with file handling (e.g., PDFs opening in the wrong application)

#### Step 3: Post to the Zotero Forums (Really! It Works!)#

If you can’t find the answer to your question in the documentation, post a message to the Zotero Forums, where you’ll get fast, expert support directly from Zotero developers as well as long-time community members. See How Zotero Support Works for more information on how this allows us to provide the best possible support.

In general, please start a new thread for your issue. It is much better to start a new thread rather than posting to thread with a potentially unrelated issue. The Zotero developers or volunteers on the forums can point you to a relevant thread if one exists.

If you’re new to the Zotero Forums, please read the forum guidelines and bug reporting procedures before posting.

**Note:** You need to create a Zotero account and log in to post messages to the Zotero Forums. You can choose a different username for the forums from your account settings.

Go to the Zotero Forums

### Google Scholar (or some other site) locked me out after using Zotero to save items. What happened? {#google-scholar-or-some-other-site-locked-me-out-after-using-zotero-to-save-items}
#### Google Scholar (or some other site) locked me out after using Zotero to save items. What happened?#

When saving multiple items, Zotero users will occasionally push the limits of what Google Scholar, or any other Zotero-compatible database (e.g., ProQuest), wants an individual to access in a given period of time. When this happens, saving to Zotero may fail, and when browsing the site you may start seeing CAPTCHAs or messages about making automated requests.

On Google Scholar, you can save individual items by simply clicking through to the publisher site (which will often result in better metadata anyway).

If you need to import more than a few items at once, it’s better to first save the items to Google Scholar’s [“My Library”](https://scholar.google.com/intl/en/scholar/help.html#library) feature, export items en masse as a BibTeX file from the My Library view, and then import the BibTeX file into Zotero. Other sites may have similar features that allow exporting a collection of items. You can also simply click through to the article page and save from there rather than saving from the search results page.

The [Google Scholar Citations for Zotero](https://github.com/beloglazov/zotero-scholar-citations) plugin can also send sufficient numbers of queries to cause Google Scholar to lock out Zotero.

#### How can I restore my access to the site?#

Sometimes a site will request a CAPTCHA verification. In this case, complete the verification test to restore access. Otherwise, in general, the site will restore your access within a few hours. You may be able to access the site using a different web browser in the interim.

If you’re still able to access the site in your browser but not able to save to Zotero, as a temporary solution you may be able to save to BibTeX or RIS, as explained above, and import the file into Zotero.

### How can I open multiple instances of Zotero and save to or cite from a specific instance? {#how-can-i-open-multiple-instances-of-zotero-and-save-to-or-cite-from-a-specific-}

You can start multiple instances of Zotero at the same time as long as they’re pointing to different profile and data directories. On Windows and Linux, the [-no-remote command-line flag](http://kb.mozillazine.org/Opening_a_new_instance_of_Firefox_with_another_profile) may also be required.

Additional configuration may be required for browser and word processor integration, as described below.

Note that using multiple instances of Zotero will use more memory, so depending on how much RAM you have in your computer, using multiple instances at the same time may be slower than using a single instance with more data.

#### Zotero Connector#

By default, when running multiple copies of Zotero (either by using multiple profiles or running Zotero in separate OS accounts), the Zotero Connector will save items to the first opened instance. To point a specific Zotero Connector installation to a specific Zotero profile, you can set the 0 hidden pref in Zotero and the 1 pref in the connector. Incrementing the port number by 1 for one pair (Zotero profile + connector installation) is sufficient.

#### Word Processor Integration#
##### Word#
##### Word for Mac (Zotero 6) / Word for Windows#

The Word plugin should automatically connect to the correct Zotero instance on a multi-user system. It’s not possible to point the Word plugin at a specific instance of Zotero within the same OS user account.

##### Word for Mac (Zotero 7)#

Open the user’s Word Startup folder and create a 0 file containing the port number you’ve configured in 1.

##### Google Docs#

Google Docs integration relies on the Zotero Connector being configured to connect to a specific Zotero instance, as described above.

##### LibreOffice#

The LibreOffice plugin uses a fixed HTTP port to connect to Zotero and cannot currently be configured, so it will connect to the first opened instance.

### How can I switch the Zotero desktop app to a different Zotero account? {#how-can-i-switch-the-zotero-desktop-app-to-a-different-zotero-account}

A Zotero database can only be synced with a single Zotero account.

Once you set up syncing for a given database in the Sync pane of the Zotero preferences, it’s no longer possible to sync that database with another account. If you unlink your account and attempt to set up syncing with a different account, Zotero will warn you that all local data will be removed if you continue. This prevents data from different accounts from being inadvertently combined and avoids difficult-to-resolve conflicts between data created in multiple accounts.

-   If you’re simply trying to change usernames, you can do so from your profile settings rather than switching to a completely new account. (If you’ve already created a new account with the desired username, you’ll need to delete it or change its username before you can use that username for your original account.)
-   If you wish to use libraries from multiple accounts on the same computer, you can set up multiple Zotero profiles pointing to separate data directories and sync each one with a different account.
-   If you’ve created data in multiple accounts and want to merge the data into a single account, you can export the data in one account to Zotero RDF with files — either on another computer or from a separate profile that you later delete — and import it into the other account. Note that using export and import will reset Date Added/Modified times and will break links to items from citations in existing word processor documents, so you may wish to export from the account with less data or data that you’ve interacted with less.

### How can I use Zotero with a SOCKS proxy? {#how-can-i-use-zotero-with-a-socks-proxy}
#### How can I use Zotero with a SOCKS proxy?#

Zotero can be configured to use a SOCKS proxy. This is useful when you have set up a local port forwarding to your institution’s servers. If you only configure your browser to use the proxy, Zotero Connector may fail to download the full-text files of references you save because it will not download them via your institution’s servers.

First, open the Config Editor, which is located under the Advanced tab of the Preferences window.

Scroll down until you find the **network.proxy…** options or use the search box.

Type the IP of the proxy server in the **network.proxy.socks** preference. If the proxy server is on your machine, use 0.

Likewise, set **network.proxy.socks\_port** to be the port of the proxy server.

**network.proxy.socks\_remote\_dns** specifies whether to use the DNS of the proxy server. If you are not sure, it is best to change it to 0.

Finally, set **network.proxy.type** to the value 0. This indicates that you have manually configured the proxy settings.

Zotero will now redirect all requests to the proxy server. This means that when you ask the Zotero Connector add-on in the browser to download a file or web page, Zotero will download it using the proxy you just specified in the settings.

### How can subsequent occurences of the same author replaced by a fixed term/symbol? {#how-can-subsequent-occurences-of-the-same-author-replaced-by-a-fixed-termsymbol}
#### How can subsequent occurences of the same author replaced by a fixed term/symbol?#

Some citation styles require to replace any subsequent occurences of the same author/bookauthor/editor by [idem/eadem](http://en.wikipedia.org/wiki/Idem), their abbreviations id./ead., or some other term or symbol like a dash.

a\) It may be that you need such replacements *within the bibliography*, e.g.

     Doe, Johnson & Williams. 2001.
     --- & Smith. 2002.
     Doe, Stevens & Miller. 2003.
     ---, --- & ---. 2004.

This is **possible** and just depends on the citation style you use. (If you want to use it in your own citation style, see [subsequent-author-subsitute rules in CSL](https://docs.citationstyles.org/en/stable/specification.html#reference-grouping).)

b\) It may be that you need such replacements *within one reference*, for example for a work in an anthology (collected or selected work), e.g.

     DOE, JOHN (1998): Article. In: IDEM (ed).: Book.

This is currently not possible, but as a **workaround** you can enter here in Zotero “IDEM” as the book author or editor, which will give you the desired result.

c\) It may also be that you need such *replacements for subsequent citations* of the same author but different books, e.g.

     V. HUGO, Les Misérables, Livre de poche, 2002, p. 127.
     ID., La Légende des siècles, Hatier, 1998, p. 45.

This is currently not possible, but as a **workaround** you can suppress the author here and add “ID., ” manually.

Main thread for this issue: 0

### How do I add a letter or memo? {#how-do-i-add-a-letter-or-memo}
#### How do I add a letter or memo?#

Click the green new item button and select “Letter”. Use Author for the sender of the letter. To add a recipient click the + sign on the author line in the right column. This will create an additional Author line. If you click the triangle to the left of the new author field you will be able to change Author to Recipient. Enter the type of letter (memo, telegram, etc.) in the “Type” field. On entering archival information, see How do I add an archival or other unpublished source?

If you have a scan of the letter (e.g., a PDF or an image), you could add it as an attachment to the Letter item.

### How do I add an archival or other unpublished source? {#how-do-i-add-an-archival-or-other-unpublished-source}
#### How do I add an archival or other unpublished source?#

Click the green new item button and select “Manuscript”. Enter the type of source (press release, draft, field notes, lecture notes, etc.) in the “Type” field. Use “Loc. in Archive” field to enter location of the document, i.e. box, folder, record group, etc. Use the “Archive” field to enter the name of the archive or collection. For adding Letters, see How do I add a letter or memo?.

### How do I add an edited volume or a book chapter? {#how-do-i-add-an-edited-volume-or-a-book-chapter}
#### How do I add an edited volume or a book chapter?#

A book chapter from an edited volume is entered as a “Book Section”. Click the green new item button and select “Book Section”. You will now see both a “Title” field for the chapter title and a separate “Book Title” field. To add an editor click the + sign on the author line in the right column. This will create an additional Author line. If you click the “Author” label you will be able to change it to an editor.

### How do I attach a file or web page to an item? {#how-do-i-attach-a-file-or-web-page-to-an-item}
#### How do I attach a file or web page to an item?#

After adding an item to your library, click on the Attachments tab in the right-hand section of the Zotero pane. A pull-down menu will allow you to add a link to a file stored elsewhere on your hard disk, copy a file to your Zotero storage directory, add a link to the current web page, or save a snapshot of the current web page.

You can also drag the item from your computer’s file system (hold down Shift \[Mac\] or Ctrl \[Windows/Linux\] while dragging to move, rather than copy the file) or right-click on the Zotero item and choose “Add Attachment”.

### How do I change the font size of text in Zotero? {#how-do-i-change-the-font-size-of-text-in-zotero}
#### How do I change the font size of text in Zotero?#

You can change the font size for the user interface and for notes from the View menu.

### How do I change the size of my Zotero pane? {#how-do-i-change-the-size-of-my-zotero-pane}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### How do I change the size of my Zotero pane?#

You can toggle between a full screen and a smaller screen by clicking on the toggle icon on the right side of the Zotero pane. You can also click and drag your Zotero pane to resize it. You can place your Zotero pane either at the bottom of your page or at the top by clicking on the Zotero gear icon and selecting Preferences.

### How do I find the browser console? {#how-do-i-find-the-browser-console}

The browser console is often used to identify problems with a webpage, and Zotero developers may ask you to check your browser console for errors to help troubleshoot a problem with the web library or ZoteroBib.

You can open your browser’s console using the following key combination:

-   Chrome and Edge:
    -   MacOS: Cmd + Option + J
    -   Windows/Linux: Ctrl + Shift + J
-   Firefox:
    -   MacOS: Cmd + Option + K
    -   Windows/Linux: Ctrl + Shift + K
-   Safari on MacOS:
    -   Cmd + Option + C

### How do I label different creator roles, such as Director or Producer, for films and other media? {#how-do-i-label-different-creator-roles-such-as-director-or-producer-for-films-an}
#### How do I label different creator roles, such as Director or Producer, for films and other media?#

For citations to films, recordings, and broadcasts, Zotero currently has limited to support for labeling producers, scriptwriters, and some other Creator roles.

To label directors, leave the main Zotero field blank (or enter the names as Contributor) and enter the director names using 0 in Extra. See Citing Fields from Extra below.

All Creator roles (Director, Producer, Scriptwriter, etc.) can also be labelled by entering the names using the default 0 role for the item (*Performer* for Audio Recording, *Podcaster* for Podcast, and *Director* for Film, Radio Broadcast, TV Broadcast, and Video Recording) and adding the appropriate labels in parentheses after the authors’ first names—e.g., MacNaughton \|\| Ian (Producer). Note that the labels will be rendered verbatim in citations; enter abbreviated terms (e.g., “Prod.”) here as needed.

If the style uses initials for author first/given names, rather than full names (e.g., APA style), if the label contains multiple words (e.g., “Executive Producer” or “Writer & Director”), Zotero will abbreviate the words of the label after the first. To avoid this, type a “Word Joiner” character (Unicode U+2060, printed here between quotes: “”) on either side of each space in the label.

See also Role Labels for Media Creators

### How do I open and close my Zotero pane? {#how-do-i-open-and-close-my-zotero-pane}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### How do I open and close my Zotero pane?#

Click on the Zotero icon at the bottom of your Mozilla Firefox browser window to open the pane. You can add a Zotero icon to your toolbar at the top of your browser window by clicking on View → Toolbars → Customize. You can close the pane by clicking the Zotero icon (either in toolbar or on bottom of your browser window or by clicking on the X icon on the right. You can also open or close pane from the keyboard: Ctrl-Shift-Z on Windows/Linux or Shift-Command-Z on Mac OS X).

### How does Zotero parse things in the name fields? {#how-does-zotero-parse-things-in-the-name-fields}
#### How does Zotero parse things in the name fields?#

There are actually three parts to the story of names in Zotero (“creators”, in techie lingo):

1.  Creator types;
2.  Field mode; and
3.  Name-part parsing.

Each of these topics is covered below. The first two are very simple.

#### Creator types#

Each name field has a label to its left, which is actually a button. Clicking on it will open a list of possible \*creator types\* for the current item type. You can change the type of an individual creator by clicking on its label and selecting from the list.

#### Field mode#

There is a small square icon to the right of each name (just before the **(+)** and **(-)** buttons used to add and remove creators). Clicking on the square icon will toggle the name between *single-field mode* and *two-field mode*.

-   In single-field mode, the field content is not parsed when generating citations. \[1\] This mode is ordinarily used for institutional names.
-   In two-field mode, the field is parsed to (even) smaller parts when generating citations. Two-field mode should ordinarily be used for personal names. *This includes Asian names!* The CSL processor in Zotero can correctly format names in a variety of languages, \[2\] and across all citation styles; but this flexibility requires correctly entered data. It is not a good practice to “force” a particular form by selecting single-field mode unnecessarily.

#### Name-part parsing#

In two-field mode only, personal names are parsed into five separate parts for formatting purposes. Here they are, with a brief explanation of each:

-   *Family name:* The family or clan name of an individual is the primary “family name” in Zotero: \[3\]
    -   The family name of “Sam Spade” is “Spade”.
    -   The family name of “Jeremy Atticus Finch” is “Finch”.
    -   The family name of “Kuruma Torajirō” is “Kuruma” (note that the family name part of this Japanese character’s name is written first).
-   *Given name:* This refers to an individual’s “own” name, or names:
    -   The given name of “Sam Spade” is “Sam”.
    -   The given names of “Jeremy Atticus Finch” are “Jeremy Atticus”.
    -   The given name of “Kuruma Torajirō” is “Torajirō” (the lead protagonist of the Japanese “Tora-san” series).
-   *Dropping particle:* Dropping particles, a feature of some European names, are descriptive elements that are placed between the given and the family name when written in “normal” order. A dropping particle is never placed with the family name when written in “sort order”.
    -   In “Ludwig van Beethoven”, “van” is a dropping particle.
    -   In “Jean de La Fontaine”, “de” is a dropping particle.
-   *Non-dropping particle:*
-   Articular

\[1\] In the Juris-M (formerly Multilingual Zotero/MLZ) variant of official Zotero, single-field names are parsed into subunits by splitting the field on pipe (“**\|**”) characters. In official Zotero, the field is printed exactly as written.

\[2\] Chinese and Japanese names will render correctly in official Zotero. Names in some languages (Khmer and Myanmar being two examples) are not yet handled correctly by official Zotero; users with special requirements may wish to explore Juris-M, which is able to apply precise name formatting rules across all language domains.

\[3\] In some other countries, individuals have no family or clan name, but only given names. Formatting conventions in such countries vary. In Myanmar and Cambodia, the entire set of names is always written in formal contexts (including citations). In Mongolia, it is customary to handle the bare patronymic in the same way as a “family” name. Where names with special requirements must be handled frequently, Juris-M may be worth a look.

### How Zotero Support Works {#how-zotero-support-works}

We know people use Zotero for critical, time-sensitive projects, and we work hard to provide fast, expert help, no matter the time of day or day of the week. Support for Zotero works a bit differently than you might be used to, and we’re incredibly proud of it — we think we’re able to provide a support experience that surpasses what you’ll get for nearly any other piece of software. It’s one of the best reasons to rely on Zotero for your research.

All Zotero support is provided in public through the Zotero Forums. Unlike with many other software forums, where users are often left to fend for themselves, Zotero developers read every forum thread, and you’ll frequently find yourself talking within hours or minutes with the person who wrote the specific part of the software you have a question about. There’s no wasting time explaining your problem to a chatbot or a customer support representative following a script. You’ll never need to escalate your report to a faceless engineering department, or wait for the right person to provide follow-up questions to pass back to you. The Zotero team is made up of conscientious people who care deeply about people’s experiences using the tools they build, and they’ll work one-on-one with you to resolve your issue.

In addition to the core Zotero developers, you’ll find a dedicated, expert community, with deep knowledge of Zotero and many related subjects: citation styles, metadata standards, Zotero plugins, workflows, etc. Community members can answer a wide range of questions, including many that Zotero developers themselves can’t answer. They can also get you assistance faster by pointing you to relevant previous posts, making sure you’ve provided enough information for a developer to be able to help you if necessary, translating troubleshooting information into other languages, or suggesting an immediate workaround until your issue is fully resolved.

Beyond technical support, the forums allow the entire community — Zotero developers and users alike — to help shape Zotero’s future. Many of Zotero’s features began with discussions in the forums, and many changes are the direct result of feedback from users there, so we strongly encourage you to get involved.

Note: While the public nature of the forums are a critical part of what makes them so effective, if you don’t want to post under your own name, you can choose a different username for the forums from your account settings. You’ll also never need to post any private information: if we need any private information to solve your problem, we’ll ask you to send that via email with a link to the relevant forum thread.

If you decide you still need in-person help, most university libraries offer Zotero instruction and support, or it might be enough to ask a technically inclined friend or colleague to read through the detailed documentation we provide on many topics.

Go to the Zotero Forums

### I have two Zotero libraries. How can I combine them? {#i-have-two-zotero-libraries-how-can-i-combine-them}
#### I have two Zotero libraries. How can I combine them?#

You will want to export one library as Zotero RDF and import the .rdf file into the second library. To do this, click “File” -> “Export Library…”. The default export format should be “Zotero RDF”, which you should leave selected. You can check the boxes for including files and notes if you would like to transfer those to the other library. Click OK, name the .rdf file, and choose a destination to save it. You can then e-mail this folder to a friend or colleague or to yourself. From the second Zotero library, click the “File” menu again and select “Import…”. Choose the .rdf file in the folder you created and click OK. This should import the items, combining the two libraries.

### Journal Abbreviations {#journal-abbreviations}

Many citation styles require that journal titles be abbreviated (e.g., using a journal abbreviation list).

Zotero 4.0 and later can automatically abbreviate journal titles in Index Medicus/MEDLINE format. This option is available in the Document Preferences window in the Zotero word processor plugin. When it is disabled, Zotero uses the abbreviation from the “Journal Abbr.” field in the client.

A planned feature for Zotero is support for choosing among several abbreviation lists, so that different styles can use different journal abbreviations (e.g., *Proceedings of the National Academy of Sciences* can be abbreviated as either *Proc. Natl. Acad. Sci. U. S. A.* or *PNAS*).

### Knowledge Base {#knowledge-base}
#### Basic Zotero Usage and Troubleshooting#

-   Can I still use Zotero if I can’t install programs on my computer?
-   Does Zotero offer installable packages of Zotero for specific Linux distributions?
-   Does Zotero support non-Western characters?
-   File Handling Issues
-   How can I quickly switch between Zotero and my browser, PDF viewer, and/or word processor?
-   How do I change the font size of text in Zotero?
-   How do I change the language of Zotero’s user interface?
-   How do I change Zotero settings?
-   How Do I Install Zotero on a Chromebook?
-   How do I uninstall Zotero?
-   I upgraded to Zotero 5.0 and now my data is missing! How do I get it back?
-   Is the Zotero web library the same as the Zotero desktop app?
-   Library Item Overview
-   What do I do if my Zotero database is corrupted?
-   What is the difference between Links and Snapshots?
-   What version of Zotero do I have?
-   Why am I getting a database version error?
-   Why can’t I access a proxied site when the Zotero Connector is enabled?
-   Why don’t I see any tabs in the item pane after selecting an item?
-   Why is my browser saying the Zotero Connector needs access to my data on all websites?
-   Why is there no save button in my browser toolbar?
-   Why isn’t Zotero detecting updates?
-   Zotero Connector and Safari
-   Zotero Connector: “Is Zotero Running?”
-   Zotero Keyboard Shortcuts

#### Getting Stuff into your Library#

-   A snapshot only captures the first of multiple pages of a New York Times article. How do I get Zotero to capture the full article?
-   Default translators
-   Google Scholar (or some other site) locked me out after using Zotero to save items. What happened?
-   How can I edit item metadata?
-   How can I import from Citavi?
-   How do I add a letter or memo?
-   How do I add an archival or other unpublished source?
-   How do I add an edited volume or a book chapter?
-   How do I attach a file or web page to an item?
-   How do I automatically add a book or article to Zotero?
-   How do I import a Mendeley library into Zotero?
-   How do I import BibTeX or other standardized formats?
-   How do I import from EndNote?
-   How do I import references into Zotero?
-   How do I label different creator roles, such as Director or Producer, for films and other media?
-   How do I manually add a bibliographic item?
-   How do I save my work?
-   How does the Import from Clipboard feature work?
-   I have bibliographies in Microsoft Word documents, PDFs, and other text files. Can I import them into my Zotero library?
-   Legal Citations: Juris-M
-   My snapshot is unreadable because flash ads appear on top of article text. How do I get a better snapshot?
-   Sometimes the icon for a book/article/etc doesn’t appear in the address bar right away. What’s going on?
-   Troubleshooting Problems Saving to Zotero
-   Why doesn’t the Zotero Connector offer to save complete data from a webpage?
-   Why doesn’t Zotero have a “watch folder” feature?
-   Why don’t I see a Zotero icon in the address bar when viewing a webpage?
-   Zotero Item Types and Fields

#### Organizing your Library#

-   Can I change the way items are sorted in my library?
-   Can I highlight and annotate PDFs with Zotero?
-   How can I see how many items I have in my Zotero library?
-   How can I see what collections my item is in?
-   How do I merge tags?
-   How do I organize my notes into an outline?
-   How do I sort my notes by page number?
-   I have two Zotero libraries. How can I combine them?
-   What ways can I organize and manage my collections?

#### Syncing#

-   “\[domain\] uses an invalid security certificate. The certificate is not trusted because \[…\]“
-   “Error connecting to server. Check your Internet connection.”
-   Do annotations sync?
-   How can I access my library from multiple computers? Can I store my Zotero library and associated files on an external drive?
-   How can I switch the Zotero desktop app to a different Zotero account?
-   List of WebDAV services
-   Why am I getting “The attached file could not be found” when I try to open a file in Zotero?
-   Why aren’t changes I make syncing between multiple devices and/or zotero.org?
-   Why can’t I connect to a WebDAV server via HTTP on iOS or Android?
-   Why do I keep getting file sync errors while syncing?
-   Why does Zotero keep asking me to reconcile the same conflicts whenever I sync?
-   Why is Zotero telling me that some data could not be downloaded?
-   Zotero Sync Reset Options

#### PDF Reader and Note Editor#

-   Where did the Extract Annotations button go?
-   Why do I see highlighted text twice in the PDF reader or in notes created from annotations?

#### Word Processor Plugins#

-   Can I prevent the “Add Citation” dialog from the word processor plugins from moving behind the word processor window?
-   How do I uninstall the Zotero word processor plugins?
-   How do I unlink all Zotero citations in a document?
-   What happened to the “classic” citation dialog?
-   Where is the Zotero toolbar in Word for Mac 2008?
-   Why are my citations underlined with a dashed line?
-   Why are Zotero citations or bibliographies always highlighted in gray or another color?
-   Why do I see code beginning with ADDIN ZOTERO\_ITEM CSL\_CITATION in my document instead of formatted citations?
-   Why is a citation not updated in my document after editing the item in Zotero?
-   Why is Zotero slow to insert citations or update the bibliography?
-   Why isn’t Zotero detecting my existing citations?
-   Zotero does not have permission to control Word

#### Citation Formatting#

-   Can I use Zotero in one language and create bibliographies in another?
-   Does Zotero support label/authorship trigraph styles, like \[ddb98\]?
-   DOI format in APA style
-   How can subsequent occurences of the same author replaced by a fixed term/symbol?
-   How do I get titles to show up in sentence case in bibliographies?
-   How do I prevent title casing of non-English titles in bibliographies?
-   How do I use rich text formatting, like italics and sub/superscript, in titles?
-   How do you cite a secondary source in Zotero?
-   How does Zotero parse things in the name fields?
-   I need to use Chicago style. Which of the three versions that come with Zotero should I use?
-   I’m the publisher/editor of a journal. What can I do to have Zotero support our style?
-   Journal Abbreviations
-   Missing Italics (or Italics-Only) in Word Bibliographies
-   References appear in the wrong font in Word/LibreOffice
-   Standard Citation Styles
-   What are these DOIs doing in my bibliography?
-   What is the official Harvard style?
-   Why do some citations include first names or initials?
-   Why isn’t the first letter of a subtitle in uppercase in bibliographies?

#### Other Topics#

-   “\[domain\] uses an invalid security certificate.”
-   How can I move my Zotero library to a different computer?
-   How can I open multiple instances of Zotero and save to or cite from a specific instance?
-   How can I use multiple profiles in Zotero?
-   How can I use Zotero with a SOCKS proxy?
-   How do I export my Zotero library?
-   How do I find the browser console?
-   How do I get my Zotero collection to work with an Exhibit or Citeline presentation?
-   How do I print my Zotero references and notes?
-   How do I turn off automatic case changes/capitalization during item import?
-   Using Zotero to create COinS metadata
-   What connections do I need to allow through a firewall for Zotero to work properly?
-   What does “Zotero” mean?
-   Where did the Extract Annotations button go?
-   Why am I getting a “disk I/O error” at Zotero startup?
-   Why can’t Zotero find a linked file?
-   Why can’t Zotero find a stored file?
-   Why do I no longer see a Save to Zotero option in the Firefox open/save dialog?
-   Why do I see highlighted text twice in the PDF reader or in notes created from annotations?
-   Why does Zotero store PDF annotations in its database instead of in the PDF file?
-   Why is Zotero still saying that my storage is full after I upgraded my storage plan or deleted files?
-   Zotero and the Text Encoding Initiative (TEI)
-   Zotero and Wikipedia/ Wikidata

#### Zotero for Firefox (pre-Zotero 5.0)#

These topics are no longer relevant for current versions of Zotero.

-   “The add-on could not be downloaded because of a connection failure on www.zotero.org.”
-   Can I open snapshots in a new tab or window? (Zotero for Firefox)
-   How can I use the Chrome or Safari Connectors with Zotero for Firefox?
-   How do I change the size of my Zotero pane?
-   How do I open and close my Zotero pane?
-   Sharing data directory between Zotero Standalone and Zotero for Firefox
-   Where are the link and snapshot buttons in Zotero 2.0?
-   Why am I getting an “unresponsive script warning”?
-   Zotero doesn’t open when I click the Zotero status bar icon or select “Zotero” from the Firefox Tools menu.

### Language Support {#language-support}

Zotero’s [Unicode](http://en.wikipedia.org/wiki/Unicode) support allows you to import, store, and cite items in any language. You can change the language of both the Zotero user interface and the citations and bibliographies created by Zotero. Finally, there is an unofficial multilingual version of Zotero, which supports storage of item metadata in more than one language (transliterations and translations).

#### Switching Languages#
##### Zotero#

In Zotero, the interface language defaults to matching the operating system’s language. To use a different language, go to the Edit menu (Windows/Linux) or Zotero menu (Mac) and select Preferences, click on the Advanced tab, and make your selection from the Language drop-down.

##### Citations and Bibliographies#

To keep your Zotero UI in one language, but use another language for the citations and bibliographies created by Zotero, simply select the citation language you’d like to use from the appropriate location:

-   The “Create Bibliography from Selected Item(s)” dialog
-   The word processor integration plugin’s document preferences window
-   The Quick Copy section in the Export pane of the Zotero preferences

#### Contributing Translations#

You can report mistakes in Zotero’s translations in the Zotero forums. If you would like to make larger contributions (like translating the Zotero client into an as of yet unsupported language), see the developer’s instructions for localization.

#### Juris-M: Multilingual fields, Translations, and Transliterations#

Juris-M (formerly called Multilingual Zotero or MLZ) is an unofficial community-driven version of Zotero that adds additional support for multilingual and legal citations. Juris-M allows you to store transliterations and translations of names, titles and other fields, and create citations and bibliographies that show this information (e.g. “Soseki, Wagahai ha neko de aru \[I am a cat\] (1905-06)”).

Juris-M is developed by Frank Bennett, a Zotero user and active Zotero contributor. If you would like to try out Juris-M, see the [project webpage](https://juris-m.github.io/).

### Licensing {#licensing}
#### Source code#

Unless otherwise indicated, the source code of the Zotero project is released under the [GNU Affero General Public License (version 3)](http://www.gnu.org/licenses/agpl.html).

#### Wiki content#

Zotero wiki content produced after April 26, 2015, is released under a [Creative Commons Attribution-ShareAlike 4.0 International License (CC-BY-SA 4.0)](http://creativecommons.org/licenses/by-sa/4.0/).

Previous wiki content may be licensed under a [Creative Commons Attribution-NonCommercial-ShareAlike 3.0 Unported License (CC BY-NC-SA 3.0)](http://creativecommons.org/licenses/by-nc-sa/3.0/). If you are interested in using content for commercial purposes, please check the revision history against the Wiki Contributor License Agreement.

### Locate Menu {#locate-menu}

The Locate menu offers several ways to access files in your library and to look up items online. The menu can be opened by clicking the straight arrow button () at the top-left of the right-hand column of the Zotero pane.

Which menu entries are available depends on the type of items you have selected in the middle column. The possible options are:

-   **View File/PDF/Snapshot** - open the files/PDFs/snapshots of the items
-   **View Online** - look up the items online, using its URL, DOI, or child link’s URL
-   **Show File** - locates the files/PDFs of the items on your computer
-   **Library Lookup** - looks up the items in your library of choice using OpenURL
-   **CrossRef Lookup** - looks up and resolves the DOI of the items
-   **Manage Lookup Engines…** — See \#Managing Lookup Engines

#### Library Lookup#

The Library Lookup option lets you locate items in an online library catalog so you can track down a physical or online full-text copy of the resource. You’ll need to select your library’s OpenURL resolver from the Advanced tab of the Zotero preferences (or the General tab in Zotero 7).

If your library’s resolver isn’t in Zotero’s OpenURL resolver directory, you can enter the address manually. Most university libraries provide their OpenURL resolver address on their websites.

#### Managing Lookup Engines#

The “Manage Lookup Engines…” option opens the Article Lookup Engine Manager window, where you can enable/disable lookup engines, remove installed lookup engines, or reset your installed lookup engines to the defaults.

It is not currently possible to add new lookup engines using through either the Zotero desktop client or the browser Connector extensions. To add new lookup engines, you must edit the 0 file located in the 1 folder in your Zotero data directory. Lookup engines can be added to this file using JSON syntax specifying an OpenSearch lookup engine. A more user-friendly way to add lookup engines will be added in a future version.

An 0 file with a variety of useful lookup engines is available from 1. To use this file, download it and place it in the 2 folder in your Zotero data directory (replacing the version of 3 that is already present). You can remove unwanted engines from the “Manage Lookup Engines…” in the Zotero Locate menu.

#### Public Library Lookup Engines#

A repository of thousands of local public library Lookup Engines (as well as engines for some universities and commercial databases) is available from 0.

### My Publications {#my-publications}

The My Publications feature allows you to automatically create a bibliography of your research and share copies of your work on zotero.org. For an example, see 0

To share an item with My Publications, drag it to the “My Publications” special collection under “My Library” in the Zotero desktop client app.

#### Sharing Items, Files, Notes, and Links#

When you drag an item to the My Publications collection, the My Publications helper window will appear. From this window, you can choose whether to also share any notes or stored attachment files attached to the items. Links (e.g., to a copy of the article hosted on your personal website, an institutional or disciplinary repository, or the publisher website) will always be shared. If an item doesn’t have an attached file or link, the bibliography shown on zotero.org will include a link drawn from the item’s DOI or URL fields.

After choosing whether to share files and notes, you will next be asked to confirm that you are a creator of the work and (if applicable) that you have the rights to publicly distribute the included files. You should only use My Publications to share your own personal work (to share items more generally, use Zotero Groups.

You should only share files for which you have public distribution rights. If your work is published, check your contributor agreements from your publisher (see also the [SHERPA/RoMEO archive](http://www.sherpa.ac.uk/romeo/index.php) of publisher copyright and self-archiving policies). In many cases, it is acceptable to share a pre-print or author manuscript version of your work, but you should verify before sharing with My Publications.

#### Distribution Rights#

If you are sharing files, you will next be asked whether you want to permit your work to be shared further. Your choices are:

-   **All rights reserved:** Retain all distribution rights over your work. Only publish your work on zotero.org and don’t permit others to share files further.
-   **Creative Commons licenses:** Share your work under a [Creative Commons license](https://creativecommons.org/). These licenses permit others to share your work so long as they attribute it to you, provide a link to the license, and indicate any changes made, potentially with limitations on modification and commercial uses. You will choose a specific license in the next window.
-   **Public domain:** Release your work into the [public domain](https://creativecommons.org/publicdomain/zero/1.0/). This waives all of your rights to the work under copyright law. Please read the Creative Commons [CC0 FAQ](https://wiki.creativecommons.org/wiki/CC0_FAQ) before placing your work in the public domain and note that this action is irreversible, even if you later choose different terms or cease publishing the work.

#### Choose a License#

Finally, if you choose to distribute your work under a [Creative Commons license](https://creativecommons.org/), you must choose the terms of the license to apply.

First, decide whether to allow adaptations or modifications to your work to be shared:

-   No
-   Yes, as long as others share alike (apply the same Creative Commons license to any derivative works)
-   Yes

Second, decide whether to permit commercial uses of your work.

Based on these questions, Zotero will give a link to the appropriate Creative Commons license. This license will appear with the files you share with My Publications.

For additional information on Creative Commons licenses, see [Considerations for licensors](https://wiki.creativecommons.org/wiki/Considerations_for_licensors_and_licensees). Note that rights waived under a Creative Commons license cannot be revoked, even if you later choose different terms or cease publishing the work.

### My snapshot is unreadable because flash ads appear on top of article text. How do I get a better snapshot? {#my-snapshot-is-unreadable-because-flash-ads-appear-on-top-of-article-text-how-do}
#### My snapshot is unreadable because flash ads appear on top of article text. How do I get a better snapshot?#

Look for a printer friendly version, then save the snapshot. Print versions of articles usually only have one ad on top and their snapshots are easier to read.

### A snapshot only captures the first of multiple pages of a New York Times article. How do I get Zotero to capture the full article? {#a-snapshot-only-captures-the-first-of-multiple-pages-of-a-new-york-times-article}
#### A snapshot only captures the first of multiple pages of a New York Times article. How do I get Zotero to capture the full article?#

The web snapshot feature only captures the current page. Many sites, like the New York Times, will only display part of the document you are interested in on the first page. A general rule for working around this issue is to look for a printer friendly version. This will generate a page that has all of the document available on one page. You can now capture the full text by clicking on the newspaper icon that appears in your location bar. The New York Times specifically also has a “Single Page” button. There is even a [greasemonkey](https://addons.mozilla.org/firefox/addon/748) script that [automatically displays](http://userscripts.org/scripts/show/56690) the single page view.

### Sometimes the icon for a book/article/etc doesn’t appear in the address bar right away. What’s going on? {#sometimes-the-icon-for-a-bookarticleetc-doesnt-appear-in-the-address-bar-right-a}
#### Sometimes the icon for a book/article/etc doesn’t appear in the address bar right away. What’s going on?#

Zotero must wait for the entire page to load before it can check for an item on the page. Sometimes there are elements of the page (a large image, e.g.) that take time to load and thus delay Zotero’s activity.

If the page has finished loading, but you still don’t see the proper item type icon, see here.

### Tips and Tricks {#tips-and-tricks}

-   When you have selected an item in the middle column, you can highlight all collections that contain this item by holding down the “Option” key on Mac OS X, the “Control” key on Windows, or the “Alt” key on Linux.
-   Press “+” (plus) on the keyboard within the collections list or items list to expand all nodes and “-” (minus) to collapse them.
-   To see the number of items in the selected library or collection, click an item in the middle column and use the *Select All* shortcut (Command-A on Mac OS X or Control-A on Windows and Linux). A count will appear in the right column (attachments are included in the count when visible, i.e. when the parent item is expanded).
-   Are you unable to adjust the size of the Zotero pane downwards beyond a certain point? Close the tag selector (e.g. by dragging the splitter above the tag selector down), as it has a minimum height.
-   You can convert the contents of the *title* and *publisher* fields to either sentence or title case by right-clicking the field and using the Transform Text menu.
-   The date fields will automatically convert “yesterday”, “today”, and “tomorrow” into the respective dates
-   When using Quick Copy, holding the “Shift” key while dragging and dropping items into a text document will insert citations instead of full references.
-   You can click the *DOI* and *URL* field labels to open the field link.
-   Manually adding authors to a Zotero item? You can use Shift+Enter after typing each name as a faster alternative to clicking the “+” button.
-   Need more than one Zotero library? You can use multiple profiles to keep your libraries separated. You can also maintain separate libraries within a single profile using Zotero Groups.
-   You can convert URL addresses in the bibliography to working hyperlinks using the AutoFormat function in Microsoft Word or LibreOffice. Follow these steps:
    1.  Click the Remove Field Codes button in the Zotero menu. This converts Zotero citations to plain text and disconnects your document from Zotero, so you should save a new copy of your document and use this function on only the final version.
    2.  Select your bibliography (or all of the text in the document, if desired) and apply the Word/LibreOffice AutoFormat (the keyboard shortcut is Ctrl/Cmd-Alt/Option-K in Word).

### Video Tutorials {#video-tutorials}
#### Video Tutorials#

A variety of library teams, Zotero users, and other community members have produced video tutorials for Zotero, including general overviews and demonstrations of specific features. Note that many of the videos available on the internet were produced using older versions of Zotero, and the current Zotero interface and functionality may look very different from what is shown.

**A list of videos produced after the release of Zotero 7.0 is available here:**

[List of Zotero tutorial videos](https://www.google.com/search?q=zotero&sca_esv=b3631baa3837d75b&sca_upv=1&biw=1916&bih=1483&source=lnt&tbs=cdr%3A1%2Ccd_min%3A8%2F9%2F2024%2Ccd_max%3A&tbm=vid)

### Viewing Site Certificate Information {#viewing-site-certificate-information}
#### Viewing Site Certificate Information#

When encountering security certificate errors in Zotero, viewing the certificate information for zotero.org or another affected domain in your web browser may provide more details on what is causing the problem on your system. If you share this information in the Zotero Forums, we may be able to help identify what is interfering (or it may be self-explanatory, based on, for example, the name of your institution or security software on your system).

1.  Load any zotero.org web page (such as this one) or, if you’re getting an error about another domain, an HTTPS URL for that domain. (If you’re getting a certificate error for s3.amazonaws.com, use [this URL](https://s3.amazonaws.com/zoterofilestorage/test).)
2.  View the certificate information. The process to do this varies by browser:
    -   In **Firefox**, click the padlock icon, click the right-arrow next to the domain name, and look at the “Verified by:” line. For full information, click More Information -> Security -> View Certificate and look at the details under Issued To, Issued By, and Period of Validity.
    -   In **Chrome (Windows)**, click the padlock icon, click on Certificate, and look at the “Issue to:” and “Issued by:” lines.
    -   In **Chrome (Mac)**, click the padlock icon, click on Certificate, and look at the “Issued by” line. For full information, expand the Details section and look at the values for Common Name and Organization.
    -   In **Safari**, click the padlock icon -> Show Certificate and look at Expires, Common Name under Subject Name, and Organization and Common Name under Issuer Name.

Secure connections to zotero.org will show that the certificate is issued by “Amazon”. If you see a different entity, something is likely intercepting your secure connections. (Something intercepting your connection could also self-identify as one of the expected entities, but this is rare.)

### What connections do I need to allow through a firewall for Zotero to work properly? {#what-connections-do-i-need-to-allow-through-a-firewall-for-zotero-to-work-proper}

While Zotero is a local program that can be used without internet access, it should generally be given the same access that browsers on the system have. Without such access, much of its functionality will be broken.

Zotero connects over HTTPS to various zotero.org subdomains for syncing, translator/style updating, retraction notifications, error reporting, version updates, and more, as documented in the Zotero privacy policy (which also explains how to disable each type of access). All Zotero domains are behind AWS load balancers, so resolved IP addresses are transient and cannot be used for allowlisting.

Beyond zotero.org domains, various core functionality depends on internet access that is unrestricted beyond standard content blocking: file saving (from any site a user saves from), Add Item by Identifier, PDF retrieval, PDF metadata retrieval, metadata updating, etc. Because of these required connections, it’s not possible to limit Zotero’s internet access to specific domains without breaking many features.

If you’re having trouble allowing Zotero to connect through a firewall or proxy, see Connection Error.

### What do I do if my Zotero database is corrupted? {#what-do-i-do-if-my-zotero-database-is-corrupted}

Zotero stores your information in a database file, zotero.sqlite, in the Zotero data directory. If the database becomes corrupted, Zotero may no longer be able to start up, or certain operations might fail.

Database corruption generally occurs when the data directory is placed in a cloud storage folder or on a network drive. If you’ve moved your data directory to one of those places, you should move it back to the default location. You should never store the data directory in cloud storage. (The same applies to any database-backed program.)

#### If You Were Using Syncing#

If your data is all in your web library, you can simply sync to pull down the latest version of your library, either from an empty local database (after moving your corrupted zotero.sqlite out of the way and restarting Zotero) or from an earlier, uncorrupted backup of the database. In the latter case, Zotero will simply pull down changes since the last time you used that database, without needing to redownload your entire library. Once you’re sure you’ve recovered your data, you can delete any copies of the corrupted zotero.sqlite file.

#### If You Weren’t Using Syncing#

If your data directory was in cloud storage or you have a local backup, you can try to restore from an earlier version of zotero.sqlite, including one of the automatic backups in your Zotero data directory, by closing Zotero, moving the current zotero.sqlite out of the way, and copying the backup file into place as zotero.sqlite. After restoring from a backup and starting Zotero, check the database integrity from the Advanced → Files and Folders pane of the Zotero preferences.

If you don’t have a backup, or the backups are corrupted as well, you can try to fix the damage with the Zotero Database Repair Tool.

### What does “Zotero” mean? {#what-does-zotero-mean}
#### What does “Zotero” mean?#

The name “Zotero” is loosely based on the Albanian (yes, Albanian) word zotëroj, meaning “to acquire, to master,” as in learning. If you are interested in more info see linguist Mark Dingemanse’s post on [The Etymology of Zotero](http://ideophone.org/zotero-etymology/).

### What is the difference between Links and Snapshots? {#what-is-the-difference-between-links-and-snapshots}
#### What is the difference between Links and Snapshots?#

When you attach a link to an item, Zotero stores only the page title, URL, and access date, and you need to return to the site to view the page content. When you save a snapshot, Zotero saves a copy of the page as it currently exists and archives it on your computer.

### What version of Zotero do I have? {#what-version-of-zotero-do-i-have}
#### What version of Zotero do I have?#
#### Zotero#

You can check your Zotero version by selecting About Zotero from the Zotero menu (Mac) or Help menu (Windows/Linux). Make sure you have the latest version listed on the changelog.

You can upgrade to the latest version of Zotero via Help -> “Check for Updates…” within Zotero or, if that’s not working correctly, from the download page.

#### Zotero for iOS / Android#

Tap Back in the top-left corner until you get to the libraries screen, and then tap the gear icon in the top right. The version will be listed at the bottom of the settings screen.

### Where are the link and snapshot buttons in Zotero 2.0? {#where-are-the-link-and-snapshot-buttons-in-zotero-20}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### Where are the link and snapshot buttons in Zotero 2.0?#

The separate link and snapshot buttons were removed from the Zotero toolbar in Zotero 2.0, but equivalent functionality still exists through the use of the Create New Item from Current Page button.

Create New Item from Current Page creates a web page item, which in Zotero 2.0 behaves much like a standalone link attachment but offers significantly more functionality and flexibility.

If “Automatically take snapshots when creating items from web pages” is enabled in the General pane of the Zotero preferences, the Create New Item from Current page button will create a child snapshot under the web page item. Holding down the Shift key while clicking the button will temporarily toggle the setting, allowing you to create a web page item with no snapshot even if the snapshot preference is enabled (or vice versa).

Double-clicking a web page item without a snapshot will take you to the web page, the same way that double-clicking a standalone web link attachment would take you to the web page in Zotero 1.0.

Double-clicking a web page item with a snapshot will display the snapshot instead. You can access the live web page by clicking the “URL:” label to the left of the URL field in Zotero’s right pane.

### Why am I getting an “unresponsive script warning”? {#why-am-i-getting-an-unresponsive-script-warning}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### Why am I getting an “unresponsive script warning”?#

While Zotero is performing long operations, Firefox may display the message:

*“A script on this page may be busy, or it may have stopped responding. You can stop the script now, or you can continue to see if the script will complete.”*

…and include a URL beginning with chrome://zotero/. Note that for all other URLs, including 0 URLs, the following advice does not apply.

**Short answer:** Click Continue until the message stops appearing. Don’t click Stop Script. This message has nothing to do with Chrome the browser.

**Long answer:** Zotero automatically disables this Firefox warning before beginning most long operations and re-enables it afterwards, but there may be places where the message still shows up, particularly on slower computers. If you receive the message repeatedly, report it in the Zotero Forums, and be sure to include the file and line number from the message in your post.

To prevent the message from appearing, type “” into the Firefox address bar and press Enter. Search for dom.max\_chrome\_script\_run\_time in the list and double-click it. To disable the warning completely, enter 0 in the dialog box that pops up. You can also set a longer timeout (in seconds) after which you should receive the message if a Firefox extension is still busy or has frozen. For example, to display the message after two minutes, enter 120. Unless you get frequent freezes afterward, there’s no downside to adjusting the timeout.

Note that, if you do receive the warning, pressing Stop Script in the middle of certain operations may disable Zotero until Firefox is restarted and/or may cause the message to reappear.

### Why can’t Zotero find a linked file? {#why-cant-zotero-find-a-linked-file}

There are a few possible reasons why a linked file might not be found:

1.  The file was moved or deleted outside of Zotero. Find the file and move it back to its original location, or use the Locate button in the File Not Found dialog to point Zotero to the new location.
2.  A Linked Attachment Base Directory is set incorrectly on one or more of your computers. If the Linked Attachment Base Directory is set correctly on your current computer but Zotero is still looking for a file at an absolute path from another computer, you must correct the Linked Attachment Base Directory setting on the other computer, which will convert attachments under the specified directory to use relative paths, and then sync both computers.
3.  Zotero is looking in the right place, but the file hasn’t yet synced from another computer. Zotero syncs only stored files, not linked files, so linked files must be synced using another tool.

If necessary, the third-party [Zutilo plugin](https://github.com/willsALMANJ/Zutilo) can be used to fix attachment paths in batch.

### Why can’t Zotero find a stored file? {#why-cant-zotero-find-a-stored-file}

There are a few reasons why a stored file might not be found:

1.  You created the file on another computer and it hasn’t yet synced to this computer. See Files Not Syncing.
2.  The individual file was moved or deleted outside of Zotero. You can use the Locate button in the File Not Found dialog to select the file. When using Locate, if you select a file that was renamed within the original storage directory, Zotero will update the database to point to the new file; if you select a file elsewhere, it will copy the file to the correct storage directory.
3.  You moved or deleted the ‘storage’ directory within the Zotero data directory.
4.  You somehow ended up using a different Zotero data directory from the one you were using previously on this computer, and you’re either viewing a different database or you synced down data — but not files, which might not be available online — from the online library. In this case, it’s better to locate the correct Zotero data directory on this computer.

Note that Zotero never deletes files on disk while their corresponding attachment items still exist in the database, so if your files are missing, either they haven’t synced or something happened outside of Zotero.

### Why do I no longer see a Save to Zotero option in the Firefox open/save dialog? {#why-do-i-no-longer-see-a-save-to-zotero-option-in-the-firefox-opensave-dialog}
#### Why do I no longer see a Save to Zotero option in the Firefox open/save dialog?#

Mozilla no longer allows extensions to modify built-in Firefox dialogs in this way.

To save a PDF from the Zotero Connector, go to the Applications pane of the Firefox preferences (“Options” on Windows) and set PDFs to “Preview in Firefox”. When viewing a PDF, you can then click the Zotero Connector save button to save the PDF to Zotero, or click the download button in the PDF viewer toolbar to get the open/save dialog.

If you see “Save to Zotero (Web Page…)” when viewing a PDF, the site is presenting the PDF in a webpage frame rather than serving the PDF directly. Right-click on the PDF and select This Frame -> Show Only This Frame to view the PDF directly, after which you will be able to save it to Zotero.

If you prefer not to preview PDFs in Firefox, you can also drag a PDF link from Firefox to Zotero to add it to your library.

Once you’ve added the PDF to Zotero, you can right-click on it and select “Retrieve Metadata for PDF” or “Create Parent Item”. This process will be automated in a future Zotero version.

### Why doesn’t Zotero have a “watch folder” feature? {#why-doesnt-zotero-have-a-watch-folder-feature}

People coming from other tools based around a “watch folder” sometimes look for similar functionality in Zotero and are surprised to find it doesn’t exist.

Zotero is designed around a different workflow from these other tools, and it reflects a bit of a philosophical difference. We don’t believe you should have to manually download files and worry about folder paths. Zotero is deeply integrated into the browser where you do most of your research, and it has an unmatched ability to save high-quality metadata from across the web.

To save to Zotero, you simply click the save button while on an article page, and Zotero will both save bibliographic details for the item and automatically download a PDF if one is available.

If only a PDF is available, you can still save it directly with the save button, and Zotero will attempt to retrieve metadata. (Make sure your browser is set to preview PDFs in the browser (at a 0 or 1 URL) rather than download them and show them at a 2 URL, which the Zotero Connector can’t access.)

If you already have a local file on disk (say, from an email), you can drag it to the Zotero app, but that’s not the primary way of adding items.

Once you’ve saved an item, Zotero manages your files for you, including renaming them based on the author, year, and title of the parent item and, if you use file syncing, making sure they stay linked across devices. There are plugins that can help you organize your files differently if you prefer, but file management is still something that’s done automatically for you based on your settings, not something you have to manage manually.

Ultimately, Zotero is designed to save time, and we try hard not to implement features that we believe encourage more tedious workflows. Since Zotero is open source, most things can be done with settings or plugins, but our goal is to create a default experience that just works for most people: one-click saves of data and files, automatic file management and syncing, advanced in-app organization and searching. Give it a try!

Still not convinced? Read through Adding Items to Zotero to better understand all the various ways of getting data and files into Zotero, and post to the Zotero Forums with other questions or comments. We’re always happy to talk about our design decisions there, and we can often recommend a better way of doing something if you’re having trouble.

### Why don’t I see a Zotero icon in the address bar when viewing a webpage? {#why-dont-i-see-a-zotero-icon-in-the-address-bar-when-viewing-a-webpage}
#### Why don’t I see a Zotero icon in the address bar when viewing a webpage?#

If you’ve used Zotero for Firefox previously and no longer see a Zotero save icon in the Firefox address bar, note that the button is now always present in the Firefox toolbar. See here for more details.

### Why don’t I see any tabs in the item pane after selecting an item? {#why-dont-i-see-any-tabs-in-the-item-pane-after-selecting-an-item}

If you don’t see an item pane at all, you’ve simply closed it. You can reopen it from the View menu by going to View → Layout and enabling “Item Pane”.

If you do see the item pane but don’t see any tabs, you’ve selected either an attachment or a note in the Zotero items list. Attachments and notes do not have full bibliographic metadata and cannot have other items attached to them, so you will not see the Info, Notes, Attachments, Tags, and Related tabs when you select them. Only *regular items* — with reference types such as books, journal articles, and newspaper articles — will show tabs in the right-hand column. See Adding PDFs and Other Files for more information on saving items with metadata or creating parent items for an existing standalone attachment.

### Why is there no save button in my browser toolbar? {#why-is-there-no-save-button-in-my-browser-toolbar}
#### Why is there no save button in my browser toolbar?#

The Zotero Connector places a save icon — representing the content it recognized on the page you’re viewing (book, journal article, webpage) — to the right of the address bar in your browser toolbar:

If you don’t see it, first make sure you’ve installed the Zotero Connector for your browser from the download page and that it shows up as enabled in the browser’s Extensions pane. If you’re having trouble installing it, make sure you’re running a compatible browser version.

If you still can’t find it, try the steps below for your browser.

##### Chrome#

When you first install a Chrome extension, it may appear initially in the Extensions popup, accessed via a puzzle-piece icon. If you find the button in that popup, click to pin it so that it stays visible.

If the extension shows as enabled in the Extensions pane but you don’t see the button anywhere in the toolbars, try uninstalling and reinstalling the extension. If it’s still not appearing, your Chrome settings may be corrupted, and you can try [resetting Chrome](https://support.google.com/chrome/answer/3296214).

##### Edge#

If the button doesn’t appear to the right of the address bar, it may appear in the Extensions popup, accessed via a puzzle-piece icon. If you find the button in the popup, click the “Show in Toolbar” icon to pin it so that it stays visible.

If the extension shows as enabled in the Extensions pane but you don’t see the button anywhere in the toolbars, try uninstalling and reinstalling the extension. If it’s still not appearing, your Edge settings may be corrupted, and you can try resetting them (Settings → Reset Settings).

##### Firefox#

If the button doesn’t appear to the right of the address bar, you can find it either in the [Extensions panel](https://support.mozilla.org/kb/extensions-button) (opened via a puzzle-piece button to the right of other toolbar buttons) or the overflow menu (opened via a “»” button to the right of your other toolbar buttons). If the button appears in the Extensions panel, click the gear icon and select Pin to Toolbar to keep it in the toolbar. If it appears in the overflow menu, you may need to resize your address bar to make more room for extension buttons.

If you still don’t see it, you may have a corrupted Firefox profile. Try creating a new [Firefox profile](http://support.mozilla.com/kb/Profiles) and installing the Zotero Connector into that profile.

##### Safari#

See Safari Compatibility.

### Why isn’t Zotero detecting updates? {#why-isnt-zotero-detecting-updates}
#### Why isn’t Zotero detecting updates?#
#### Updates require Administrator privileges#

Zotero does not require Administrator privileges for any of its operations and generally does not require them to install updates. If you are asked for a password or for Administrator privileges during Zotero update, one of two things is occurring:

1.  Zotero was initially installed with Administrator privileges. This is not necessary. Uninstall Zotero and re-install it from zotero.org/download. If prompted to give Administrator privileges during the installation process, click Cancel, rather than OK.
2.  Your IT administrators have set your system to require Administrator privileges to install program updates. Inquire with your IT administrator if Zotero can be whitelisted to allow it to update itself.

#### No automatic updates are found, but manual update checks work#

If you’re not receiving automatic updates to Zotero or an add-on but you can find updates manually by clicking Help → Check for Updates… (for Zotero) or Tools → Add-ons → Gear → Check for Updates… (for add-ons), check that automatic updates are enabled. For Zotero, open the Config Editor from the Advanced pane of Zotero preferences, and ensure that 0 and 1 are both set to 2. For add-ons, check Update Add-Ons Automatically in the Gear menu of the Zotero Add-ons window.

#### No automatic or manual updates are found#

If new versions of Zotero or add-ons aren’t being installed automatically and aren’t being detected when you manually check for updates, something on your system or network may be intercepting secure (HTTPS) connections to zotero.org or the add-on’s update server. To determine whether your connection is being intercepted, check the site certificate info.

#### Quick Fix#

If the site certificate information points to security software on your system (Bitdefender, Avast), disable the SSL/TLS/HTTPS scanning feature of that software. The exact name of the feature will vary. Consult the software’s documentation for help. Read on for more details.

#### More Details#

To ensure the security and privacy of its users, Zotero requires all connections to be made over HTTPS, which ensures that you’re connecting directly to a remote website and that your connection is encrypted. However, software installed on your system, or your network administrator, can override the security protections of HTTPS, essentially masquerading as any website. Some security software does this in an attempt to provide additional security: it intercepts HTTPS connections, scans the contents itself, and then re-encrypts the data and sends it to the original website in a new connection. While the makers of such software would argue that they’re protecting you with this feature by searching for malware served over HTTPS, this behavior breaks a fundamental security feature built into web browsers. You can think of it as someone going through all your postal mail, reading every letter, and warning you if they find any junk mail. While what they’re doing is potentially useful, they really have no business snooping around in your mail. (And in some cases, antivirus vendors have even [stuck their own ads and trackers into the envelopes](http://www.howtogeek.com/199829/avast-antivirus-was-spying-on-you-with-adware-until-this-week/) before resealing them.)

To receive updates automatically, you have two options:

1\) Disable the SSL/TLS/HTTPS scanning feature in the security software and try the update again. If the certificate information identifies your institution as the intercepting party, you’ll need to speak to your network administrator and request that they stop intercepting your secure connections to websites, though it may be a condition of your use of the network.

2\) If you trust the software or institution that is intercepting your connection, you can force Zotero to download updates over intercepted connections. Open the Zotero Config Editor from the Advanced pane of Zotero preferences. Right-click on the list of settings that appears and select New → Boolean. For Zotero, enter 0 for the property name and choose 1 for the value. For add-ons, enter 2 and choose 3 for the value. Be aware that there will no longer be a guarantee that you are receiving legitimate versions of Zotero or add-ons unless you return to the Config Editor and disable these preferences.

#### Zotero Connector for Firefox not detecting updates#

If the Zotero Connector for Firefox is not updating automatically, verify that automatic updates are enabled from the Firefox Add-ons window (Tools → Add-ons → Gear → Update Add-ons Automatically). Also check the Zotero Connector preferences and ensure that automatic updates are enabled there.

If automatic updates are enabled, try to update the connector manually by clicking the Gear button in the Firefox Add-ons window and choosing Check for Updates. If updates aren’t being detected manually, you may be encountering the secure connection interception issue described above. To determine whether your connection is being intercepted, check the site certificate info.

If the site certificate information points to security software on your system (Bitdefender, Avast), disable the SSL/TLS/HTTPS scanning feature of that software. The exact name of the feature will vary. Consult the software’s documentation for help.

As a temporary fix, you can manually install the updated extension from zotero.org/download. (Normally Firefox prevents manual installations over such connections as well, but we have implemented a workaround to allow them.) Without automatic updates, however, you may run into compatibility issues or bugs later that have already been fixed.

To receive updates automatically, you have two options:

1\) Disable the SSL/TLS/HTTPS scanning feature in the security software and try the update again. If the certificate information identifies your institution as the intercepting party, you’ll need to speak to your network administrator and request that they stop intercepting your secure connections to websites, though it may be a condition of your use of the network.

2\) If you trust the software or institution that is intercepting your connection, you can force Firefox to download add-on updates over intercepted connections. Enter “” in the Firefox address bar, right-click on the list of settings that appears, select New → Boolean, enter 0 for the property name, and choose 1 for the value. Be aware that there will no longer be a guarantee that you are receiving legitimate versions of Zotero and other Firefox add-ons unless you return to and disable that preference.

### Zotero 2.1 {#zotero-21}

We are excited to announce the release of Zotero 2.1. Zotero 2.1 is a major update over 2.0, offering many new features and bug fixes, including:

-   Firefox 4.0 compatibility\[1\]
-   A next-generation citation engine
    -   [citeproc-js](http://bitbucket.org/fbennett/citeproc-js/wiki/Home), written by Frank Bennett
    -   Supports [CSL](http://citationstyles.org/) 1.0 styles and includes many new features and bug fixes
-   Improved word processor integration
-   Support for displaying Zotero as a Firefox tab
-   A customizable Locate menu
-   Improved performance

See the changelog for a complete list of changes.

To install Zotero 2.1, click the “Download” button on the Zotero home page. Zotero 2.1 requires Firefox 3.6 or newer.

#### Word Processor Plug-ins for Zotero 2.1#

Installing word processor plugins for Zotero 2.1

#### Upgrading#
##### Upgrading from Zotero 2.0#

Upgrading from Zotero 2.0 is easy: just click the Download button on the Zotero home page. After installation, Zotero will prompt you to convert your 2.0 database to the new format used by 2.1, after which you will no longer be able to use your database with Zotero 2.0. Zotero will automatically back up your database in your Zotero data directory before upgrading it.

##### Upgrading from Zotero 1.0#

To upgrade from Zotero 1.0, click the Download button on the Zotero home page. After installation, Zotero will prompt you to convert your 1.0 database to the new format used by 2.1. When upgrading from Zotero 1.0, you should take extra precautions in case you run into trouble:

-   **Back up your data before upgrading your database**. While Zotero automatically backs up its database before every upgrade, the upgrade process from Zotero 1.0 to Zotero 2.1 requires irreversible changes to both the Zotero database and the attachments directory. Without a backup of your entire Zotero data directory, you will be unable to go back to Zotero 1.0 without losing data.
-   **Pick the right time to upgrade**. If you have a large database and/or a slow computer, the upgrade process may take a number of hours to run. You may also run into unforeseen issues after upgrading. Consider delaying your upgrade if you have an important deadline coming up soon and you don’t have time for troubleshooting.

Don’t want to upgrade now? If you’ve installed 2.1 but haven’t yet upgraded your database, you can reinstall Zotero 1.0, but note that Zotero 1.0 is no longer supported and no longer receives translator or style updates.

\[1\] Mac Word integration requires Mac OS X 10.6 (Snow Leopard) or later for Firefox 4 compatibility

### Zotero and the Text Encoding Initiative (TEI) {#zotero-and-the-text-encoding-initiative-tei}

The [Text Encoding Initiative](http://www.tei-c.org/) is a consortium that curates an international standard for the XML markup of texts. The standard enjoys broad support among libraries, museums, publishers, and scholars. Web-based TEI-compliant documents are found in increasing numbers online.

TEI provides for rich bibliographic metadata in the element 0 and its children, some of which are modeled after the [International Standard Bibliographic Description (ISBD)](http://en.wikipedia.org/wiki/International_Standard_Bibliographic_Description) (see TEI’s [note for library cataloguers](http://www.tei-c.org/release/doc/tei-p5-doc/en/html/HD.html#HD8)). For encoding bibliographic citations that may appear in a text, the element 1 is commonly used for a structured bibliographic citation, in which only bibliographic sub-elements appear and in a specified order.

TEI and Zotero are in theory a perfect fit, with Zotero able to detect and import a TEI-compliant text and, in turn, able to export to a well-formed 0 element. At this time, such interaction is still under development:

-   Zotero records can be exported to TEI-compliant XML
-   Zotero cannot currently import TEI-compliant XML. An import translator could be written, or a stylesheet could be created to convert TEI tags to Zotero-ready RDF.

And there is a TEI group on Zotero, in case you want to get acquainted with what has been published on TEI. As users in the Zotero and TEI communities build bridges to each other, they are encouraged to add new tools to this page.

### Zotero and Wikipedia/ Wikidata {#zotero-and-wikipedia-wikidata}
#### Zotero and Wikipedia/ Wikidata#
#### Wikipedia#

Zotero can be a powerful tool for Wikipedia users, be they casual visitors or regular contributors. By incorporating print-based references, Wikipedia users are able lend greater credibility to their articles and provide readers with a broader reference base, forging new relationships between online and offline resources. Wikipedia employs its own format for bibliographic citations, known as [Wikipedia Citation Templates](http://en.wikipedia.org/wiki/Citation_templates). Because the Wikipedia Citation Templates for books and articles embed COinS tags in each bibliographic entry—see, for example, the raw HTML markup of the [Wikipedia article for Zotero](http://en.wikipedia.org/wiki/Zotero)—Zotero can import references from articles that employ them. In addition, a Wikipedia Citation Templates export format in Zotero allows for the exporting of references to Wikipedia. This read/write relationship promises to allow better circulation of information through different scholarly communication networks.

*N.B.: if you wish to cite a web page in Wikipedia you must include a date in the accessed field.*

See: 0

#### Quick Copy#

Zotero’s Quick Copy feature makes it easy to export Zotero items to Wikipedia. Open the Export pane of Zotero preferences and select “Wikipedia Citation Templates” as the Default Format. Then, you can drag and drop items from your library to the Wikipedia to insert properly formatted citations. You can also copy Wikipedia Citation Templates data to your clipboad by pressing Ctrl/Cmd-Shift-C.

#### Wikidata#

Wikidata is a sister project of Wikipedia, operated by the same non-profit organisation (the Wikimedia Foundation), in the same volunteer-edited manner. It provides linked, open data about the entities described in Wikidata, and many more besides, under a Public Domain licence.

Zotero has two translators for use with Wikidata; one for reading Wikidata pages and adding their data to Zotero; the other for outputting data from other sources in a format that can be used create new Wikidata items, using a tool called QuickStatements.

For details, see : 0

### Zotero Beta Builds {#zotero-beta-builds}

If you want to run the latest prerelease version of a Zotero component (e.g., to take advantage of a recent bug fix or test a new feature), you can install a beta build. These versions are built regularly and may be less stable than release versions.

(If you want to make code changes, see Zotero Source Code.)

#### Zotero Beta#

**Don’t want to run a beta? Download the release version of Zotero 9.**

Beta versions of Zotero are currently built from the development line for Zotero 10.

-   Mac
-   Windows 64-bit Installer
-   Windows 64-bit ZIP
-   Windows ARM64 Installer
-   Windows ARM64 ZIP
-   Windows 32-bit Installer
-   Windows 32-bit ZIP
-   Linux 64-bit
-   Linux ARM64
-   Linux 32-bit

Beta versions will update automatically. You can update to the latest version at any time via Help → Check for Updates.

To revert to the release version, reinstall it from the download page, though note that beta versions occasionally update the database such that downgrading isn’t possible without reverting to a backup copy of zotero.sqlite in the Zotero data directory or removing the local data directory and syncing to pull down data from the online library. You’ll always be able to switch to the next major production release of Zotero.

On Windows, you may wish to use the ZIP version, which doesn’t contain an installer, to avoid overwriting your primary installation. (Your existing Zotero database will still be used.) Note that the ZIP version does not register Zotero as a handler for various file types.

#### Zotero Connector Beta#

The Zotero Connector Beta is used for testing connector-specific features. It’s not necessary to use a beta version of the connector when using the Zotero Beta above.

##### Firefox#

-   Zotero Connector for Firefox (beta)

Beta versions will update automatically, or you can update to the latest version at any time from the Firefox Extensions pane. To revert to the release version, simply reinstall it from the download page.

##### Safari#

A beta version of the Zotero Connector is included in the Zotero beta above.

##### Chrome/Edge#

There is no beta connector for Chrome or Edge.

#### Zotero for iOS Beta#

To try the latest beta of the iOS app, join the TestFlight group from your iOS device:

0

You can switch back to the release version at any time by reinstalling it from the App Store.

We ask that you remove yourself from the beta group in TestFlight if you switch back to the release version, as beta slots are limited.

#### Zotero for Android Beta#

To try the latest beta of the Android app, first [install Zotero for Android from the Play Store](https://play.google.com/store/apps/details?id=org.zotero.android), and then follow these steps:

1.  Open the Play Store app.
2.  At the top right, tap the Profile icon.
3.  Tap “Manage apps & devices” and then “Installed”.
4.  Tap on Zotero to open its details page.
5.  Under “Join the beta”, tap “Join” and then “Join”.

### The Zotero Data Directory {#the-zotero-data-directory}
#### Locating Your Zotero Data#

The easiest and most reliable way to find your Zotero data is by clicking the “Show Data Directory” button in the Advanced tab of the Zotero settings. This will reveal the folder on your computer that contains your Zotero database and attachment files.

##### Default Locations#

Unless you have selected a custom data directory in the Advanced tab of the Zotero settings, your Zotero data is stored within the following OS-dependent directories:

- **macOS**: **Windows 7 and higher**; 0: 0
- **macOS**: **Windows XP/2000**; 0: 0
- **macOS**: **Linux**; 0: 0


The “Show Data Directory” button will always reveal the data directory currently in use and is the recommended method for finding your data directory. If you’re unable to access the Zotero settings, a search for the file name ‘zotero.sqlite’ can also help you locate the Zotero data directory.

**Older Versions**

##### Zotero 4 for Firefox (2017 and earlier)

[TABLE]

##### Zotero 4 Standalone (2017 and earlier)

[TABLE]

#### Data Directory Contents#

The most important file in the data directory is the 0 file, which is the database containing the majority of your data: item metadata, notes, tags, etc. When Zotero starts up, it reads the 1 file in the active data directory.

The directory also contains a 0 folder with 8-character subfolders (e.g., “N7SMB24A”) containing all of your file attachments, such as PDFs, web snapshots, audio files, or any other files you have imported. (Files that are linked are not copied into this subfolder.)

Your data directory will likely contain several other files and folders. These can include 0 (an automatic backup of 1, which is updated periodically if the existing 2 file hasn’t been updated in the last 12 hours) and 3 files (automatic backups of 4 that are created during certain Zotero updates), as well as folders such as 5, 6, 7, 8, and 9 that are created automatically at Zotero startup.

**Warning**: Before you copy, delete, or move any of these files, be sure that Zotero is closed. Failure to do so before moving these files can damage your data.

#### Backing Up Your Zotero Data#

We strongly recommend that you regularly back up your Zotero data directory. While syncing is a great way to make sure you can restore your libraries if something happens to your computer, it’s not a complete substitute for a proper backup: the Zotero servers only store the most recent version of your libraries, and it takes just a single (possibly automatic) sync to change the server copy (though some inadvertent changes can be restored from Zotero’s automatic backups).

Rather than backing up just your Zotero database, we recommend using a backup utility that automatically backs up your entire hard drive to an external device on a regular basis and keeps incremental backups so that you can restore to a given version. Most modern operating systems offer such functionality (e.g., Time Machine on Macs).

If you really want to back up your Zotero data specifically, locate your Zotero data, close Zotero, and copy your data directory (the *entire folder*, including 0 and 1 and the other subfolders) to a backup location, preferably on another storage device. As with all important data, it’s a good idea to back up your Zotero data frequently, which is why we recommend an automated full-system backup instead.

Note that if you’re using “download files as needed” for file syncing, your attachment files may not all exist locally and may not be included in a backup. Zotero Storage provides reliable storage of uploaded files, so you might choose to exclude the 0 folder from your backup, but if you’d like a local backup of attachments as well, you would need to use “download files at sync time” on one computer and make a backup of the data directory from that computer.

**Warning**: You shouldn’t use export (e.g., to Zotero RDF, BibTeX, or RIS) as a backup method. Exporting and re-importing a library doesn’t produce an exact copy — it will reset Date Added/Modified times and break links to existing citations in word processor documents, along with other potential changes.

#### Restoring Your Zotero Data From a Backup#

Between manual backups, automatic backups, and synced data, it’s often possible to restore a lost Zotero library or restore data that was accidentally deleted.

Before following these steps, be sure that Zotero is looking in the right place for your data.

##### Restoring Your Zotero Data Using Zotero Syncing#

If you were using Zotero syncing and have an empty local library, you can likely restore your data simply by syncing with your online library. After verifying that your library is correct on zotero.org, simply reenter your username and password in the Sync tab of the Zotero settings and click the Sync button in the toolbar. (Zotero only syncs explicit deletions, so just syncing an empty library won’t overwrite the server data **unless you deleted items manually**.)

If you have a local Zotero library that you want to overwrite, close Zotero and delete the old Zotero data directory before syncing. Syncing your database with a different Zotero account will also prompt you to remove the existing local database.

##### Restoring Your Zotero Data From a Backup#

If you were not using Zotero syncing (or were but don’t want to perform a full sync) and have a backup of your Zotero data directory, you can restore your library by replacing your active data directory with your backed-up data directory.

Open the Advanced tab of the Zotero settings and make a note of the specified path under Data Directory Location. (By default, this will be “Zotero” within your home folder.) Click “Show Data Directory”, which should reveal your active data directory containing zotero.sqlite and possibly a ‘storage’ subdirectory. Close Zotero, change to the parent folder of the active data directory (Cmd-up-arrow on macOS, Alt-up-arrow on Windows), and rename the folder to “Zotero-Old”. Next, copy the data directory from your backup to the original location (e.g., “Zotero”).

When you reopen Zotero, you should see your restored Zotero data.

Once you’ve successfully restored your data, you can delete the “Zotero-Old” folder, but it’s a good idea to keep it for a while until you’re sure your data is correct.

Note that, if you were using Zotero syncing, any changes you made to your library since the backup and subsequently synced to your online library will be applied to your restored database as soon as you sync. If you don’t want that to happen, see the following section.

##### Restoring Your Zotero Data From a Backup and Overwriting Synced Changes#

If you or someone else made unwanted changes to your Zotero library and synced those changes to your online library, you may be able to restore data by using a local backup of your Zotero data directory.

1.  Temporarily disable auto-sync in the Sync tab of the Zotero settings.
2.  Follow the steps in the preceding section to restore from a backup of your Zotero data directory.
3.  Once you see your restored data, were you to sync again, the more recent data in the online library would replace the data you just restored, and you’ll need to take steps to prevent that:
    -   If you’re trying to restore a small number of deleted items or notes, you can simply duplicate the items — by right-clicking and choosing “Duplicate Item(s)” — so that the new copies remain even after syncing. You could also make a local change (e.g., adding items to a collection) to trigger a conflict, and then choose the local versions when you sync.
    -   If you’re trying to restore deleted collections, you can create duplicate collections and drag items from the old collections to the new ones. When you sync, the old collections will be deleted but the new ones will remain.
    -   If many items were affected or collections were deleted, you can use Replace Online Library to force Zotero to upload the local version of the library, overwriting previously synced changes. Note that will delete any other changes you’ve made locally since the backup.

If you’re happy with the results, you can re-enable auto-sync and continue working.

##### Restoring From the Last Automatic Backup#

If you make a critical mistake while using Zotero — for example, if you accidentally delete a large set of items — you may be able to revert to the last automatic backup. Note that automatic backups contain only data, not files.

1.  If you’re using syncing, temporarily disable auto-sync in the Sync tab of the Zotero settings.
2.  Locate your Zotero data and make a backup copy of any zotero.sqlite.bak files. The timestamps of the files may help you determine which file would contain the data you’re trying to restore.
3.  Close Zotero. In your data directory, rename zotero.sqlite to zotero.sqlite.old, rename one of the original .bak files (based on the timestamp) to zotero.sqlite, and restart Zotero. You should now see the backed-up version of your library.
4.  If you were using syncing and the undesired changes were already synced, syncing again now would cause the more recent data in the online library to replace the data you just restored, and you’ll need to take steps to prevent that:
    -   If you’re trying to restore deleted collections, you can create duplicate collections and drag items from the old collections to the new ones. When you sync, the old collections will be deleted but the new ones will remain.
    -   If you’re trying to restore a small number of deleted items or notes, you can simply duplicate the items — by right-clicking and choosing “Duplicate Item(s)” — so that the new copies remain even after syncing.
    -   If many items were affected or collections were deleted, you can use Replace Online Library to force Zotero to upload the local version of the library, overwriting previously synced changes.

If you’re happy with the results, you can re-enable auto-sync and continue working. Keep zotero.sqlite.old and your .bak file backups until you’re sure all your data is intact and in sync across all your computers.

##### Restoring From the Last Upgrade Backup#

When you upgrade to a major new version of Zotero, Zotero will automatically update your database to work with the new version. If you would like to revert to a previous version of Zotero at a later point, you will have to manually replace your database with the automatic backup Zotero made during the upgrade. In most cases this will be the highest-numbered “zotero.sqlite.\[num\].bak” file in your Zotero data directory.

It’s a good idea to make a backup of your entire Zotero data directory before making any changes.

If you have synced your data with the Zotero servers, reverting to a previous version is as simple as reinstalling the previous version, closing Zotero, replacing “zotero.sqlite” in your Zotero data directory with “zotero.sqlite.\[highest-number\].bak”, and restarting Zotero. (Note that if you try to open an upgraded database in an earlier version, Zotero will display an error. Just close Zotero and replace the .sqlite file.) Zotero will then sync from the online library any changes made since you last used the older database.

If you were not using syncing, you may wish to export to Zotero RDF any items added since the database upgrade and then reimport those into the earlier version. Sorting your library by Date Added may help you find such items.

To temporarily disable further updates, go to the Config Editor in the Advanced tab of the Zotero settings and set 0 to false. Note that staying on an old version is not a long-term solution, as old versions are no longer supported and may stop syncing or receiving site-compatibility updates at any time. Be sure to post to the Zotero Forums and explain whatever is causing you to downgrade, and make a note to check back periodically to see whether it makes sense for you to re-enable automatic updates.

#### Locating Missing Zotero Data#

If you open Zotero to find your library blank or missing lots of data, there are a few main possibilities:

-   If you were using a very old version of Zotero — from 2017 or earlier — without installing any updates and just upgraded to a newer version of Zotero, see Missing Data After Zotero 5 Upgrade. Zotero 5 was released many years ago, so this no longer applies to most people.
-   If you’re using a different computer from the one where you created the missing data, and your data is also missing in your online library, your data simply hasn’t synced from the computer where you created it. See Changes Not Syncing.
-   If you know you’ve had the data on this computer previously, something may have happened to your previous Zotero database, or Zotero may be looking in the wrong place for your data. Read on for instructions.

To determine what happened to your data on this computer, first locate your current Zotero data directory by going to the Advanced → Files and Folders section of the Zotero settings and using the “Show Data Directory” button. Take note of the names, sizes, and dates of the files beginning with “zotero.sqlite” in this folder, which are your Zotero database (zotero.sqlite) and automatic database backups (\*.bak). An empty Zotero database will be either approximately 1 MB (~1,000 KB) or 5 MB.

If you see only 1 MB or 5 MB zotero.sqlite files, look in the ‘storage’ folder, if one exists, for folders with dates corresponding with your previous usage of Zotero.

-   If you see folders in ‘storage’, this is likely the Zotero data directory you were using previously, but something happened to the zotero.sqlite database outside of Zotero — for example, you might have accidentally deleted zotero.sqlite using system tools while trying to clear disk space. In this case, you can restore your data through backups and/or Zotero syncing:
    1.  Look for larger zotero.sqlite.bak files in the data directory, or look for a larger zotero.sqlite file in any separate backups you have. (It’s not possible to restore your data from the ‘storage’ files alone.) When Zotero starts up, it reads the zotero.sqlite file in the active data directory, so you can try other copies of zotero.sqlite by copying them to that location and filename. Do not try to import an .sqlite file into Zotero via File → “Import…” — it won’t work.
    2.  Whether or not you have a backup, if you’ve been using Zotero syncing, you can sync to pull down all data in your online library. If you do have a backup, all data more recent than the backup will be downloaded. If you only have an empty database, all data will be downloaded. In either case, you won’t overwrite data in the online library simply by syncing — syncing doesn’t work that way.
    3.  If you can’t find any other copies of zotero.sqlite and weren’t using Zotero syncing, you’ll unfortunately probably need to recreate your database from scratch. Close Zotero, move the Zotero data directory to your desktop as “Zotero Old”, and restart Zotero to create a new library. You can search for all PDFs within your “Zotero Old” folder and drag them to Zotero, and Zotero will attempt to retrieve metadata for as many of them as possible. You can also extract data from any Word or LibreOffice documents you used with the Zotero word processor plugin by using [Reference Extractor](https://rintze.zelle.me/ref-extractor/), though note that any data you re-import this way won’t be linked to your existing documents.
-   If this isn’t the location you were expecting to be using, or if you don’t see a ‘storage’ folder or it’s empty, you’ll need to locate your previous data directory on this computer. Once you find it, either select that data directory from the Zotero settings or, with Zotero closed, rename the current directory (e.g., to “Zotero-Old”) and move your desired Zotero directory to the specified location. If you’re not sure where your most recent Zotero data is located, look for versions of zotero.sqlite or zotero.sqlite.bak larger than 5 MB with appropriate modification times stored elsewhere on your computer and look at the dates of the folders within the ‘storage’ folder.
    -   Unless you have a good reason to use a custom data directory location, we strongly recommend using the default location in your home directory.
    -   When specifying a custom data directory location, keep in mind that Zotero doesn’t move or copy any data. You still need to copy your data into the specified location. Also, when pointing the data directory location to an existing folder, be sure to specify the folder containing zotero.sqlite and ‘storage’, not the ‘storage’ folder.

If you’ve gone through these steps and aren’t sure what to do, post to the Zotero Forums with the following info:

-   The names, sizes, and dates of all files beginning with “zotero.sqlite” in your current data directory
-   Whether there’s a ‘storage’ folder containing subfolders with dates corresponding to your previous usage of Zotero
-   Whether your current data directory is in the default location (“Zotero” in your home folder)
-   When you last used Zotero on this computer, and what happened on your computer since then
-   What you’ve tried so far

### Zotero Documentation {#zotero-documentation}
#### Quick Links#

-   Installation
-   Quick Start Guide
-   Getting Help
-   Zotero Storage Subscriptions
-   Frequently Asked Questions
-   Version History
-   System Requirements
-   For Developers

#### Using Zotero#

-   **Getting Stuff Into Your Library**
    -   Adding Items
    -   Adding Files
    -   Feeds
    -   Retrieve PDF Metadata
    -   Importing from Other Reference Managers
-   **Organizing Your Library and Taking Notes**
    -   Collections and Tags
    -   Searching
    -   Sorting
    -   PDF Reader
    -   Notes
    -   Related Items
    -   Duplicate Detection
-   **Generating Bibliographies, Citations, and Reports**
    -   Creating Bibliographies within Zotero
    -   Word Processor Integration
    -   Citation Styles
    -   Reports
-   **Syncing, Collaboration, and Backup**
    -   Data and File Syncing
    -   Groups
    -   Share your work with My Publications
    -   Backup
-   **Zotero Preferences**
    -   Preferences
    -   Zotero Connector Preferences
    -   Proxies
    -   Languages and Localization
-   **Getting the Most Out of Zotero**
    -   Knowledge Base
    -   Locate items in your institution’s library or other databases
    -   Plugins
    -   Zotero for Mobile
    -   Tips and Tricks
    -   Community-developed Video Tutorials

#### Small Print#

-   Contact Us
-   Accessibility
-   Credits and Acknowledgments
-   Licensing
-   Security
-   Privacy Policy

### Zotero doesn’t open when I click the Zotero status bar icon or select “Zotero” from the Firefox Tools menu. {#zotero-doesnt-open-when-i-click-the-zotero-status-bar-icon-or-select-zotero-from}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### Zotero doesn’t open when I click the Zotero status bar icon or select “Zotero” from the Firefox Tools menu.#

Disable all other Firefox extensions, beginning with those that are known to cause problems with Zotero.

If disabling all extensions allows you to open Zotero, re-enable your extensions one-by-one, restarting Firefox each time, until you find the conflicting extension. If the conflicting extension isn’t listed on the Known Issues page, report it in the forums. You should also contact the author of the other extension, who will likely need to fix the problem.

### Zotero Item Types and Fields {#zotero-item-types-and-fields}

This page provides general descriptions of Zotero item types and fields. It can be helpful when entering item data into Zotero manually. The page also provides guidance on ways to store information for unusual item types.

#### Item Types#

Item types in Zotero should be regarded as flexible, broad categories. Item types are generally determined based on how items should be cited. See descriptions of Zotero’s item types below. If you are citing an unusual item type that does not perfectly fit one of the supported categories, choose an item type that is reasonably close with an adequate selection of fields (e.g., 0, 1, or 2 can be made to fit most unusual item types well).

You can store more specific information about the type of item (e.g., a novel versus a biography book) in the Type, Format, or Extra fields or using tags.

- Item Type: Artwork; Description: A piece of artwork (e.g., an oil painting, photograph, or sculpture). Also use this item type for other types of images or visual items (e.g., scientific figures).
- Item Type: Audio Recording; Description: Any form of audio recording, including music, spoken word, sound effects, archival recordings, or audio-based scientific figures.
- Item Type: Bill; Description: A proposed piece of legislation.
- Item Type: Blog Post; Description: An article or entry posted to a personal blog website. For online articles published as part of a larger online publication (e.g., [NYT Blogs](http://www.nytimes.com/interactive/blogs/directory.html)), using 0 or 1 generally yields better results.
- Item Type: Book; Description: A book or similar published item. For government documents, technical reports, manuals, etc., use 0 instead. This item type can also be adapted to fit many types of unusual items.
- Item Type: Book Section; Description: A section of a book. Usually chapters, but also forewords, prefaces, introductions, appendices, afterwords, comments, etc.
- Item Type: Case; Description: A legal case, either published or unpublished.
- Item Type: Conference Paper; Description: A paper presented at a conference and subsequently published in a formal conference proceedings publication (e.g., as a book, report, or issue of a journal). For conference papers that have not been published in a proceedings, use 0.
- Item Type: Dataset; Description: A collection of data.
- Item Type: Dictionary Entry; Description: An entry published as part of a dictionary.
- Item Type: Document; Description: A generic document item. This item type has a poor selection of fields and poor support in citation styles, so it should generally be avoided.
- Item Type: E-mail; Description: A message sent via email. This type could also be used for other forms of personal communication.
- Item Type: Encyclopedia Article; Description: An article or chapter published as part of an encyclopedia.
- Item Type: Film; Description: A film or motion picture. Generally, use this type for artistically-oriented films (including fictional, non-fictional, and documentary films). For other types of video items, use 0.
- Item Type: Forum Post; Description: A post on an online discussion forum. Also use this type for items such as Facebook posts or tweets.
- Item Type: Hearing; Description: A formal hearing or meeting report by a legislative body.
- Item Type: Instant Message; Description: A message sent via an instant message or chat service. This type could also be used for other forms of personal communication.
- Item Type: Interview; Description: An interview with a person, including recordings, transcripts, or other records of the interview.
- Item Type: Journal Article; Description: An article published in a scholarly journal (either print or online).
- Item Type: Letter; Description: A letter sent between persons or organizations. This type could also be used for other forms of personal communication.
- Item Type: Magazine Article; Description: An article published in a non-scholarly, popular, or trade magazine (either print or online).
- Item Type: Manuscript; Description: An unpublished manuscript. Use this type for both historical documents and modern unpublished work (e.g., unpublished manuscripts, manuscripts submitted for publication, working papers that are not widely available). Can also be used for other forms of historical or archival documents. This item type can also be adapted to fit many types of unusual items.
- Item Type: Map; Description: A map. Also use this type for geographic models.
- Item Type: Newspaper Article; Description: An article published in a newspaper (either print or online).
- Item Type: Patent; Description: A patent awarded for an invention.
- Item Type: Podcast; Description: A [podcast](https://en.wikipedia.org/wiki/Podcast) (an episode of an audio or video program distributed online, often via subscription).
- Item Type: Preprint; Description: A version of a scholarly or scientific paper that precedes formal peer review and publication in a peer-reviewed scholarly or scientific journal. The preprint may be available, often as a non-typeset version available for free, before or after a paper is published in a journal, through institutional repositories or preprint servers. This item type can also be used for working papers.
- Item Type: Presentation; Description: A presentation made as part of a conference, meeting, symposium, lecture, etc. This item type refers to the presentation itself, not a written version published as part of a conference proceedings (use 0 for such published versions).
- Item Type: Radio Broadcast; Description: An audio broadcast, such as a radio news show, an episode of a radio entertainment series, or similar. Includes broadcasts from online radio stations and audio broadcasts archived online (cf. 0).
- Item Type: Report; Description: A report published by an organization, institution, government department, or similar entity. Also used for working papers and preprints distributed through institutional repositories or preprint servers. This item type can also be adapted to fit many types of unusual items.
- Item Type: Software; Description: A piece of software, an app, or another computer program.
- Item Type: Standard; Description: A published document describing a standard, or a similar set of industrial or governmental specifications.
- Item Type: Statute; Description: A law or other piece of enacted legislation.
- Item Type: Thesis; Description: A thesis submitted as part of a student applying for a degree (either published or unpublished).
- Item Type: TV Broadcast; Description: An episode of a television series.
- Item Type: Video Recording; Description: A video recording. Use this type for general video items that do not fit into one of the more specific video item types (e.g., Film, TV Broadcast), such as YouTube videos or video-based scientific figures.
- Item Type: Webpage; Description: An online page of a website. When possible, use one of the more specific item types above (e.g., Magazine Article, Blog Post, Report).
- Item Type: Attachment; Description: A standalone attachment file (e.g., a PDF, JPEG, DOCX, PPTX, XLSX, or ODT file). Standalone attachment files have limited functionality in Zotero (e.g., they cannot be properly searched or cited). Always attach files to proper Zotero items.
- Item Type: Note; Description: A standalone note. Notes can be used for organizing and annotating in Zotero. If you cite a standalone note, Zotero will use the first 120 characters as the item title (and will treat the note as an author-less and date-less item). Citing notes is not a reliable way to add standalone commentary to a bibliography or reference list.


For additional legal and historical item types, see Legal Citations.

#### Item Fields#

Fields included for various item types in Zotero are described below. Fields are organized into categories based on the kinds of items they are most typically associated with. Some fields (e.g., Volume, Series) may also be present for other types of items. Specific meanings of fields for particular item types are noted where necessary.

Fields marked with an asterisk (\*) cannot be used in citations

##### General Fields#

The following fields exist across multiple item types and have the same purpose for all (or nearly all) types.

[TABLE]

##### Fields for Books and Periodicals#

[TABLE]

##### Fields Related to When and How Items Were Accessed#

- Field: Accessed; Description: Date an electronic resource was accessed. Typically filled automatically. You can enter “today,” “yesterday,” and “tomorrow” to quickly enter the corresponding date.
- Field: DOI; Description: The [Digital Object Identifier](http://en.wikipedia.org/wiki/Digital_object_identifier) of an item. For items other than journal articles, the DOI can be stored in Extra. See Citing Fields from Extra below.
- Field: URL; Description: URL (web-address) where the full item was accessed. Do not use this field for links to catalog records (e.g., libraries or Google Books) or abstracts—these should be instead added as links.
- Field: Archive; Description: Mainly for archival resources, the archive where an item was found. Also used for repositories, such as government report databases, institutional repositories, or subject repositories.
- Field: Loc. in Archive; Description: The location of an item in an archive, such as a box and folder number or other relevant location information from the finding aid. Include the subcollection/call number, box number, and folder number together in this field. For additional tips on citing archival sources in Zotero, see [here](https://web.archive.org/web/20150220133939/0).
- Field: Library Catalog; Description: The catalog or database an item was imported from. This field is used, for example, in the MLA citation style. Uses of this field are broader than actual library catalogs.
- Field: Call Number; Description: The call number of an item in a library. For citing archival sources, also include the Call Number in 0 (if applicable).


##### Fields for Reports and Theses#

[TABLE]

##### Fields for Presentations and Performances#

[TABLE]

##### Fields for Recordings and Broadcasts#

[TABLE]

##### Fields for Images, Artwork, and Maps#

- Medium: Artwork Size; The type of artwork or figure or the medium it was created with (e.g., "Watercolor painting", "Wood sculpture", "X-ray Crystallograph", "Scatterplot").: The dimensions of the artwork or figure.
- Medium: Scale; The type of artwork or figure or the medium it was created with (e.g., "Watercolor painting", "Wood sculpture", "X-ray Crystallograph", "Scatterplot").: The scale of a map or model.
- Medium: Type; The type of artwork or figure or the medium it was created with (e.g., "Watercolor painting", "Wood sculpture", "X-ray Crystallograph", "Scatterplot").: For maps, a description of the specific type or genre of map.


##### Fields for Primary Sources and Personal Communications#

- Medium: Type; For interviews, the format in which an interview was recorded (e.g., "Audio Recording", "Video Recording", "Transcript").: For letters, a description of the specific type or purpose of letter (e.g., "Private correspondence). For manuscripts, a description of the type or status of the manuscript (e.g., "Unpublished manuscript", "Manuscript submitted for publication", "Working paper").
- Medium: Subject; For interviews, the format in which an interview was recorded (e.g., "Audio Recording", "Video Recording", "Transcript").: The title of an email.


##### Fields for Websites#

[TABLE]

##### Fields for Software#

- Version: System; The version of a computer program.: The operating system or platform a computer program is written for.
- Version: Company; The version of a computer program.: The organization publishing a computer program.
- Version: Language; The version of a computer program.: The programming language a computer program is written in. *Note that* Computer Program *uses the* 0 *field differently than other items.*


##### Additional Fields#

- Rights\\: Date Added; The copyright terms, license, or release status for an item.: The date and time an item was added in Zotero. This field is filled automatically.
- Rights\\: Date Modified; The copyright terms, license, or release status for an item.: The date and time an item was last modified in Zotero. This field is filled and updated automatically.
- Rights\\: Extra; The copyright terms, license, or release status for an item.: Free field for storing additional information. You can also store additional variables not included in an item’s fields that can be used when making citations and bibliographies . See Citing Fields from Extra below.


##### Fields for Legal Items#

For additional and more flexible support for citing legal materials, see Legal Citations.

##### Legislation and hearings#

- Name of Act: Bill Number; The full title of a statute.: The identification number assigned to a proposed piece of legislation.
- Name of Act: Code; The full title of a statute.: The name of the code in which a bill or statute is published.
- Name of Act: Code Volume; The full title of a statute.: The volume number of the code containing a bill or statute.
- Name of Act: Code Number; The full title of a statute.: The identification number assigned to a piece of legislation in the code in which it is published.
- Name of Act: Public Law Number; The full title of a statute.: The identification number assigned to an enacted piece of legislation. See [here](https://www.loc.gov/law/help/statutes.php) for more details.
- Name of Act: Date Enacted; The full title of a statute.: The date a statute was enacted.
- Name of Act: Section; The full title of a statute.: The section of a bill or statute being cited.
- Name of Act: Committee; The full title of a statute.: The committee holding a hearing.
- Name of Act: Document Number; The full title of a statute.: The document identification number assigned to a published transcript or record of a committee hearing.
- Name of Act: Code Pages; The full title of a statute.: The pages in the code volume containing a bill or statute.
- Name of Act: Legislative Body\\; The full title of a statute.: The legislative body (parliament, senate, etc.) debating a bill, holding a hearing, or passing legislation.
- Name of Act: Session; The full title of a statute.: The session of a legislative body in which a bill was introduced, statute was passed, or hearing was held.
- Name of Act: History; The full title of a statute.: Resources related to the procedural history of a bill or statute.


##### Legal cases#

- History: Case Name; Resources related to the procedural history of a legal case.: The title of a legal case.
- History: Court; Resources related to the procedural history of a legal case.: The court where a legal case was argued.
- History: Date Decided; Resources related to the procedural history of a legal case.: The date a legal case was decided.
- History: Docket Number; Resources related to the procedural history of a legal case.: The docket number assigned to a to-be-heard legal case.
- History: Reporter; Resources related to the procedural history of a legal case.: The reporter in which a legal case is published.
- History: Reporter Volume; Resources related to the procedural history of a legal case.: The volume number of the reporter in which a legal case is published.
- History: First Page; Resources related to the procedural history of a legal case.: The first page of the reporter volume on which a case appears.


##### Patents#

- Country: Assignee; The country issuing a patent.: The entity to which a ownership rights for a patent are assigned.
- Country: Issuing Authority; The country issuing a patent.: The authority or office reviewing the application and issuing the patent.
- Country: Patent Number; The country issuing a patent.: The identification number assigned to a patent.
- Country: Filing Date; The country issuing a patent.: The date on which a patent application was filed.
- Country: Issue Date; The country issuing a patent.: The date on which a patent is formally issued.
- Country: Application Number; The country issuing a patent.: The identification number assigned to a patent application.
- Country: Priority Numbers; The country issuing a patent.: The international application number of a patent used for priority rights claims.
- Country: References; The country issuing a patent.: Resources related to the history of a patent.
- Country: Legal Status; The country issuing a patent.: The legal status of a patent or application.


#### Item Creators#

Creator types marked with an asterisk (\*) cannot be used in citations.

[TABLE]

##### Role Labels for Media Creators#

For citations to films, recordings, and broadcasts, Zotero currently has limited to support for labeling producers, scriptwriters, and some other Creator roles.

To label directors, leave the main Zotero field blank (or enter the names as Contributor) and enter the director names using 0 in Extra. See Citing Fields from Extra below.

All Creator roles (Director, Producer, Scriptwriter, etc.) can also be labelled by entering the names using the default 0 role for the item (*Performer* for Audio Recording, *Podcaster* for Podcast, and *Director* for Film, Radio Broadcast, TV Broadcast, and Video Recording) and adding the appropriate labels in parentheses after the authors’ first names or last names as appropriate for your citation style—e.g., MacNaughton \|\| Ian (Producer) for APA style. Note that the labels will be rendered verbatim in citations; enter abbreviated terms (e.g., “Prod.”) here as needed.

If the style uses initials for author first/given names, rather than full names (e.g., APA style), if the label contains multiple words (e.g., “Executive Producer” or “Writer & Director”), Zotero will abbreviate the words of the label after the first. To avoid this, type a “Word Joiner” character (Unicode U+2060, printed here between quotes: “”) on either side of each space in the label.

See also Media Creator Roles

#### Additional Item Types and Fields#
##### Citeable Item Types not Included in Zotero#

These item types are not yet formally supported in Zotero. For citation purposes, you can convert an item of a different type to one of these types by entering them in the Extra field in the following format:

    Type: CSL Type

For example:

    Type: review-book

- Item Type: Figure; CSL Type: 0; Description: A figure included in a scientific or academic work.
- Item Type: Musical Score; CSL Type: 0; Description: The written score for a musical work.
- Item Type: Pamphlet; CSL Type: 0; Description: An informally-published work. Typically smaller and less technical than a Report.
- Item Type: Book Review; CSL Type: 0; Description: A review of a book. Enter these as Journal, Magazine, or Newspaper Article, depending on where they were published, by providing a “Reviewed Author” creator.
- Item Type: Treaty; CSL Type: 0; Description: A legal treaty between two nations.


##### Citeable Fields not Included in Zotero#

These item fields are not yet formally supported in Zotero. For citation purposes, you can convert an item of a different type to one of these types by entering them in the Extra field in this format:

    CSL Variable: Value

For example:

    PMID: 123456
    Status: in press
    Original Date: 1886-04-01
    Director: Kubrick || Stanley

- Field: PMID; CSL Variable: 0; Description: The [PubMed identifier](http://en.wikipedia.org/wiki/PMID#PubMed_identifier).
- Field: PMCID; CSL Variable: 0; Description: The [PubMed Central identifier](https://en.wikipedia.org/wiki/PubMed_Central#PMCID).
- Field: Status; CSL Variable: 0; Description: The publication status of an item (e.g., “forthcoming”, “in press”, “advance online publication”).
- Field: Submitted Date; CSL Variable: 0; Description: The date an item was submitted for publication.
- Field: Reviewed Title; CSL Variable: 0; Description: The title of a reviewed work.
- Field: Chapter Number; CSL Variable: 0; Description: The number of the chapter within a book.
- Field: Archive Place; CSL Variable: 0; Description: The geographic location of an archive.
- Field: Event Date; CSL Variable: 0; Description: The date an event took place. Enter in ISO format (year-month-day).
- Field: Event Place; CSL Variable: 0; Description: The geographic location of an event.
- Field: Original Date; CSL Variable: 0; Description: The original date an item was published. Enter in ISO format (year-month-day).
- Field: Original Title; CSL Variable: 0; Description: The original title of a work (e.g., the untranslated title).
- Field: Original Publisher; CSL Variable: 0; Description: The publisher of the original version of an item (e.g., the untranslated version).
- Field: Original Publisher Place; CSL Variable: 0; Description: The geographic location of the publisher of the original version of an item (e.g., the untranslated version).
- Field: Original Author; CSL Variable: 0; Description: A type of Creator. The original creator of a work.
- Field: Director; CSL Variable: 0; Description: A type of Creator. The director of a film, recording, or broadcast. In Zotero, “Director” is mapped to CSL author. If you need special labels for directors—”(Dir.)”, enter the 1 label in Extra.
- Field: Editorial Director; CSL Variable: 0; Description: A type of Creator. The managing editor of a publication (“Directeur de la Publication” in French).
- Field: Illustrator; CSL Variable: 0; Description: A type of Creator. The illustrator of a work.


##### Citing Fields from Extra#

If a Zotero item type is missing fields that are needed for citations, it is possible to add these fields to the Extra field.

Enter each variable on a separate line at the top of the Extra field in the following format:

    CSL Variable: Value

For example:

    DOI: 10.1128/AEM.02591-07
    Original Date: 1824
    PMCID: PMC3531190

With the exception of Item Type (CSL 0) and Date variables (CSL 1, etc.), variables entered in Extra will not override corresponding values entered in proper Zotero fields.

##### Dates#

Dates entered in Extra will override the date entered in Zotero’s Date field. Dates must be entered in ISO format (year-month-day). Date ranges can be entered in this format:

    Issued: 2001-12-15/2001-12-31

##### Names#

For Creator variables, separate two-field names entered in Extra with two vertical bar characters (‘\|\|’), like this:

    Editorial Director: De Gaulle || Charles

### Zotero Keyboard Shortcuts {#zotero-keyboard-shortcuts}
#### Zotero Desktop App#

Note that the Shortcuts pane of the Zotero preferences allows you to change some of these keyboard shortcuts.

##### Adding Items to your Zotero Library#

- Function: Save to Zotero; Windows/Linux: 0+1+S; Mac OS: 2+3+S
- Function: Create a New Item by Hand; Windows/Linux: 0+1+N; Mac OS: 2+3+N
- Function: Create a New Note; Windows/Linux: 0+1+O; Mac OS: 2+3+O
- Function: Import; Windows/Linux: 0+1+I; Mac OS: 2+3+I
- Function: Import from Clipboard; Windows/Linux: 0+1+2+I; Mac OS: 3+4+5+I


##### Editing Items (Info Tab)#

- Add Another Author/Creator when Editing Creator: Save Abstract or Extra field; 0+1: 0+1; 2+3: 2+3


#### Removing or Deleting Items and Collections#

- Function: Move to Trash; from My Library: 0; from a Collection: 1+2 / 3+4
- Function: Move to Trash without Confirmation Dialog; from My Library: 0+1 / 2+3; from a Collection: *Not Available*
- Function: Remove from Collection; from My Library: *Not Applicable*; from a Collection: 0 *(Only for top-level items)*


- Function: Delete Collection (and Keep Items in Library and Others Collections, if any); Key: 0
- Function: Delete Collection and Move Items to Trash; Key: 0+1 / 2+3


##### Creating Citations and Bibliographies (Quick Copy)#

- Function: Copy Selected Item Citations to Clipboard; Windows/Linux Defaults: 0+1+A; Mac OS Defaults: 2+3+A
- Function: Copy Selected Items to Clipboard; Windows/Linux Defaults: 0+1+C; Mac OS Defaults: 2+3+C


##### Navigating between Zotero Panes#

- Function: Focus Libraries (Left) Pane; Windows/Linux: 0+1+L; Mac OS: 2+3+L
- Function: Move through Panes and Fields; Windows/Linux: 0/1+2; Mac OS: 3
- Function: Move through Info/Notes/Tags/Related Tabs; Windows/Linux: 0 and 1 (or 2+3/4+5+6 or 7+8/9+10); Mac OS: 11+12/13+14\[1\]
- Function: Quick Search; Windows/Linux: 0+1+K; Mac OS: 2+3+K
- Function: Quick Search; Windows/Linux: 0+F; Mac OS: 1+F


##### Moving Between Tabs#

Zotero supports most standard shortcuts for switching between tabs:

-   0+1/2 (3+4+5/6 on most Mac keyboards)
-   0+1 / 2+3-4
-   0+1+2/3 (macOS only)
-   0+1+2/3 (macOS only)
-   0/1+2 through 3
-   0/1+2 (opens the list all tabs menu)

##### Searching#

- Function: Quick Search; Windows/Linux: 0+1+K; Mac OS: 2+3+K
- Function: Quick Search; Windows/Linux: 0+F; Mac OS: 1+F
- Function: Find/Highlight Collection(s) an Item belongs to; Windows/Linux: Hold down 0 (Windows) or 1 (Linux); Mac OS: Hold down 2


##### Tags#

- Function: Toggle Tag Selector; Windows/Linux: 0+1+T; Mac OS: 2+3+T
- Function: Assign Colored Tag to an Item; Windows/Linux: 1 to 6 keys; Mac OS: 1 to 6 keys


##### Feeds#

- Function: Mark All Feed Items as Read/Unread; Windows/Linux: 0+1+R; Mac OS: 2+3+R
- Function: Mark Feed as Read/Unread; Windows/Linux: 0+1+\2Cmd3Shift4


##### Other Shortcuts#

- Function: Expand/Collapse Collections or Items List; Windows/Linux: 0/1; Mac OS: 2/3
- Function: Highlight All Collections an Item is in; Windows/Linux: *Hold down* 0 *(Windows) or* 1 *(Linux)*; Mac OS: *Hold down* 2
- Function: Count Items (Result Appears in Right-Hand Pane); Windows/Linux: 0+A; Mac OS: 1+A
- Function: Edit Collection Names (Left Pane); Windows/Linux: 0; Mac OS: 1


##### Reader#

These shortcuts apply in Zotero’s built-in PDF, EPUB, and snapshot reader. Some are specific to a particular document type, as noted.

##### Annotation Tools#

- Function: Toggle highlight tool; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle underline tool; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle note tool; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle text tool *(PDF)*; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle area (image) selection tool *(PDF)*; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle draw (ink) tool *(PDF)*; Windows/Linux: 0+1; macOS: 2+3
- Function: Toggle eraser tool *(PDF)*; Windows/Linux: 0+1; macOS: 2+3
- Function: Select (pointer) tool *(PDF)*; Windows/Linux: 0; macOS: 1
- Function: Hand (pan) tool *(PDF)*; Windows/Linux: 0; macOS: 1
- Function: Cycle to next annotation color; Windows/Linux: 0+1 *(PDF)*, 2+3 *(EPUB/snapshot)*; macOS: 4+5 *(PDF)*, 6+7 *(EPUB/snapshot)*
- Function: Choose annotation color (while a tool is active); Windows/Linux: 0–1; macOS: 2–3


##### Creating and Editing Annotations#

- Function: Create annotation from selected text *(PDF)*; Windows/Linux: 0+1+2/3/4/5/6(highlight/underline/note/text/area); macOS: 7+8+9/10/11/12/13 (highlight/underline/note/text/area)
- Function: Create annotation from selected text *(EPUB/snapshot)*; Windows/Linux: 0+1+2/3/4 (highlight/underline/note); macOS: 5+6+7/8/9 (highlight/underline/note)
- Function: Select all annotations; Windows/Linux: 0+1; macOS: 2+3
- Function: Move selected annotation *(PDF)*; Windows/Linux: 0/1/2/3; macOS: 4/5/6/7
- Function: Resize selected highlight/underline *(PDF)*; Windows/Linux: 0+1+2/3/4/5; macOS: 6+7+8/9/10/11
- Function: Resize selected image/text/ink annotation *(PDF)*; Windows/Linux: 0+1/2/3/4; macOS: 5+6/7/8/9
- Function: Resize/extend selected annotation *(EPUB/snapshot)*; Windows/Linux: 0+1/2 (add 3 for word, 4+5 to adjust start); macOS: 6+7/8 (add 9 for word, 10+11 to adjust start)
- Function: Delete selected annotation; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Undo annotation change; Windows/Linux: 0+1; macOS: 2+3
- Function: Redo annotation change; Windows/Linux: 0+1+2; macOS: 3+4+5
- Function: Deselect annotation/tool, close popup; Windows/Linux: 0; macOS: 1


##### Navigation#

- Function: Scroll down / go to next page; Windows/Linux: 0, 1; 2+3 *(PDF)*; 4/5 *(EPUB)*; macOS: 6, 7; 8+9 *(PDF)*; 10/11 *(EPUB)*
- Function: Scroll up / go to previous page; Windows/Linux: 0+1, 2; 3+4 *(PDF)*; 5/6 *(EPUB)*; macOS: 7+8, 9; 10+11 *(PDF)*; 12/13 *(EPUB)*
- Function: Go to beginning / end of document; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Move focus to the “go to page” box; Windows/Linux: 0+1+2; macOS: 3+4+5
- Function: Go back (after following a link); Windows/Linux: 0+1, 2+3 *(Linux)*; macOS: 4+5, 6+7
- Function: Go forward; Windows/Linux: 0+1, 2+3 *(Linux)*; macOS: 4+5, 6+7


##### Zoom and Display#

- Function: Zoom in / increase font size; Windows/Linux: 0+1 / 2+3; macOS: 4+5 / 6+7
- Function: Zoom out / decrease font size; Windows/Linux: 0+1; macOS: 2+3
- Function: Reset zoom; Windows/Linux: 0+1; macOS: 2+3


##### Finding Text#

- Function: Find text in the document; Windows/Linux: 0+1; macOS: 2+3
- Function: Find next; Windows/Linux: 0+1; macOS: 2+3
- Function: Find previous; Windows/Linux: 0+1+2; macOS: 3+4+5


##### Miscellaneous#

- Function: Print the document; Windows/Linux: 0+1; macOS: 2+3


##### Read Aloud#

Zotero’s reader can read your documents aloud using natural-sounding voices. Press 0 or 1 to start reading from the current location or text selection; the shortcuts below apply while Read Aloud is active.

- Function: Start Read Aloud; Windows/Linux: 0 or 1; macOS: 2 or 3
- Function: Play / pause; Windows/Linux: 0; macOS: 1
- Function: Skip to previous / next sentence; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Skip to previous / next paragraph; Windows/Linux: 0+1 / 2+3; macOS: 4+5 / 6+7
- Function: Highlight / underline the current passage; Windows/Linux: 0 / 1; macOS: 2 / 3


After annotating a passage:

- Function: Dismiss; Windows/Linux: 0 or 1; macOS: 2 or 3
- Function: Delete annotation; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Move annotation; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Extend annotation by sentence; Windows/Linux: 0+1/2; macOS: 3+4/5
- Function: Switch to highlight / underline; Windows/Linux: 0 / 1; macOS: 2 / 3
- Function: Set annotation color; Windows/Linux: 0–1; macOS: 2–3


The annotation is saved immediately, and the popup will close on its own after a short time.

##### Notes#

**This section hasn't been fully updated for the new note editor in Zotero 6.**

- Function: Bold; Windows/Linux: 0+B; Mac OS: 1+B
- Function: Italic; Windows/Linux: 0+I; Mac OS: 1+I
- Function: Underline; Windows/Linux: 0+U; Mac OS: 1+U
- Function: Select All; Windows/Linux: 0+A; Mac OS: 1+A
- Function: Undo; Windows/Linux: 0+Z; Mac OS: 1+Z
- Function: Redo; Windows/Linux: 0+Y or 1+2+Z; Mac OS: 3+Y or 4+5+Z
- Function: Cut; Windows/Linux: 0+X; Mac OS: 1+X
- Function: Copy; Windows/Linux: 0+C; Mac OS: 1+C
- Function: Paste; Windows/Linux: 0+V; Mac OS: 1+V
- Function: Paste without formatting; Windows/Linux: 0+1+V; Mac OS: 2+3+V
- Function: Format Heading 1 to 6; Windows/Linux: 0+1+1/6; Mac OS: 2+3+1/6
- Function: Format as Paragraph; Windows/Linux: 0+1+7; Mac OS: 2+3+7
- Function: Format as Div; Windows/Linux: 0+1+8; Mac OS: 2+3+8
- Function: Format as Address; Windows/Linux: 0+1+9; Mac OS: 2+3+9
- Function: Find and Replace; Windows/Linux: 0+F; Mac OS: 1+F
- Function: Insert Link; Windows/Linux: 0+K; Mac OS: 1+K
- Function: Focus/jump to toolbar; Windows/Linux: 0+F10; Mac OS: 1+2


#### Zotero Connector#

Chrome, Firefox, and Edge all provide a way to assign a keyboard shortcut to the Save to Zotero button. Consult your browser documentation for more information.

#### Word Processor Plugins#

-   Word
-   LibreOffice
-   Google Docs

#### Customize your Shortcuts#

The third-party [Zutilo plugin](https://github.com/willsALMANJ/Zutilo) adds various functions not available in Zotero itself through extra menu items and keyboard shortcuts.

\[1\] MacBooks and the Apple Wireless Keyboard don’t have dedicated Page-Up/Page-Down keys, so you have to use Fn+Up/Down-arrow on those to simulate the Page-Up/Page-Down keys.

### Zotero Privacy Policy {#zotero-privacy-policy}
#### Overview#

Zotero is an open-source project committed to providing the best tool for managing your research. Our philosophy is that what you put into Zotero is yours, and one of our founding principles is to make sure you remain in control of your data and can share it how you like — or choose not to share it at all.

**We are an independent, nonprofit organization and have no financial interest in your private information.** We fund further development by offering additional online storage space to people who find the software useful, not by selling data.

#### Data We Collect#

Zotero is designed as a local program that saves data to your own computer by default, and it doesn’t require sharing any data with us to be usable. However, some of Zotero’s advanced features require you to supply us with information.

-   We collect the information you voluntarily provide (e.g., your name and email address) if you create a Zotero account and use optional Zotero services. Zotero can be used without creating an account.
-   We collect the library data you upload if you choose to synchronize your library with the Zotero servers. Syncing is entirely optional and is disabled by default.
-   We collect the attachment files you upload if you choose to use Zotero file syncing. File syncing is optional, and you can choose to sync files in your personal library using a WebDAV server instead.
-   We keep a record of the most recent IP addresses used to access your synchronized library in order to let you verify the security of your account and allow you to revoke access in your Zotero settings.
-   We log visits to our website, including IP address and browser, in order to prevent abuse and to diagnose technical issues. We retain these access logs for 90 days.
-   We log requests made to our servers by Zotero or third-party software, including IP address and client information, in order to prevent abuse, diagnose technical issues, and assess usage. We retain these logs for up to 90 days. You can opt out of all requests to our servers.
-   We collect error reports that you submit. See Support Interactions.
-   If you attempt to save something to Zotero and the save fails, we collect information about the failure, including the URL, browser, and version information, so that we can more quickly fix site compatibility issues. We store this information for up to one week. No additional personally identifying information (e.g., username or IP address) is stored, and reports are generally only viewed in aggregate. This error reporting can be disabled in the Zotero and Zotero Connector preferences.

#### Library Statistics#

Zotero anonymizes and aggregates synchronized user and group library data to generate statistics on readership. This anonymized and aggregated data only includes publicly available metadata (e.g., publication title and author). This data is never sold or made available in any forms other than ones offered publicly. We may also use this anonymized and aggregated user information for auditing, research, and analysis to operate and improve Zotero services.

#### Security of Stored Data#

See Security of Zotero Data.

#### Disabling Automatic Requests#

You can disable all automatic communication with Zotero servers from the Zotero and Zotero Connector preferences:

-   **Syncing:** Sync preferences → leave unconfigured or disable automatic syncing
-   **Automatic PDF metadata retrieval:** General preferences → disable “Automatically retrieve metadata for PDFs”
    -   We do not log any information about the contents of PDF metadata requests.
-   **Open-access PDF retrieval:** General preferences → disable “Automatically attach associated PDFs and other files when saving items”
    -   If a PDF can’t be saved for an item with a DOI, Zotero will send the DOI to Zotero servers to check for open-access versions. We do not log the contents of these requests. Disabling this preference will disable all automatic attachment saving.
-   **Broken site translator reporting:** disable “Report broken site translators” in the Advanced pane of Zotero and the Zotero Connector
-   **Translator/style update checking:** Advanced preferences → disable “Automatically check for updated translators and styles”
-   **Zotero update checking:** Advanced preferences → Config Editor → set 0 to false
    -   Automatic update checking is strongly recommended for security and stability reasons.
-   **Retracted item checking:** Advanced preferences → Config Editor → set 0 to false
    -   Retraction checks are performed without sharing the items you have in your database.
-   **Proxy authentication checking:** Advanced → Config Editor → set 0 to false.
    -   At Zotero startup, HEAD requests are made to a test file on Amazon S3 and selected publisher websites (controlled by 0) to trigger a proxy authentication prompt if and only if Zotero detects that a proxy is required to connect to the internet. If you disable this option and require an authenticated proxy, Zotero network connections will fail.

If automatic syncing or automatic translator/style updates are enabled, Zotero will maintain a persistent connection to Zotero servers when it is open in order to provide immediate updates. You can disable this connection by disabling both of those options or by setting 0 to false in the Config Editor.

If you use the Zotero Connector without having Zotero open, the Connector will make a daily request to Zotero servers for information on available site translators. It will then download translators for the sites you visit. For example, if you load a *New York Times* article, the Connector will download Zotero’s *New York Times* translator and cache it. If Zotero doesn’t have a translator for a specific site, no request will be made. No information on the specific pages you visit is transmitted, and subsequent requests won’t be made for the same translator until you restart your browser or the translator is updated. You can avoid these requests by keeping Zotero open while you browse the web.

#### Permissions Warnings#

When using third-party platforms, we request the most restrictive permissions available that still allow Zotero to perform its advertised functions. In some cases, the necessary permissions can sound a bit scary, so we want to explain why they’re necessary.

##### Zotero Connector#

When installing the Zotero Connector, your browser will warn you that the extension can **“Read your browsing history”**, **“Access your data for all websites”**, or similar. These different wordings all mean the same thing: that Zotero can interact with each page as you browse the web. This is the standard permission that browser extensions that run on all pages require. Zotero uses it to determine what content it can save on a given page and update the save button accordingly, to fetch metadata and other resources that are served from outside of the domain being saved from, and to provide advanced features such as automatic proxy redirection. Zotero in no way reads your previous browsing history, and no data about your browsing activity is stored except when you choose to save a page to either your local or online Zotero library.

The Zotero Connector also requires permission to **“Block content on any page”**, which is the technical mechanism it uses to enable saving of CSL files to the Zotero app and to show CAPTCHAs when saving PDFs on some websites. No content is being blocked.

##### Google Docs Integration#

When you first use Google Docs integration, Google will ask you to grant Zotero Google Docs Integration permission to “See, edit, create, and delete all your Google Docs documents”. The plugin requires this permission to insert citations into your documents. The plugin doesn’t do anything else with your document content and doesn’t access documents other than the ones on which it’s triggered. The integration works entirely locally on your computer, so even when you trigger the plugin on a given document, nothing is sent to Zotero servers.

#### Storage Purchases#

We send two pieces of personal data to our payment processor at the time of purchase:

-   your Zotero username, in order to associate your payment with your Zotero account
-   your email address, so that the payment processor can send you a payment receipt

We do not collect or store your credit card, banking, or other payment information. When you enter your full name, billing address, and account numbers to complete a purchase, this information passes directly to our payment processor. We never have access to payment account numbers. We send your name and address to our tax service provider in order to calculate applicable taxes for the sale and record the transaction for tax remittance.

#### Support Interactions#

Most Zotero support occurs in the public Zotero Forums. If you would like to remove forum posts you have made, you may clear them yourself at any time, though we encourage you to leave your posts up for the benefit of others. If you’d prefer your forum posts to appear under a different name, you can change your forums username from your account settings.

You may be asked to submit an error report or debug output to help us troubleshoot problems. These reports contain technical information about your computer, such as your operating system and installed browser extensions, and may include incidental personal information such as URLs of sites you visited before or while generating the report. You can review the output of these reports before submitting them. We don’t store any personal information (username, IP address) that links the report to you, and we generally don’t look at reports unless they are referenced by a Report ID or Debug ID in the Zotero Forums. Reports are stored for up to one year.

If you email us, we collect your email address and any other information you provide, and we may store your messages indefinitely to provide context for any future support requests you make.

#### Third-Party Services Used#

-   All Zotero server data is stored in the United States in Amazon Web Services. See Security of Zotero Data for more information.
-   Certain operations you perform in Zotero may trigger requests to public third-party services such as Crossref or the Library of Congress for metadata retrieval. These third parties may log your IP address and search terms (e.g., DOI or ISBN) according to their privacy policies, but no other identifying information is provided.
-   When you save items to or update items in Zotero, Zotero may, depending on your settings, connect to the associated sites to download metadata or save PDFs or snapshots. This is equivalent to loading those sites in your browser, and similar privacy implications apply. See Zotero and Firewalls for access requirements.
-   When using Premium voices with Read Aloud, document text is sent to external providers for text-to-speech processing. Currently, these providers may include Inworld AI or Google. No additional information beyond the document text is sent to the third-party providers. Standard Read Aloud voices are processed entirely on Zotero servers.
-   Some fonts on Zotero’s websites are licensed from myfonts.com. In order to verify Zotero’s compliance with this license, myfonts.com collects your IP address and the URL of the accessed Zotero webpage.
-   When an account is registered, we use Google reCAPTCHA, hCaptcha, or Cloudflare Turnstile to verify that it is not an automated registration attempt. We also do this for login attempts.
-   Payment processing is provided by [Stripe](https://stripe.com).
-   For institutional storage purchases, invoices and payments may be processed by [Intuit](https://www.intuit.com/privacy/statement).
-   Tax calculation, processing, and remittance services are provided by [Anrok](https://www.anrok.com/) and [VAT IT](https://vatit.com/).

#### Deleting Your Data#

You may delete your Zotero account to remove information you voluntarily provided when you registered to use Zotero services and to remove the library data you provided if you chose to synchronize your library with Zotero servers. To delete your account, visit your Zotero settings and click “Permanently Delete Account”.

#### Backed-up Data#

We make regular automated backups of data on our servers to protect against accidental loss of user data. These backups are intended for disaster recovery and would be accessed only in the event of significant data loss. Backups may be retained for up to 6 months.

#### Legally Compelled Disclosure#

We may be legally required to comply with requests for data from law enforcement or government agencies.

#### Changes#

We may update our privacy policies over time. Up-to-date information, including details of new features, will always be available from this page.

#### Questions#

If you have any questions or concerns regarding Zotero’s privacy policies, please ask us in the Zotero Forums or email privacy@zotero.org.

### Zotero Security {#zotero-security}

*If you believe you’ve found a security issue in Zotero software, please contact security@zotero.org.*

Zotero was created with the philosophy that your research data belongs to you and should be kept secure and private by default.

All Zotero software is [open source](https://github.com/zotero) and can be audited for security and privacy practices. Zotero builds for macOS and Windows are code-signed, and all builds are distributed with transport encryption, ensuring that the version you run is the version we released.

Unlike many cloud-based tools, the Zotero desktop application is a local program that runs on your computer and saves all research data locally by default. Unless you explicitly set up syncing, your research data never leaves your computer.

If institutional policies prevent uploading of data to third-party servers, Zotero can always be used locally without syncing any data, but syncing is required to use group functionality.

If you choose to sync your data with the Zotero servers, all data is encrypted in transit with current best practices (Zotero’s API endpoint receives an A+ score on the well-respected SSL Labs test) and stored within the Amazon cloud, where access is tightly restricted to the small number of Zotero staff members who need access to maintain the service. Data in newly created accounts is also encrypted at rest using AES-256. All data is currently stored in the us-east-1 AWS region in the United States.

While library data and group files can be synced only with Zotero servers, for syncing of files in personal libraries you can choose between Zotero servers and a WebDAV server under your control, or you can use linked files that are stored in a location of your choosing and aren’t synced by Zotero.

The Zotero data server is open-source and can be run locally, which some organizations choose to do, but this can be technically challenging, and we don’t currently provide support for such installations.

See our privacy policy for further details on Zotero’s collection and use of data you choose to share.

### Zotero System Requirements {#zotero-system-requirements}
#### Zotero#

-   macOS 10.15 or later
-   Windows 10 or later
-   Linux (same library requirements as [Firefox 140](https://www.firefox.com/en-US/firefox/140.0/system-requirements/#gnulinux))

#### Zotero Connector#

-   Chrome (current Stable or Extended Stable version)
-   Edge (current Stable or Extended Stable version)
-   Firefox 115 or later
-   Safari 16.6+ on macOS 11 Big Sur or later (details)

#### Word Processor Plugins#
##### Word for Windows#

-   Word 2010-2024 or Office 365, excluding Word 2010 Starter Edition

##### Word for Mac#

-   Word 2016–2024 or Office 365

##### LibreOffice#

-   LibreOffice 5.2 or later
-   Java Runtime Environment (JRE) or Java Development Kit (JDK)
    -   You will generally be prompted to install a JRE automatically upon installation if necessary. Some Linux users may need to install the JRE included in their distribution.
    -   On macOS, LibreOffice requires the JDK, not the JRE. See the troubleshooting instructions for more information.

### ZOTERO TERMS OF SERVICE {#zotero-terms-of-service}
#### ZOTERO TERMS OF SERVICE#

Welcome to zotero.org (the “Site”). Zotero is a collection of services, including storage subscription services (the “Services”), integrated with the Zotero desktop software and the Site, operated by the Corporation for Digital Scholarship (“CDS”) from its offices within the United States on a not-for-profit basis. Your access to and use of the Services and the Site are subject to these Terms of Service (“Terms of Service”) and all applicable laws. CDS makes no representation that the Services made available on or accessed through the Site are appropriate or available for use in other locations, and access to them from territories where such access is illegal is prohibited. CDS may change or modify these Terms of Service, in whole or in part, at any time by updating this posting, without prior notice or liability to users. By accessing and/or using the Services - whether you are a “Visitor” (which means that you simply browse the Site) or a “Registered User” (which means that you have registered as a Zotero user with CDS) - you acknowledge that you have read, understood and agree to be bound by these Terms of Service and to comply with all applicable laws. The terms “you”, “your” or “user”, as used in these Terms of Service, refer to a Visitor or a Registered User.

#### 1. Your Account/Registration#

Registration is required to subscribe to the Services. You must be 13 years or older to subscribe to the Services. If you are between age 13 and 18, you confirm that you have your parent’s or legal guardian’s consent to use the Services, that both you and they have read and agreed to these Terms of Service, and they have agreed to be considered a Registered User for purposes of the account. By registering, you represent and warrant to CDS that: (a) you are 18 years of age or older and the age of majority in your state of residence as of the time you register as a Registered User; (b) all information provided by you to CDS during the registration process is truthful, accurate and complete; (c) you will comply with all terms and conditions of these Terms of Service; and (d) you will not use the Services for any purpose that is unlawful or prohibited by these Terms of Services.

As a Registered User, you agree to maintain and promptly update your registration data as necessary to keep it accurate, current and complete. CDS may terminate your access to the Services, without prior notice or liability to you, if any of the information provided is found to be inaccurate, false, out of date or incomplete, or for violating these Terms of Service and/or the law.

As a Registered User, you acknowledge that you are solely responsible for all activities that occur under your account while using the Services. You agree to abide by all applicable laws in connection with your use of the Services, including those related to intellectual property rights, data privacy, international communications and the transmission of technical or personal data.

You are responsible for maintaining the security of your account as well as monitoring and controlling access to your account. You agree to notify CDS immediately of any unauthorized use of any account, or any other known or suspected breach of security. CDS cannot and will not be liable for any loss or damage in the event of an unauthorized use by a third party of your account. CDS, in its sole discretion, has the right to suspend or terminate your account and refuse any and all current or future use of the Services, or any other service, for any reason at any time. Termination of the Services will result in the deactivation or deletion of your account or your access to your account, and the forfeiture and relinquishment of all content in your account. If CDS terminates your account or services without cause, it will issue a pro-rata refund of any unused Services. CDS reserves the right to refuse service to anyone for any reason at any time.

#### 2. Your Submissions and Other Data#

CDS does not claim ownership of any data or other content you transmit, upload or store on or through the Services by any means (“Submissions”). You retain your rights to the Submissions you transmit, upload or store on or through the Services. You are solely responsible for all your Submissions and all activity that occurs under your account. You have sole responsibility for the accuracy, quality, integrity, legality, reliability, appropriateness, and intellectual property ownership or right to transmit, post or upload your Submissions and to grant the rights granted by you herein. By using the Services, you automatically grant to CDS and its service providers a royalty-free license and right to store, display, process, modify, and retransmit your Submissions.

In order to provide and ensure the quality of the Zotero Services, we may collect and store certain data consistent with our Privacy Policy, which can be found at 0 and is incorporated herein by reference.

CDS is not responsible for screening or monitoring data transmitted, uploaded or stored through the Services by you or other users. If notified of any data transmitted, posted or uploaded through the Services allegedly in violation of these Terms of Service, CDS may investigate the allegation and determine in good faith and its sole discretion whether to remove such data or any portion thereof. CDS shall have no liability or responsibility to users for performance or nonperformance of such activities.

CDS reserves the right not to store and to remove from storage any data uploaded or submitted through the Services for any reason, including, without limitation, any data that, in its sole discretion, it deems to be in violation of these Terms of Service.

#### 3. Fees and Payment#

The Services are offered under free and/or paid subscriptions. One person or legal entity may not maintain more than one free subscription. If a free account is inactive (i.e., there has been no synchronization of files) for ninety (90) days, CDS reserves the right to delete any and all files in that account without notice or liability to you.

The fees for paid subscriptions are not license fees, but charges due for storage and related services. The Zotero research tool and server software are released for free public use under several open source licenses, primarily the AGPLv3 - see the license text included with the code for details. The source code is available at 0.

You agree to provide CDS with complete and accurate billing and credit card information in connection with your paid subscription(s). If the information you have provided is false or fraudulent, CDS reserves the right to terminate your access to the Services in addition to any other legal remedies.

The fees for the paid services are set out at 0 and are subject to change from time to time. You agree to pay the subscription fees and any other charges (including any applicable taxes) for the specific paid Services plan you select in accordance with the fees and billing terms in effect at the time you subscribed to such services. All fees and charges, when paid, are nonrefundable and accrue on the first day of the initial subscription term or successive renewal term until terminated, regardless of whether or not you actually use the Services.

Once you begin your annual billed paid subscription, your next payment would be due on the same date the following year. Should you elect to upgrade your paid subscription to a larger amount of storage space, you will be billed on the commencement date for the first year of the upgraded level of Services. Your initial new subscription term will be 12 months, plus the amount of time your remaining prepaid fees on your prior subscription would buy on a pro-rata basis under the new subscription. For example, if you initially signed up for a $20/year subscription on January 1, 2000, your renewal date would be January 1, 2001. However if on July 1, 2000 you decided to upgrade to the $60/year subscription, your initial new subscription term would be 14 months — from July 1, 2000 to August 31, 2001. This is because your initial new subscription would include an additional two months, as the prior unused subscription fee (here $10) would cover an additional two months under the $60/year subscription plan. After your initial new subscription term you would return to a 12-month subscription term.

You may downgrade your plan, but downgrading may cause the loss of content, features or capacity of your account. If you opt to downgrade, you agree to be solely responsibility for any related loss; CDS cannot and will not be held liable for such loss.

Payment for all paid subscriptions is handled for CDS by a third party service provider, and your credit card statement may include this provider’s identifier. If you believe any fees or charges to your account are incorrect, you must contact us in writing within thirty (30) days of the date of the bill containing the amount in question to be eligible to receive an adjustment or credit.

CDS reserves the right to terminate a free account at any time in its sole discretion. CDS also reserves the right to impose fees, with at least fourteen (14) days prior notice to you, for your continued use of Services that were offered for free at the time you subscribed.

#### 4. Renewals#

Renewal charges will be based on fees in effect at the time of renewal. CDS agrees to provide you at least fourteen (14) days prior notice of a fee increase via an email to the email associated with your account. Such increase shall be effective upon renewal. The paid subscription will automatically renew for fixed periods of time equal to your initial paid subscription purchase on the anniversary of your initial purchase. A paid subscription, however, will not be renewed automatically if you terminate such subscription in the manner and time described in Section 5 below. If renewal fees are not paid in a timely manner, or if we are unable to process your transaction using the credit card information provided, CDS will provide you notice of your account delinquency. If you do not bring your balance current within thirty (30) days after we provide you with notification that your account is in arrears, CDS reserves the right to delete some or all of your Submissions so as to reduce your storage space to below the ceiling of free subscriptions and to convert your paid subscription to a free subscription.

#### 5. Paid Subscription Termination#

You may terminate your paid subscription at any time through 0. You must terminate your paid subscription before its renewal date in order to avoid automatic renewal and a charge of the next period’s fees to your credit card. If you terminate your paid subscription, you will not be charged at the next renewal date. We will not charge a fee for terminating your paid subscription.

If you terminate your paid subscription and opt to continue using the Services on a free subscription basis, you will have reduced storage capacity and may lose content and features. If you opt to do so, you agree to be solely responsible for such loss; CDS cannot and will not be held liable for such loss.

#### 6. Acceptable Use and Conduct/User Restrictions#

As a condition of your access and use of the Services, you agree that you will use the Services in compliance with these Terms of Service and all applicable laws, including any laws regarding the transmission of technical data exported from your country of residence and all export controls and embargo restrictions under the laws of the United States.

You further agree that you shall not: (a) impersonate any individual or entity or misrepresent your affiliation with any other individual or entity; (b) use the Services in any manner with the intent to interrupt, damage, disable, overburden or impair the Services; (c) use the Services in violation of CDS’ or any third party’s intellectual property or other proprietary or legal rights; (d) use the Services in violation of any applicable laws and/or to encourage illegal activities; (e) attempt (or encourage or support anyone else’s attempt) to circumvent, reverse engineer, decrypt or otherwise alter or interfere with the Services or make any unauthorized use thereof; (f) access, tamper with, or use non-public areas of the Services, CDS’ computer systems, or the technical delivery systems of CDS’ providers or otherwise obtain or attempt to obtain any content, materials or information through any means not intentionally made publicly available or provided for through the Services; (g) probe, scan, or test the vulnerability of any system or network or breach or circumvent any security or authentication measures or otherwise attempt to gain unauthorized access to the Services through hacking, password mining or any other means; (h) use or attempt to use any “spider”, “robot”, “bot”, “scraper”, “data miner” or any other program, device or algorithm, process or methodology to access, acquire, copy, or monitor the Services (or portions thereof); (i) forge any TCP/IP packet header or any part of the header information in any email or posting, or in any way use the Services to send altered, deceptive or false source-identifying information; (j) interfere with, or disrupt, (or attempt to do so), the access of any user, host or network, including, without limitation, sending a virus, overloading, flooding, spamming, mail-bombing the Services, or by taking any action that would interfere with or create an undue burden on the Services; or (k) use the Services in any way that might bring CDS into disrepute.

Any commercial or promotional distribution, publishing or exploitation of the Services is strictly prohibited unless you have received the express prior written permission from authorized personnel of CDS, CDS’ licensors or the otherwise applicable rights holders.

#### 7. Changes to and Termination of the Services#

We aim to continually improve the delivery and content of the Services and, as a result, CDS will make changes to the Services from time to time. New features may be added, but we also may modify or discontinue (temporarily or permanently) any element of the Services, in whole or in part. We will notify you if there are any material changes to the Services either via an email to the email associated with your account or via a notice displayed in the Zotero software or on the Site.

There are some circumstances under which the Services may be terminated:

A. In the event that we cannot obtain commercially-practical rates or terms from a service provider or supplier, we may cease to offer the Services. In such case, we will provide thirty-days prior notice via email to the email associated with your account or via a notice displayed in the Zotero software or on the Site.

B. We may also cease to offer the Services for any other reason, in which case we will provide you with a thirty-days prior notice via email to the email associated with your account or via a notice displayed in the Zotero software or on the Site. In such case, we will not charge you for Services after their termination, and will refund you any fees paid in advance for Services that have not been received.

CDS reserves the right, at any time, to disable the Services temporarily for security or maintenance reasons.

#### 8. Privacy#

As a non-profit enterprise intended to support the development and use of open source scholarship and cultural collaboration, CDS takes privacy very seriously. In order for you to access certain areas and functions/features of the Site and register with CDS (become a Registered User), we may ask you to establish certain login information and provide us with information that personally identifies you, such as your name and email address (“Personal Information”). If you communicate with us by email or otherwise complete online forms or the like, any information provided in such communication may be collected as Personal Information. The information collection and use policies of CDS with respect to the privacy of such Personal Information are set forth in the Site’s Privacy Policy, which is located at 0 and incorporated herein by reference.

#### 9. Copyright#

You may access third-party content in your use of the Services. You may not remove any third party’s copyright notices or other identifier, except as allowed by the third-party’s license of that content. CDS is not responsible for any content provided by any third party.

All Submissions and other data and content made available through the Services must comply with U.S. copyright law. Pursuant to the Digital Millennium Copyright Act (17 U.S.C. §512, as amended), if you believe in good faith that copyrighted work has been copied, adapted, reproduced or exhibited on the Site or through use of the Services in a manner that constitutes copyright infringement, you may submit written notification of the claimed infringing activity to our Designated Agent, Alex H. Pyle, Sheehan Phinney Bass + Green PA, 255 State Street, 5th Floor, Boston, MA 02109. To be effective, the notification of claimed infringement must include the following information:

A. A physical or electronic signature of a person authorized to act on behalf of the owner of the exclusive right that is allegedly infringed;

B. Identification of the copyrighted work claimed to have been infringed, or, if multiple copyrighted works at a single online site are covered by a single notification, a representative list of such works at that site;

C. Identification of the material that is claimed to be infringing or to be the subject of infringing activity and that is to be removed or access to which is to be disabled, and information sufficient to permit CDS to locate the material;

D. Information reasonably sufficient to permit CDS to contact the complaining party, such as an address, telephone number and, if available, an electronic mail address at which the complaining party may be contacted;

E. A statement that the complaining party has a good faith belief that use of the material in the manner complained of is not authorized by the copyright owner, its agent, or the law; and

F. A statement that the information in the notification is accurate, and under penalty of perjury, that the complaining party is authorized to act on behalf of the owner of the exclusive right that is allegedly infringed.

Please consult your legal advisor before submitting written notification, as the above-stated requirements may have changed. For further information about the DMCA, please visit the website of the U.S. Copyright Office at: 0.

In appropriate circumstances, CDS, at its sole discretion, may suspend or terminate any user’s access to the Services and/or take other action against users where infringing activity is apparent, regardless of whether the material or activity is ultimately determined to be infringing.

#### 10. Limitation of Liability#

TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, CDS AND ITS SERVICE PROVIDERS, SUPPLIERS AND LICENSORS, AND THEIR RESPECTIVE DIRECTORS, SHAREHOLDERS, OFFICERS, EMPLOYEES AND AGENTS WILL NOT BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL OR PUNITIVE DAMAGES, INCLUDING WITHOUT LIMITATION, LOSS OF PROFITS, DATA, USE, GOODWILL, OR OTHER INTANGIBLE LOSSES, RESULTING FROM (i) YOUR ACCESS TO OR USE OF OR INABILITY TO ACCESS OR USE THE SERVICES; (ii) ANY CONDUCT OR CONTENT OF ANY THIRD PARTY ON THE SITE OR THROUGH USE OF THE SERVICES, INCLUDING WITHOUT LIMITATION, ANY DEFAMATORY, OFFENSIVE OR ILLEGAL CONDUCT OF OTHER USERS OR THIRD PARTIES; (iii) ANY CONTENT OBTAINED FROM OR THROUGH THE SERVICES; AND (iv) UNAUTHORIZED ACCESS, USE OR ALTERATION OF YOUR SUBMISSIONS OR OTHER CONTENT, WHETHER BASED ON WARRANTY, CONTRACT, TORT (INCLUDING NEGLIGENCE) OR ANY OTHER LEGAL THEORY, WHETHER OR NOT CDS HAS BEEN INFORMED OF THE POSSIBILITY OF SUCH DAMAGE.

#### 11. Disclaimer of Warranties#

YOUR ACCESS TO AND USE OF THE SITE AND THE SERVICES OR ANY CONTENT IS AT YOUR OWN RISK. YOU UNDERSTAND AND AGREE THAT THE SITE AND THE SERVICES ARE PROVIDED TO YOU ON AN “AS IS” BASIS WITHOUT ANY WARRANTIES OF ANY KIND, EITHER EXPRESS OR IMPLIED. WITHOUT LIMITING THE FOREGOING, TO THE FULLEST EXTENT PERMITTED BY LAW, CDS AND ITS SERVICE PROVIDERS, SUPPLIERS AND LICENSORS, AND THEIR RESPECTIVE DIRECTORS, SHAREHOLDERS, OFFICERS, EMPLOYEES AND AGENTS, HEREBY DISCLAIM ALL EXPRESS, IMPLIED, STATUTORY AND OTHER WARRANTIES, GUARANTEES OR REPRESENTATIONS, INCLUDING, WITHOUT LIMITATION, ANY WARRANTIES OF TITLE, MERCHANTABILITY, NON-INFRINGEMENT OF THIRD PARTIES’ RIGHTS, OR FITNESS FOR PARTICULAR PURPOSE OR NEED.

WE MAKE NO WARRANTY AND DISCLAIM ALL RESPONSIBILITY AND LIABILITY FOR THE COMPLETENESS, ACCURACY, AVAILABILITY, TIMELINESS, SECURITY OR RELIABILITY OF THE SERVICES OR ANY CONTENT THEREON. CDS WILL NOT BE RESPONSIBLE OR LIABLE FOR ANY HARM TO YOUR COMPUTER SYSTEM, LOSS OF DATA, OR OTHER HARM THAT RESULTS FROM YOUR ACCESS TO OR USE OF THE SERVICES, OR ANY CONTENT MADE AVAILABLE THOROUGH THE SERVICES. YOU ALSO AGREE THAT CDS HAS NO RESPONSIBILITY OR LIABILITY FOR THE DELETION OF, OR THE FAILURE TO STORE OR TO TRANSMIT, ANY CONTENT AND OTHER COMMUNICATIONS MAINTAINED BY THE SERVICES. YOU SHOULD ENSURE THAT YOU HAVE SUITABLE BACKUPS OF ALL OF YOUR DATA. WE MAKE NO WARRANTY THAT THE SERVICES WILL MEET YOUR REQUIREMENTS OR BE AVAILABLE ON AN UNINTERRUPTED, SECURE, OR ERROR-FREE BASIS. NO ADVICE OR INFORMATION, WHETHER ORAL OR WRITTEN, OBTAINED FROM CDS OR THROUGH THE SERVICES, WILL CREATE ANY WARRANTY NOT EXPRESSLY MADE HEREIN.

#### 12. Indemnity#

You agree to defend, indemnify, and hold harmless CDS, its directors, shareholders, officers, employees, agents, licensors and service providers from and against any claims, actions or demands, including, without limitation, reasonable legal and accounting fees, arising out of or resulting from your use of the Services, your Submissions, your breach of these Terms of Service and/or violation of any applicable laws. We reserve the right, at our own expense, to assume the exclusive defense and control of any matter otherwise subject to indemnification by you, in which event you will cooperate with us in asserting any available defense.

#### 13. Links#

The Site may contain links to third-party websites or resources. These links are provided solely as a convenience to you. You acknowledge and agree that we are not responsible or liable for: (i) the availability or accuracy of such websites or resources; or (ii) the content, products, or services on or available from such websites or resources. Links to such websites or resources do not imply any endorsement by CDS of such websites or resources or the content, products, or services available from such websites or resources. Nor do links to such websites and resources imply that CDS is affiliated or associated with the linked-third-party site. You acknowledge sole responsibility for and assume all risk arising from your use of any such websites or resources.

#### 14. Local Laws and Export Control#

CDS, through the Site, provides services and uses software and technology that may be subject to United States export controls administered by the U.S. Department of Commerce, the United States Department of Treasury Office of Foreign Assets Control and other U.S. agencies. You acknowledge and agree that the Site shall not be used, and none of the underlying information, software or technology may be transferred or otherwise exported or re-exported to countries as to which the United States maintains an embargo (collectively, “Embargoed Countries”), or to or by a national or resident thereof, or any person or entity on the U.S. Department of Treasury’s List of Specially Designated Nationals or the U.S. Department of Commerce’s Table of Denial Orders (collectively, “Designated Nationals”). The lists of Embargoed Countries and Designated Nationals are subject to change without notice. By using the Services, you represent and warrant that you are not located in, under the control of, or a national or resident of an Embargoed Country or Designated National. You agree to comply strictly with all United States export laws and assume sole responsibility for obtaining licenses to export or re-export as may be required.

CDS and its licensors make no representation that the Services are appropriate or available for use in other locations outside the United States. If you use the Services from outside the United States, you are solely responsible for compliance with all applicable laws, including without limitation export and import regulations and intellectual property laws of other countries. Any diversion of the Services and/or any content obtained from or through the Services contrary to United States laws is prohibited.

#### 15. Controlling Law and Jurisdiction#

These Terms of Service and any action related thereto will be governed by the laws of Virginia without regard to or application of its conflict of law provisions or your state or country of residence. All claims, legal proceedings or litigation arising in connection with the Services will be brought solely in Virginia, and you consent to the jurisdiction of and venue in such courts and waive any objection as to inconvenient forum. If you are accepting these Terms of Service on behalf of a United States federal government entity that is legally unable to accept the controlling law, jurisdiction or venue clauses above, then those clauses do not apply to you but instead these Terms of Service and any action related thereto will be will be governed by the laws of the United States of America (without reference to conflict of laws) and, in the absence of federal law and to the extent permitted under federal law, the laws of Virginia (excluding choice of law).

#### 16. General#

CDS makes no claims that the materials provided on or accessed through the Site are appropriate or may be downloaded outside of the United States. Access to the Site or Services may not be legal by certain persons or in certain countries. If you access the Site or Services from outside of the United States you do so at your own risk and are responsible for compliance with all applicable laws. If any provision contained herein is found to be invalid by any court having competent jurisdiction, the invalidity of such provision shall not affect the validity of the remaining provisions set forth herein, which shall remain in full force and effect. No waiver of any term hereunder of these Terms of Service shall be deemed a further or continuing waiver of such term or any other term. These Terms of Service constitute the entire agreement between you and CDS with respect to the use of the Site. Any changes to these Terms of Service must be made in writing, signed by an authorized representative of CDS to be binding on CDS. Notwithstanding the foregoing, CDS, at its sole discretion and without notice, may change, modify, add or remove any portion of these Terms of Service, in whole or in part, at any time. Changes in these Terms of Service will be effective when posted. If the changes, in our sole discretion, are material, we will notify you via an email to the email associated with your account or via a notice displayed in the software or on the Site. Your continued use of the Site and/or the Services made available on or accessed through the Site after any changes to these Terms of Service are posted will be considered acceptance of those changes.

Except as otherwise provided in these Terms of Service, all notices should be sent by registered mail to the Corporation for Digital Scholarship, 7245 Arlington Blvd., Suite 300, Falls Church, VA 22042, USA.

Effective Date: October 7, 2014

### Zotero Wiki Contributor License Agreement {#zotero-wiki-contributor-license-agreement}
#### Zotero Wiki Contributor License Agreement#

Due to [ambiguity](https://groups.google.com/d/topic/zotero-dev/YUzcpS-Ffms/discussion) concerning the license status of Zotero wiki content prior to April 27, 2015, wiki contributions prior to that date may be understood to be licensed under a CC Attribution-Noncommercial-Share Alike 3.0 Unported (CC-BY-NC-SA) license and therefore unusable for commercial purposes, even if those purposes benefit the Zotero community.

By adding your username below, you agree to license your contributions to this wiki prior to April 27, 2015, under a [Creative Commons Attribution-ShareAlike 4.0 International License](http://creativecommons.org/licenses/by-sa/4.0/):

-   br
-   dstillman
-   fcheslack
-   rmzelle
-   sean
-   zuphilip
-   adamsmith

### “[domain] uses an invalid security certificate. The certificate is not trusted because […]” {#domain-uses-an-invalid-security-certificate-the-certificate-is-not-trusted-becau}
#### “\[domain\] uses an invalid security certificate. The certificate is not trusted because \[…\]”#

If a Zotero error report shows an error similar to the above for your institution’s **proxy** or **WebDAV server** or a site you’re trying to save from, there are two possibilities:

1.  You’re connecting to a server with a “self-signed certificate”. For a proxy or WebDAV server, you would need to whitelist the certificate in Zotero. This is uncommon for public servers.
2.  The server is misconfigured and will need to be fixed by your IT department or the IT department of the site operator. See the technical details below for more information.

If you’re using an institutional proxy or WebDAV server and are unsure which is the case, point your IT department to this page along with the URL from the error report.

If you’re getting a certificate error for a zotero.org or s3.amazonaws.com URL — for example, while syncing — that’s a different issue.

##### Technical Details: Missing Intermediate Certificate#

If the server isn’t using a self-signed certificate (i.e., if it’s chained to a root certificate that’s trusted in browser stores), this error generally occurs because the server isn’t serving the necessary “intermediate certificate” for secure connections, and Zotero (like Firefox, on which it is based) won’t download it on its own. Without an intermediate certificate, it’s impossible to determine whether the connection is secure, and the connection fails.

To verify that this is the case, submit the URL from the error report to the [SSL Labs server test](https://www.ssllabs.com/ssltest/) and view the results. If you see “Chain issues: Incomplete” in orange under “Additional Certificates (if supplied)”, you’re experiencing this issue. The report will then also say “Extra download” (instead of “Sent by server” or “In trust store”) for one or more certificates listed under “Certification Paths”. Alternatively, one or more bundled intermediate certificates may be listed as expired. The missing intermediate certificate(s) should be provided along with the site’s primary certificate when HTTPS clients connect.

Note that loading the same HTTPS URL in a browser may still work. In that case, either the browser is downloading intermediate certificates automatically (as Chrome does) or you previously loaded another site (perhaps even another from your institution) that included the intermediate certificate, which the browser cached and is using even on sites that don’t serve it properly. Sites should always serve their intermediate certificates, however, and are misconfigured if they don’t. If you create a new profile in Firefox, you should get a certificate error trying to load the same URL, which is essentially the situation Zotero is in.

### “The add-on could not be downloaded because of a connection failure on www.zotero.org.” {#the-add-on-could-not-be-downloaded-because-of-a-connection-failure-on-wwwzoteroo}

**This article applies to the deprecated Zotero for Firefox (pre-Zotero 5.0) plugin. It no longer applies to the current versions of Zotero.**

#### “The add-on could not be downloaded because of a connection failure on www.zotero.org.”#

This message previously could appear when attempting to download Zotero Firefox extensions on systems or networks that intercepted secure (HTTPS) connections to zotero.org. We’ve since put a workaround in place that allows installations to take place over such connections.

Automatic updates will continue not to work over such connections, however. See Updates Not Detected for more details.

## Preferences and troubleshooting {#preferences-and-troubleshooting}
### Advanced {#advanced}

The Advanced pane has four tabs: “General”, “Files and Folders”, “Shortcuts”, and “Feeds”.

#### General#
##### Miscellaneous#

-   **Automatically check for updated translators:** Allow Zotero to automatically update its translators (for detecting and saving bibliographic data from different websites) and citation styles. You manually check for updates immediately by clicking the “Update now” button.
-   **Report broken site translators:** Allow Zotero to notify its developers when a translator fails to save an item. This information is submitted anonymously.

##### Language#

Set the language for the Zotero interface.

##### OpenURL#

Here you can specify a different OpenURL resolver for use with Zotero’s Library Lookup feature.

If you are on an institution network, you can click the “Search for resolvers” button. If Zotero finds an OpenURL resolver belonging to your institution, you can select it using the “Custom…” drop-down menu.

You can also enter an OpenURL URL by hand in the Resolver field. Most resolvers use OpenURL version 1.0, but 0.1 is still in use. Ask your librarian for more information, or check our own list of OpenURL resolvers.

##### Advanced Configuration#

Click the “Config Editor” button to configure Zotero’s hidden preferences.

#### Files and Folders#
##### Linked Attachment Base Directory#

*If you store attached files in Zotero — the default — this setting does not affect you. It only applies to* linked *files.*

This setting allows you to access linked files on multiple computers even when they’re stored in different locations on each computer. You should set the base directory to the folder on each computer under which you store linked files. For example, if the folder with your linked files is at 0 on your laptop and at 1 on your work computer, set those paths as the base directory on each respective machine. If you add a linked file within the base directory, Zotero stores a relative path to that base directory rather than an absolute path.

Note that this setting does not control where files are stored — only whether linked files within the specified folder are referenced by absolute or relative paths. If you’re using a plugin to help with a linked-file workflow, you should configure it to store linked files within the base directory you’ve configured.

Note that linked files are an advanced configuration, and we’re not able to help troubleshoot problems with specific setups. We recommend that most people use stored files instead.

##### Automatic Linked-File Relinking#

If you previously added linked files on another computer without having set a base directory and your files are now in a different location on your current computer, you can simply set your base directory to the desired containing folder on your current computer (e.g., 0). When you next open a missing file with an absolute path (e.g., 1), Zotero will automatically locate the file within your specified base directory and offer to relink all files that have the same base path (2).

##### Data Directory Location#

By default, Zotero stores your data directory (which contains your library database, attachment files, and several other files) in your user home folder on your computer. This is the best location for most users, but it is possible to change the location. After changing the data directory location, Zotero will store newly created data in the new location.

Note that Zotero will not, however, copy over existing data to the new location. If you want to keep your data, you need to move the files manually. Click the “Show Data Directory” button to open the Zotero folder in your computer’s file browser. You will need to move everything in the folder (including ‘zotero.sqlite’, ‘storage’, ‘styles’, and all other files) to the new location.

##### Unsafe Data Directory Locations#

There are several data directory locations that are **likely to lead to database corruption or even data loss**.

-   **Cloud storage folders**: Storing your Zotero data directory in a cloud storage folder (e.g., Dropbox, Google Drive, or other similar folder synchronization services) is extremely unsafe and will almost certainly lead to database corruption and potentially data loss.
-   **Network drives**: If you store your Zotero data directory on a network drive and then access it from multiple computers at the same time, you are very likely to encounter database corruption. For example, if you leave Zotero open on your laptop, then open the same Zotero database on the network drive from your work computer, you will likely experience corruption. You should never use a network drive to permit multiple users to access the same Zotero database on different machines (use groups or Zotero syncing for that).
-   **Virtual machines**: Similar to issues with network drives, it can be unsafe to use the same Zotero database file from both a virtual machine and the computer’s host operating system (or another virtual machine). If the same Zotero database is accessed from two locations at the same time (e.g., if Zotero is open on both the virtual machine and the host operating system), corruption is likely. If you want to use Zotero in a virtual machine, it is better to set up a separate Zotero data directory in the virtual machine and to keep it up to date using Zotero syncing.
-   **External disks**: While keeping your data directory on an external disk is usually safe, your database could become corrupted if, say, the disk becomes unmounted while Zotero is open.

##### Database Maintenance#

-   **Check Database Integrity:** This function checks your Zotero database for invalid data and database corruption. Database corruption is rare, and in most cases is caused by storing your data directory in an unsafe location. Checking database integrity can take a long time if your database is very large. If your database is corrupted, you can try to use the Database Repair Tool to repair the corruption.
-   **Reset Translators…:** Reset web and import/export translators to the versions bundled with the application or provided as updates from Zotero servers.
-   **Reset Styles…:** Reset citation styles to the versions bundled with the application or provided as updates from Zotero servers.

#### Shortcuts#

This tab allows you to change Zotero’s default keyboard shortcuts.

-   **Create a New Item:** Create a new blank item in the current collection.
-   **Create a New Note:** Create a new standalone note in the current collection.
-   **Focus Libraries Pane:** Sets the focus in Zotero to left (libraries, collections, and feeds) pane.
-   **Quick Search:** Sets the focus to the Quick search box. 0-1 will also focus the Quick search box.
-   **Copy Selected Item Citations to Clipboard:** Copies an inline citation for the selected item(s) to the clipboard. (Depending on the style, this could be long and detailed or, if the style demands footnotes, simply a number.)
-   **Copy Selected Items to Clipboard:** Copies the full bibliographic reference for the selected item(s) to the clipboard.
-   **Toggle Tag Selector:** Shows/hides the tag selector.
-   **Mark All Feed Items As Read/Unread:** Marks all items in the selected feed as read/unread.

Any changes made to this page will only take effect after you restart Zotero. a new browser window is opened.

##### Windows Defaults#

- Function: Create a New Item; Command: Ctrl-Shift-N
- Function: Create a New Note; Command: Ctrl-Shift-O
- Function: Focus Libraries Pane; Command: Ctrl-Shift-L
- Function: Quick Search; Command: Ctrl-Shift-K
- Function: Copy Selected Item Citations to Clipboard; Command: Ctrl-Shift-A
- Function: Copy Selected Items to Clipboard; Command: Ctrl-Shift-C
- Function: Toggle Tag Selector; Command: Ctrl-Shift-T
- Function: Mark All Feed Items As Read/Unread; Command: Ctrl-Shift-R


##### Mac OS X Defaults#

- Function: Create a New Item; Command: Cmd-Shift-N
- Function: Create a New Note; Command: Cmd-Shift-O
- Function: Focus Libraries Pane; Command: Cmd-Shift-L
- Function: Quick Search; Command: Cmd-Shift-K
- Function: Copy Selected Item Citations to Clipboard; Command: Cmd-Shift-A
- Function: Copy Selected Items to Clipboard; Command: Cmd-Shift-C
- Function: Toggle Tag Selector; Command: Cmd-Shift-T
- Function: Mark All Feed Items As Read/Unread; Command: Cmd-Shift-R


#### Feeds#

This tab contains preferences for Zotero’s Feeds feature.

-   **Sorting:** Change between sorting newest or oldest items first.
-   **Feed Defaults:** Change how frequently feeds are checked for new items and how long read and unread items are kept in your database before being removed.

### Cite {#cite}

The Cite pane has two tabs: “Styles” and “Word Processors”.

#### Styles#
##### Style Manager#

The Style Manager displays the currently installed citation styles and the date they were last updated. You can download additional styles directly from the Zotero Style Repository by clicking the “Get additional styles…” link. You can also install a local [Citation Style Language](http://citationstyles.org/) (CSL) style file by clicking the “+” button and locating the style file on your computer. To delete a style, select the style and click the “-” button.

If you aren’t sure what style you need, you can [Search by Example](http://editor.citationstyles.org/searchByExample/) to find a style. Note that this tool requires you to format the reference data shown on the page, not just any example reference.

##### Citation Options#

-   **Include URLs of paper articles in references:** When this option is unchecked, Zotero will only include a URL when citing journal, magazine, and newspaper articles when the article does not have a page range specified.

##### Tools#

-   **Style Editor:** Opens Zotero’s CSL Editor window for editing and testing CSL citation styles. You can also edit styles using a plain text editor program on your computer (e.g., Notepad, Text Edit) or the [CSL Visual Editor](http://editor.citationstyles.org/visualEditor/).
-   **Style Preview:** Opens Zotero’s CSL Preview window to preview how the references selected in your library will be formatted using installed styles.

#### Word Processors#

Zotero’s word processor plugins integrate Zotero into either Microsoft Word or LibreOffice. Zotero will install plugins into Word and LibreOffice automatically when you install Zotero. You can re-install the word processor plugins from this pane. Reinstalling the LibreOffice plugin may be helpful if upgrading LibreOffice causes the path to the LibreOffice program files to change.

If one of the Install buttons is disabled, check that the respective word processor extension is installed and enabled in Zotero Add-ons window (click “Tools -> Add-ons”).

-   **Use classic Add Citation dialog:** By default, the Zotero word processor plugins will use a Quick Citation interace that lets you intuitively search for items across all of your libraries, add multiple items to the same citation, and easily add page numbers, prefixes, and suffixes. To browse your libraries and collections for an item, you can click the “Z” on the left of the window to switch to the “Classic View” citation dialog. Check this option to switch the default interface to the Classic View. Note that some features are not supported in the Classic View.

### Debug Output Logging {#debug-output-logging}

If you’ve been asked to provide a Debug ID (which is different from a Report ID) to help troubleshoot a problem, follow these simple steps:

#### Zotero#

1.  In the Help menu, go to Debug Output Logging and select Enable, or, to generate debug output from Zotero startup, select “Restart with Logging Enabled…”. (If you’re not able to access the Help menu, see Reporting Startup Errors instead.)
2.  Immediately perform the relevant action (syncing, saving, importing, etc.) and reproduce the problem you’re experiencing.
3.  Before doing anything else, return to Help -> Debug Output Logging and click Submit Output, which will disable logging and submit the output to zotero.org. A window should pop up containing a Debug ID (e.g., “D12345678”). Click “Copy to Clipboard” and paste the Debug ID into your forum thread.

If submitting output fails, you can return to the Debug Output Logging menu and select View Output, go to File -> “Save…”, choose Format: “Text Files”, and save the output to a file, which you can email to support@zotero.org with a link to your forum thread. It can be helpful to ZIP the file before emailing it.

#### Zotero Connectors (Firefox, Chrome, and Safari)#

1.  Open the Zotero Connector preferences
    -   **Firefox:** right-click on the Zotero Connector extension button and click “Preferences”.
    -   **Chrome:** right-click on the Zotero Connector extension button and click “Options”.
    -   **Safari:** right-click anywhere on a webpage and select “Zotero Preferences…”
2.  In the Advanced tab of the Zotero Connector preferences, under “Debug Output Logging”, check the box next to “Enable Logging”. Do not close this tab.
3.  Immediately perform all the relevant actions (e.g., import an item from a web page).
4.  Go back to the Advanced tab of the Zotero Connector preferences and click Submit Output.
5.  You will be provided with a Debug ID (e.g., “D12345678”). Please post the Debug ID to the forums.
6.  Uncheck the box next to “Enable Logging.”

#### Zotero for iOS#

1.  Tap Back in the top-left corner until you see the list of libraries.
2.  Tap the Gear icon.
3.  Tap “Debug Output Logging” and then “Start Logging”.
4.  Close the settings window.
5.  Immediately perform all relevant actions (e.g., pulling down on the items list to trigger a sync or switching to the browser and trying to save).
6.  When you’re done, tap the circular stop button in the bottom-left corner of the screen.
7.  In the alert that pops up, tap Copy to copy the Debug ID to the clipboard, and then paste it into your forum thread.

#### Zotero for Android#

1.  From the items list, tap Collections in the top-left corner and then Libraries.
2.  Tap the Gear icon.
3.  Tap “Debug Output Logging” and then “Start Logging”.
4.  Close the settings window.
5.  Immediately perform all relevant actions (e.g., pulling down on the items list to trigger a sync or switching to the browser and trying to save).
6.  When you’re done, tap the circular stop button in the bottom-left corner of the screen.
7.  In the alert that pops up, tap Copy to copy the Debug ID to the clipboard, and then paste it into your forum thread.

#### Logging to a Terminal Window#

If you’d like to regularly follow Zotero’s debug output in real-time, it may be preferable to have Zotero log to a terminal window.

##### macOS#

-   Open Terminal via Spotlight or from /Applications/Utilities.
-   Go to the Terminal menu and open Settings. In Profiles → Window, make sure Scrollback is set to “Limit to available memory”.
-   Paste 0 into the Terminal window.
-   Press Return

You can add 0 to the end of the command to redirect the output to a file on your desktop.

##### Windows#

-   Open cmd.exe, Cygwin shell, or another terminal
-   Paste 0 into the console window. It may be necessary to add 1 as well.
-   Press Enter.

Due to limitations of the available console windows on Windows, you may have a better experience using 0 instead to use Zotero’s internal debug output window.

##### Linux#

Start Zotero via the command line, adding the 0 command-line flag.

##### Logging to a File#

To capture output when Zotero is crashing or hanging, you can use 0 to redirect output to a file.

##### Developer Note#

To enable Zotero debug output permanently, set extensions.zotero.debug.log to true in the Zotero config editor, accessible from the Advanced pane of the Zotero preferences, and then start Zotero from the command line.

### Export {#export}

The Export preferences pane is used to choose a default citation style or bibliographic data format and configure other settings for copying data from Zotero.

For information on adding and removing citation styles from Zotero, see Citation Styles and the Cite pane in Zotero preferences.

#### Quick Copy#

Quick Copy allows you to quickly export items in the specified format. You can set your Quick Copy **Default Format** to either a citation style, such as Chicago Manual of Style (note), or a bibliographic export format, such as BibLaTeX or Zotero RDF. Press Ctrl/Cmd-Shift-C to copy the full reference to the clipboard. You can also drag the item to any text box in another program. If you set your default format to a citation style, press Ctrl/Cmd-Shift-A or hold down Shfit before dragging to copy the in-text ciation or footnote.

Besides setting the default Quick Copy format, you can also change these Quick Copy settings:

-   **Language:** What languages should be used for Quick Copy bibliographies and citations?
-   **Copy as HTML:** Quick Copy citations as regular text (default, generally recommended) or HTML (useful for pasting in web pages).
-   **Site-Specific Settings:** These settings can change the default Quick Copy format when a specific website is open in your web browswer with the Zotero Connector plugin installed. For example, you can set the default export style to Wikipedia Citation Templates when copying references to insert into Wikipedia articles. This feature makes it easy to use different citation styles and formats across contexts.
    -   See Zotero and Wikipedia for more information.
    -   You can also assign multiple citation styles or export formats to different keyboard shortcuts using the [Zutilo plugin](https://github.com/willsALMANJ/Zutilo/releases).
-   **Disable Quick Copy when dragging more than … items:** Setting this option (default: 50) can prevent Zotero from slowing down when draggin a very large number of items.

#### Character Encoding#

-   **Import Character Encoding:** By default, Zotero auto-detects the proper character encoding when you import bibliographic data. If this doesn’t work correctly for a file, you can specify the correct character encoding here.

### General {#general}

The General preferences pane controls Zotero’s user interface, item import settings, file handling, and group library behavior.

#### File Handling#

-   **Automatically take snapshots when creating items from web pages:** When importing items from websites or archiving a web page using the Zotero save button in your browser, should Zotero save a snapshot of the web page as an attachment to the new item (enabled by default)?
-   **Automatically attach associated PDFs and other files when saving items:** Should Zotero automatically download the full-text PDF version of articles (or sometimes other related files) when importing items using the Zotero save button in your browser (enabled by default)?
    -   To automatically download supplementary material, as well as the main article file, see Supplemental Material.

#### Miscellaneous#

-   **Automatically tag items with keywords and subject headings:** Some repositories use tags to annotate and organize items (examples are the Library of Congress catalog, which includes subject headings, and the online version of the New York Times, which uses keywords). When this option checked (the default), these annotations are attached to saved items as automatic tags.
-   **Automatically remove items in the trash deleted more than … days ago:** Change how long items are kept in the trash before being automatically deleted (default: 30 days).

#### Groups#

By default, when you copy items between your personal My Library and a group library (or between different group libraries), any child notes, snapshots and other stored file attachments, linked URIs (to web pages or other programs), and tags are copied along. You can uncheck “child notes”, “child snapshots and imported files”, “child links”, or “tags” to prevent notes, attachments, links, or tags from being copied. Note that linked attachments are not supported in group libraries and will not be copied.

### Hidden Preferences {#hidden-preferences}

You can edit most Zotero preferences through the Preferences window in Zotero or the Preferences pane in the Zotero Connector in your browser. However, both Zotero and the Zotero Connector support additional hidden preferences. These settings may have received less testing and/or are intended for more advanced use.

#### Zotero#

To view the the full list of Zotero’s preferences, including many hidden preferences, go to the Advanced pane of the Zotero preferences and click “Config Editor”. Enter “zotero” into the Filter field at the top of the list that comes up. Preferences that can be safely changed by users are described below.

Most Zotero hidden preferences are preceded by “extensions.zotero.”

##### General Preferences#

These general hidden preferences allow you to refine your Zotero configuration.

- Preference Name: backup.interval; Default Value: 1440; Description: Determines, at most, how often (in minutes) Zotero makes an automatic backup of the database. The default is every 24 hours (1440 minutes)
- Preference Name: backup.numBackups; Default Value: 2; Description: Determines how many automatic database backups Zotero should keep. Excess backups are deleted oldest first. This does not include backups made during database upgrades.
- Preference Name: capitalizeTitles; Default Value: true; Description: By default, Zotero will recase titles of items you capture (e.g., to remove all caps). Switch this preference to false and you will preserve case information for titles.
- Preference Name: debug.level; Default Value: 5; Description: When debug.log is enabled, determines the lowest of the debug levels (1-5, with 5 being the lowest) that is displayed
- Preference Name: debug.log; Default Value: false; Description: Used for debugging Zotero. See debug output.
- Preference Name: debug.time; Default Value: false; Description: When debug.log is enabled, shows the milliseconds from the previous debug call
- Preference Name: fontSize; Default Value: “1.0”; Description: This preference allows you to increase or decrease the size of text in the Zotero interface.
- Preference Name: httpServer.enabled; Default Value: true; Description: If set to true, Zotero will listen for requests from the Zotero Connector (e.g., to allow saving items to Zotero from the Connector).
- Preference Name: httpServer.port; Default Value: 23119; Description: If 0 is enabled, this is the port on which Zotero will listen for connections from the Zotero Connector.
- Preference Name: sortAttachmentsChronologically; Default Value: false; Description: If set to true, your attachments will be sorted by the order you added them instead of alphabetically.
- Preference Name: sortNotesChronologically; Default Value: false; Description: If set to true, your notes will be sorted by the order you added them instead of alphabetically.


##### PDF Reader#

- Preference Name: sortNotesChronologically.reader; Default Value: true; Description: Sort item notes in reverse chronological order. If 0, sort alphabetically.


##### Note Editor#

- Preference Name: note.css; Description: Custom CSS to apply to note content
- Preference Name: note.fontSize; Default Value: 14; Description: Note font size — settable from the View menu, but other values (including decimals) can be set manually
- Preference Name: note.smartQuotes; Default Value: true; Description: Automatically convert straight quotes to typographic quotes


##### Translator Preferences#

These hidden preferences allow you to control behavior for import/export translators for some specific bibliographic formats. **All translator hidden preferences are preceded by “extensions.zotero.translators.”**

[TABLE]

##### Full-Text Indexing#

These preferences deal with Zotero’s ability to create full-text indexes from imported files.

- Preference Name: search.useLeftBound; Default Value: true; Description: This preference determines whether Zotero only finds word matches based on the left bound or whether it finds matches anywhere within words. Switching this to false may be beneficial for languages other than English, but it may significantly slow down Zotero’s search functions.


##### Reports#

These options allow you to customize your reports.

- Preference Name: report.includeAllChildItems; Default Value: true; Description: By default, selecting only parent items for a report causes those items’ child notes and attachments to be included as well. If includeAllChildItems is set to false, only the items you have selected will be included. Selecting a combination of parent and child items will cause only the selected items to be displayed regardless of this setting.
- Preference Name: report.combineChildItems; Default Value: true; Description: By default, Zotero groups child notes and attachments in reports together under their parent items. Switching this to false will cause notes to appear separately from their parent items. This can be helpful for people interested in using Zotero’s note-taking features as an outlining tool.


##### Citation QuickCopy Settings#

- Preference Name: export.quickCopy.compatibility.indentBlockquotes; Default Value: true; Description: Word and TextEdit don’t indent blockquotes on their own and need this enabled. Results in an extra indent in LibreOffice, which handles blockquotes correctly.
- Preference Name: export.quickCopy.compatibility.word; Default Value: false; Description: Add Word Normal style to paragraphs and enable double-spacing. LibreOffice inserts the conditional style code as a document comment.
- Preference Name: quickCopy.quoteBlockquotes.plainText; Default Value: true; Description: Add quotes around blockquote paragraphs in plain-text output
- Preference Name: quickCopy.quoteBlockquotes.richText; Default Value: true; Description: Add quotes around blockquote paragraphs in rich-text output


##### Word Processor Plugin#

- Preference Name: integration.keepAddCitationDialogRaised; Default Value: false; Description: If you switch this to true, you can keep the Zotero word plugin interface for adding citations always at the front. and prevent it from going hidden behind the Word window you’re working with.


#### Zotero Connector#

To view hidden preferences for the Zotero Connector, open the preferences for the connector (by right-clicking on the save button and choosing Preferences/Options in Chrome and Firefox, or by long-pressing the save button in Safari). Then, click “Advanced”, then “Config Editor”.

##### Translator Preferences#

Zotero Connectors support some translator preferences that apply to all translators generally or to specific websites. To use these preferences, in the Zotero Connector Config Editor, click “Add Preference”. Type or paste the preference’s name and click “OK”. Enter the appropriate preference value from the table below (e.g., **true** or **1**) and click “OK” again.

- Preference Name: translators.attachSupplementary; Default Value: false; Description: Translators should attempt to attach supplementary data when importing items.; Applies to: All web translators implementing this behavior
- Preference Name: translators.supplementaryAsLink; Default Value: false; Description: Supplementary data attachments should be attached as links instead of being downloaded. This option has no effect if attachSupplementary is disabled. Setting this oprtion to “true” maintains the convenience of quick access to supplementary data, but speeds up saving items from the web.; Applies to: All web translators implementing this behavior
- Preference Name: translators.ACS.highResPDF; Default Value: 0; Description: Determines which version of the Full Text PDF is attached: 0 - PDF w/ links; 1 - high res PDF; 2 - both; Applies to: ACS Publications


**Note:** The supplementary data preferences will only work for sites whose translator supports this behavior. If you encounter sites with supplementary data that are not imported, please report it on the Zotero forums.

### How can I use multiple profiles in Zotero? {#how-can-i-use-multiple-profiles-in-zotero}

Zotero allows you to create multiple profiles, each with its own settings and associated data directory. This is an advanced configuration and not recommended for most users, but if you’re familiar with using multiple profiles in Firefox, Zotero works the same way and supports the same command-line flags.

A default profile is created when you first start Zotero. To create an additional profile, start Zotero from the command line and pass the 0 flag to open the Profile Manager:

#### macOS#

-   Open Terminal via Spotlight or /Applications/Utilities.
-   Paste 0 into the Terminal window.
-   Press Return

#### Windows#

-   Open the Run dialog (Search/Cortana → type “Run” → Run (Windows 10) or Start → Run (Windows 7)
-   Paste 0
-   Press Enter

#### Linux#

-   Start Zotero via the command line, adding the 0 command-line flag.

The Profile Manager window should appear, allow you to select, create, and delete Zotero profiles.

When you create a new profile (e.g., “Work”), if there’s already a profile pointing to the default data directory location, Zotero will create a new data directory named after the new profile (e.g., “Zotero Work”) when you first start it. Your original data directory won’t be affected. \[**Note, June 2023:** This seems to currently not always happen. If you see your existing data directory in the new profile, create a new data directory manually and point the new profile to there from the Advanced → Files and Folders pane of the Zotero preferences.\]

You can open a specific profile from the command line with the 0 flag (e.g., 1), which may be useful for creating shortcuts that automatically open a given profile. (On a Mac, you can save [an AppleScript with command-line flags embedded](https://superuser.com/a/116237) as an application in Script Editor.)

Additional configuration may be required when running multiple instances of Zotero at a time.

### How do I change Zotero settings? {#how-do-i-change-zotero-settings}
#### How do I change Zotero settings?#

The Zotero Settings menu is located in the “Zotero” menu on Mac and in the “File” menu on Windows and Linux.

### I upgraded to Zotero 5.0 and now my data is missing! How do I get it back? {#i-upgraded-to-zotero-50-and-now-my-data-is-missing-how-do-i-get-it-back}
#### I upgraded to Zotero 5.0 and now my data is missing! How do I get it back?#

Zotero 4.0 and earlier stored data by default in a ‘zotero’ directory within either the Zotero profile directory or the [Firefox profile directory](https://support.mozilla.com/kb/Profiles).

When you first start Zotero 5.0, it attempts to migrate your data directory to a new default location, a “Zotero” directory in your home directory. (See Zotero Data for the default location on your platform.) If you previously used both Zotero for Firefox and Zotero Standalone and chose not to share a data directory, you may have a Zotero database in both profile directories, and Zotero will automatically use whichever database was modified more recently.

In some cases, Zotero may not copy the correct directory and instead either create a new data directory or migrate the data directory from the wrong profile, potentially leaving you with an empty or outdated database.

If Zotero 5.0 then syncs with your online library, it will update that database, but if your online library is out of date — for example, if you haven’t synced in several months — you could still end up with missing local data.

#### Restoring Your Data#

If you used Zotero for Firefox at any point in the past, your data may be stored in the Firefox profile directory, and you may need to temporarily disable security software so that Zotero can access that directory on startup and offer to restore your data.

Otherwise, to restore your data manually, first identify your current Zotero data directory from the Advanced → Files and Folders pane of the Zotero preferences. (Again, this will generally be “Zotero” within your home directory, but it could also be ‘zotero’ within your Zotero profile directory.)

Then find your Zotero profile directory and/or [Firefox profile directory](https://support.mozilla.com/kb/Profiles) and look for a ‘zotero’ subdirectory. If you find one, check the timestamp of the zotero.sqlite file and confirm that there’s a ‘storage’ directory with other directories below it, corresponding to the range of dates that you added items to Zotero. (These are your attachment files.)

If you ever used Firefox’s “Refresh Firefox” feature, you can also look for your Zotero data in a ‘zotero’ folder folder located in the “Old Firefox Data” folder on your desktop.

To restore the data directory you located, follow these steps:

1.  Close Zotero
2.  Rename your current data directory in your home directory from “Zotero” to “Zotero-Old”
3.  Move the ‘zotero’ directory from the profile directory (or from “Old Firefox Data”) to your home directory and rename it to “Zotero” (with a capital “Z”).
4.  When you then start up Zotero 5.0 again, you should see your previous data.

If there isn’t a ‘zotero’ directory within either profile directory or in the “Old Firefox Data” folder on your desktop, it’s possible you were previously using a custom data directory location, which Zotero 5.0 also wouldn’t be able to locate automatically if it was blocked from accessing the Firefox preferences file. If you know the location, you can simply point Zotero to it from the Advanced → Files and Folders pane of the Zotero preferences. Otherwise, open the prefs.js file in your Firefox profile in a text editor and search for 0 to see if a custom path was set.

#### Further Assistance#

If you’re still having trouble locating your previous data, post to the Zotero forums with the following information:

-   Your current data directory location from the Advanced → Files and Folders pane of the Zotero preferences
-   The timestamp and sizes of zotero.sqlite\* files within that directory
-   Which version of Zotero you were using previously
-   The names of a few files you see in your Firefox profile directory

### Overriding Security Certificate Errors in Zotero {#overriding-security-certificate-errors-in-zotero}

**Note:** These instructions are only for use with security software that intercepts/scans HTTPS connections, a WebDAV server with a self-signed certificate, or an institutional network that monitors encrypted traffic using a custom root certificate authority (CA). You should never override certificate errors unless you understand the consequences. When in doubt, please contact your network administrator or ISP.

#### Self-Signed Certificate#

Zotero does not currently provide a graphical way to whitelist self-signed certificates, so you will need to copy files from a working Firefox installation.

If you are using a WebDAV server with a self-signed certificate, you can open the WebDAV URL in Firefox, accept the certificate, and then copy the cert\_override.txt file from the [Firefox profile directory](http://support.mozilla.com/kb/Profiles) to the Zotero profile directory.

##### Zotero 8.0#

Zotero 8.0 can read a cert\_override.txt file from [Firefox 140 ESR](https://ftp.mozilla.org/pub/firefox/releases/140.7.0esr/). A file from a later version of Firefox may or may not work.

##### Zotero 7.0#

Zotero 7.0 can read a cert\_override.txt file from [Firefox 115 ESR](https://ftp.mozilla.org/pub/firefox/releases/115.29.0esr/). A file from a later version of Firefox may or may not work.

##### Zotero 6#

Zotero 6 expects a cert\_override.txt file created by [Firefox 60 ESR](https://ftp.mozilla.org/pub/firefox/releases/60.9.0esr/), with a line in this form:

    192.168.xxx.xxx:1234    OID.2.16…    1D:E4:07:…    U    AAAA…

If you create an override file with a newer version of Firefox, your cert\_override.txt file may contain a line with a trailing colon after the port number (“1234” in this example) and may be missing one or more letters before “AAAA” (“U” in the above example):

    192.168.xxx.xxx:1234:    OID.2.16…    1D:E4:07:…    AAAA…

To use such a file in Zotero 6, strip the colon from after the port number and add a “U” (untrusted cert) before “AAAA”. To allow for a hostname mismatch, add “M”.

#### Custom Certificate Authority#

If you or your organization is using a custom certificate authority, which can be the case when using security software or connecting via a proxy server, Zotero may need to be configured to accept the custom CA:

-   **Windows/Mac:** Zotero 7 will automatically use the system root certificate store, which in most cases should allow it to work automatically like other browsers on the system.
-   **Linux**: Zotero is based on Firefox and uses the same certificate mechanism, so you or your IT department will need to configure Firefox for the custom CA in a new Firefox 115 ESR profile and then copy the cert9.db, key4.db, and pkcs11.txt files from the [Firefox profile directory](http://support.mozilla.com/kb/Profiles) to the Zotero profile directory.

### Profile directory location {#profile-directory-location}
#### Profile directory location#
#### Zotero#

Zotero’s profile system is based on the [Firefox profile system](https://support.mozilla.org/kb/profiles-where-firefox-stores-user-data).

You can find your Zotero profile directory in the following location:

- Mac: Windows 11/10/8/7/Vista; 0 Note: The /Users/<username>/Library folder is hidden by default. To access it, click on your desktop, hold down the Option key, and click the Finder’s Go menu, and then select Library from the menu.: 0 Note: If AppData is hidden on your system, click the search bar (or Start before Windows 10), type 1, and press Enter, which should take you to the AppData\Roaming directory. You can then open the rest of the path.
- Mac: Windows XP/2000; 0 Note: The /Users/<username>/Library folder is hidden by default. To access it, click on your desktop, hold down the Option key, and click the Finder’s Go menu, and then select Library from the menu.: 0
- Mac: Linux; 0 Note: The /Users/<username>/Library folder is hidden by default. To access it, click on your desktop, hold down the Option key, and click the Finder’s Go menu, and then select Library from the menu.: 0

### Reporting Zotero Problems {#reporting-zotero-problems}

Is something in Zotero not working correctly for you? Here’s the information you’ll need to provide in the Zotero Forums to allow Zotero developers and others to help you most effectively.

**Please provide both \#1 and \#2 listed below.**

#### 1. Provide a Report ID#
#### Zotero#

If you’re experiencing a problem, before restarting Zotero, open the Help menu in Zotero and select “Report Errors…”.

In the window that pops up, submit the error report, and then copy the numeric Report ID (not the contents of the report) and paste it into your Zotero Forums thread. Error reports aren’t reviewed unless referred to in the forums, since the reports generally aren’t helpful without context and/or follow-up.

If you’re not able to provide a Report ID, be sure to include your operating system and Zotero version in your forum post.

**Alternative:** If you’re reporting a reproducible problem with a particular operation, a Debug ID, which logs activity over a specific period of time, may be more useful than a Report ID. It’s always fine to provide a Debug ID instead of a Report ID if you think it might be helpful, but you can also wait for a Zotero developer to request one.

#### Zotero Connector#

If you’re experiencing a problem with the Zotero Connector, before restarting the browser, open the Zotero Connector preferences:

-   **Chrome:** Right-click on the “Save to Zotero” button and click Preferences/Options, or type chrome://extensions into the address bar, press Enter, click “Details” under Zotero Connector, and then click “Extension options”.
-   **Safari:** Right-click anywhere on a webpage and select “Zotero Preferences…”
-   **Firefox:** Type into the address bar, press return, and click the “Preferences” button next to the entry for “Zotero Connector”. Click on the “Advanced” Tab.

In the Report Errors section, click “Submit Report”. A dialog box should pop up containing a Report ID. Post the Report ID to the Zotero Forums.

If you are unable to provide a Report ID, be sure to include your Zotero version, Zotero Connector version, browser, and operating system in your forum post.

#### Zotero for iOS#

Generating a Report ID manually isn’t possible on iOS, but you can provide a Debug ID for a specific action. The app may also automatically provide a Report ID if a crash occurs.

#### 2. Provide Steps to Reproduce#

In addition to your Report ID, you’ll need to explain exactly what’s happening and, ideally, how to consistently reproduce it: if you can, there’s a very good chance we can fix it quickly or tell you how to fix it. If you’re not able to reproduce the problem, explain what happened and what you were doing when it occurred.

Post a message to the forums with the following info:

1.  The exact steps you took to reproduce the problem, including specific URLs or files you accessed, any text you entered, buttons or other interface elements that you clicked on, etc.
2.  What happened, including **exact error messages** or other relevant text that you see on the screen
3.  What you expected to happen (unless you’re reporting a clear error message)

Note that, after a problem occurs, other things in Zotero may temporarily break, so it’s important to restart Zotero and try to repeat what you did before it occurred.

#### Bad:#

> I can’t add unfiled items to a collection after selecting a tag. Does anyone know how to fix this?

#### Good:#

> Report ID: 1892199645
>
> When I drag an item from the Unfiled Items collection into another collection while a tag is selected in the bottom left, Zotero tells me it has to be restarted.
>
> Steps to reproduce:
>
> 1\. Start Zotero.
>
> 2\. In My Library, click the New Item menu and select Book.
>
> 3\. With the new item selected, add the tag “Foo” from the Tags tab in the right-hand pane.
>
> 4\. Click on “Unfiled Items” in the left pane.
>
> 5\. Click on “Foo” in the tag selector.
>
> 6\. Drag the item from the middle pane to another collection.
>
> Zotero displays the message “An error has occurred. Please restart Zotero.” in the middle pane.

#### Bad:#

> I can’t figure out how to add a web page to Zotero.

#### Good:#

> Report ID: 19672347
>
> I can’t figure out how to add a web page to Zotero.
>
> Steps to reproduce:
>
> 1\. Click the “New Item” button in the Zotero toolbar.
>
> 2\. Open the “More” menu.
>
> I expected to see “Web Page” in this menu. I see “Artwork”, “Audio Recording”, “Bill”, etc., but nothing about adding a web page. How can I add a web page?

#### Reporting Startup Errors#

For serious problems that prevent you from using the reporting wizard — such as Zotero not opening at all — we may need more information:

1\. First, be sure you’ve tried restarting your computer. Many startup errors will go away after a computer restart.

2a. If you’re able to access the Zotero Help menu, go to “Help -> Debug Output Logging -> Restart with Logged Enabled…”.

2b. If you can’t access the Help menu, or if the problem doesn’t occur during a restart, start Zotero via the command line. The steps for that depend on your platform:

##### macOS#

1.  Open Terminal via Spotlight or from /Applications/Utilities.
2.  In the terminal window that opens, paste the following and press Return. 0

If you can’t select the logging window, you can press Cmd-\` (backtick, above Tab) to cycle between windows until it’s selected.

##### Windows#

1.  Press Windows-R or search for “run” to open the Run dialog.
2.  Click Browse and locate the Zotero application directory. This is typically “C:\Program Files\Zotero\\.
3.  Select the “zotero.exe” file and click Open.
4.  The complete path to the “zotero.exe” file will be displayed in the text box. Add ” -ZoteroDebug” to the end, after any closing quotes, with a space before the hyphen. For example: 0
5.  Click OK.

##### Linux#

-   From a terminal window, run 0 within the Zotero program directory.

3a. If you used “Restart with Logging Enabled…”, you should be able to return to the Help menu, submit the output, and copy the given Debug ID to your forums post.

3b. If you used the command line, Zotero should start with a separate debug output window. Once it has stopped logging activity, click “Submit…” in the top section, click the clipboard icon to copy the Debug ID to the clipboard, and then paste the Debug ID into your forums post.

##### Alternative: Error Console#

If the above steps don’t work, repeat the steps above, try replacing 0 with 1. Zotero should open with a separate Error Console window showing errors that have occurred. Right-click on any lines with a pink background, select Copy, and paste them into your forums post.

##### Alternative: Logging to the Terminal#

If Zotero is crashing, such that the debug window opened by 0 closes as well, you can log to a terminal window instead. Upload the debug output somewhere and provide a link in your forums thread or email the output to support@zotero.org with a link to your forum thread.

### Reports {#reports}
#### Reports#

Reports are simple HTML pages that give an overview of the item metadata, notes, and attachments of the selected items. You can print them, post them to the web, and email them.

##### Generating Reports#

To create a report, right-click (ctrl-click on macOS) an item or a selection of items in the center pane and select “Generate Report from Selected Item(s)…”. You can also right-click a collection in the left column and select “Generate Report from Collection”.

##### Sharing and Printing Reports#

Reports can be saved by selecting “File -> Save…” in the File menu, and printed by selecting File -> “Print…”.

##### Working with and Searching Reports#

To copy text from a report, highlight the text and type Ctr/Cmd-C or select “Copy” from the “Edit” menu. Searching currently does not work in the Zotero Report Viewer. However, if you save a Report to your computer (“File -> Save…”), you can open it in your browser and search there.

##### Sort Order#

By default reports sort items alphabetically by title in ascending order. Sorting within the Zotero report window is not currently possible. You can, however, customize the sort order for reports by generating them from a Collection or Saved Search.

If you right-click on a collection or Saved Search in Zotero’s left pane, then choose “Generate Report from Collection/Saved Search”, Zotero will use the current sort order of the columns in the Zotero center pane for the report. To generate a report for an entire library, first make a Saved Search with the parameters: 0 1 2, then right-click on this Saved Search.

##### Customizing Reports#

It’s not currently possible to customize which fields are included in Reports within Zotero itself, but there are third-party options for doing so.

##### Uses For Reports#
##### Reviewing Abstracts#

If you need to review a large number of papers’ titles, authors, and abstracts (e.g., if you are conducting a systematic review using Zotero), reports can provide a convenient layout for reading the abstracts and writing notes in the margins.

##### Teaching#

Reports can also be used in teaching to track and assess students during the process of collecting information and writing. Reports show when items were collected, how students associate their items with notes and tags, and how students are interpreting their research items. Reports can also be a useful tool for discussing sources with students and guiding the research, organization, and writing process.

##### Organizing Notes into Outlines#

While Zotero has not been designed to be an outlining tool, you can create outlines from notes. By default, reports list child notes together with their parent items. To include child notes in your outline and separate them from their parent items, change the “extensions.zotero.report.combineChildItems” hidden preference to “false”.

Then, to build your outline, add an outline number at the beginning of each note you want to include, e.g. 1.1, 1.2, 2.1. Select the notes in Zotero, then right-click and generate a report from them.

If you are working with a large number of notes and you do not want to manually select each one, Tags and Advanced Searches can make life easier. First, tag each note with a description, such as “chapter one” or “methods”. Then create an Advanced Search for “Item Type” “is” “Note” and “Tag” “is” “chapter one”. Save the Advanced Search, then right-click the Saved Search and choose “Generate Report from Saved Search…”. This will create a report including only the notes tagged “chapter one”.

##### Disabled Features#

Zotero 5.0 opens reports in a window without an address bar or a right-click menu. As a result, several features that were previously available in Zotero for Firefox are currently disabled.

-   Sorting (but see the workaround above)
-   Searching (but see the workaround above)
-   Copying from right-click menu (but see available methods above)

### Sort Order {#sort-order}

The Zotero 5.0 Report Viewer does not include an address bar, so it is not currently possible to change item sorting in Reports. The Search function is also currently not working (but see a workaround).

The Sort syntax that was previously available in Zotero for Firefox is described below.

#### Sort Order#

By default reports sort items alphabetically by title in ascending order. You can change the sort order by appending “?sort=” to the report’s URL, followed by the item field(s) you would like to sort by. Use a comma to separate the different fields. To use a descending sort order, add “/d” after the field name. For example, a URL of a report that is first sorted by title in ascending order, and then by date in descending order looks like:

    zotero://report/items/0_KKZSDPI2/html/report.html?sort=title,date/d

You can sort by the following fields:

- ?sort=title: ?sort=date; ?sort=firstCreator: ?sort=accessed
- ?sort=title: ?sort=dateAdded; ?sort=firstCreator: ?sort=dateModified
- ?sort=title: ?sort=publicationTitle; ?sort=firstCreator: ?sort=publisher
- ?sort=title: ?sort=itemType; ?sort=firstCreator: ?sort=series
- ?sort=title: ?sort=type; ?sort=firstCreator: ?sort=medium
- ?sort=title: ?sort=callNumber; ?sort=firstCreator: ?sort=pages
- ?sort=title: ?sort=archiveLocation; ?sort=firstCreator: ?sort=DOI
- ?sort=title: ?sort=ISBN; ?sort=firstCreator: ?sort=ISSN
- ?sort=title: ?sort=edition; ?sort=firstCreator: ?sort=url
- ?sort=title: ?sort=rights


When a report is generated from a collection rather than from items selected in the center column, Zotero by default uses the order in which the items are shown in the center column.

### Why am I getting a database version error? {#why-am-i-getting-a-database-version-error}
#### Why am I getting a database version error?#

After reinstalling or downgrading Zotero, you may see the following error message when you you try to open Zotero:

*This version of Zotero is older than the version last used with your database. Please upgrade to the latest version from zotero.org.*

This message results from trying to use a Zotero database from a later version of Zotero in an earlier version of the software. For example, if you installed Zotero 10 but then reinstalled Zotero 9, you would get this error. Most Zotero versions preserve backward database compatibility, but occasionally it’s necessary for Zotero developers to make changes to the database that prevent it from working with previous versions.

The best solution is generally to reinstall the latest version of Zotero from zotero.org. (If you were using the Zotero Beta, you may need to reinstall the latest beta build.)

If you’ve fully synced your data and files with your online account, you can keep your current Zotero version by closing Zotero, deleting the zotero.sqlite, zotero.sqlite-wal, and the accompanying .bak files from the root of your Zotero data directory completely, and then starting Zotero and syncing to restore your previous data.

### Why is Zotero telling me that some data could not be downloaded? {#why-is-zotero-telling-me-that-some-data-could-not-be-downloaded}

If you or someone in a group you’re a member of have used a newer version of Zotero on another computer, you may receive this warning when syncing Zotero:

“Some data in Zotero could not be downloaded. It may have been saved with a newer version of Zotero.”

When new features or fields are added to Zotero, older versions of Zotero may not know how to handle the data required to support those changes. To avoid problems, Zotero will refuse to download data it doesn’t understand, while continuing to sync other data in your libraries.

You can click “Check for Updates” in the warning dialog to see if there’s a newer version of Zotero available. If no newer version is available, the data was likely created in the Zotero Beta, in which case you’ll need to install the beta to sync all data or ignore the warning until a newer version of Zotero is available on the main release channel.

If you’re receiving this message and don’t believe you’ve used a newer version of Zotero, please post to the Zotero Forums with a Debug ID for a sync attempt that produces the error.

### Zotero Settings {#zotero-settings}

**Note:** This section hasn’t been updated for Zotero 7, so the screenshots will look different. Most of the settings are unchanged, however.

#### Preferences Window#

Many of Zotero’s features can be customized via the Zotero preferences window. Open the preferences by clicking “Edit → Settings” (Windows/Linux) or “Zotero → Settings” (Mac). You can also press 0-1.

The Preference window is divided into the following panes:

-   **General**: Adjust appearance, import settings, and other general features.
-   **Sync**: Set up data and file syncing.
-   **Search**: Manage PDF fulltext indexing and see relevant statistics.
-   **Export**: Set default settings for generating bibliographies and citations.
-   **Cite**: Add, remove, edit, and preview citation styles and install word processor plugins.
-   **Advanced**: Zotero data location, library lookup, and other advanced settings.

#### Hidden Preferences#

In addition to the settings shown in the Zotero settings window, Zotero has a number of hidden preferences that only can be changed by opening the Config Editor under the Advanced tab of the settings window.

### “[domain] uses an invalid security certificate.” {#domain-uses-an-invalid-security-certificate}
#### “\[domain\] uses an invalid security certificate.”#

The above or a similar message generally indicates that something on your computer or network is intercepting and possibly monitoring your connection to the internet.

(If you’re getting a certificate error for a proxy or WebDAV URL, that’s likely a different issue.)

Try the following steps to debug the problem:

1.  Restart Zotero and/or your computer and try again. This error can be caused by intermittent network issues.
2.  Make sure your system clock and time zone are set correctly.
3.  If you get the same error after restarting, Zotero’s network connection is likely getting intercepted, possibly due to a proxy server on your network or security software or malware on your computer. View the certificate information for the affected domain in your browser, which may show you what is intercepting your connections.
    1.  If your browser shows the affected domain as validated by an expected entity (e.g., “Amazon” for zotero.org), Zotero’s network connection may be configured differently from your browser’s.
    2.  If your browser shows the affected domain as verified by something else and you recognize the listed entity (e.g., the name of security software or your institution), take appropriate action:
        -   If you see security software listed, disable it, or disable its SSL/TLS/HTTPS-scanning feature.
        -   If you see your institution listed, your IT department is likely intercepting your traffic and has installed a certificate for a “custom certificate authority” in your browser to avoid security errors resulting from the interception. Depending on how your system is configured, Zotero may not trust the same custom certificate by default, in which case it will properly warning you of the intercepted connection. You can try following the certificate override instructions for Zotero, but be aware that your connection to Zotero servers is being monitored by your institution.
    3.  If you don’t recognize the listed entity:
        -   Check your system for malware.
        -   If you’re using any security software, try temporarily disabling it.
        -   If you’re in an institutional environment, ask your IT department if they have installed a “custom certificate authority” in your browser. If so, you can try following the certificate override instructions for Zotero, but be aware that your connection to Zotero servers is being monitored by your institution.
        -   If you’re using a laptop, try from a different network.
        -   In rare circumstances, reinstalling Zotero may help.

## Syncing and storage {#syncing-and-storage}
### Can I store my Zotero data directory in a cloud storage folder? {#can-i-store-my-zotero-data-directory-in-a-cloud-storage-folder}

No. Storing your Zotero data directory in a cloud storage folder (Dropbox, Google Drive, OneDrive, etc) is extremely likely to corrupt your database or break Zotero in unexpected ways, and it shouldn’t be done. The same applies to essentially any database-backed program.

Database-backed programs like Zotero rely on file locking to ensure file integrity, but cloud storage systems generally don’t honor such locks. If you wake up your computer with Zotero running, and then your cloud storage tool pulls down a change from another linked computer and updates part of the database file, Zotero will be unaware of the change and will corrupt the database when it next writes to the file. Even if a cloud storage tool did honor file locks, you would simply end up with two conflicted copies that were impossible to merge, and those conflicted copies would quickly proliferate in the Zotero data directory.

The Zotero Forums contain countless reports over many years of database corruption or unexpected errors resulting from the use of cloud storage folders. Sometimes people are able to restore from a backup or recover by using the Zotero Database Repair Tool, but other people lose some or all of their Zotero data. Don’t be one of those people.

The easiest and safest way to access Zotero data across multiple computers is to keep your data directory in the default location and use Zotero Sync. For ways to safely use external cloud storage without jeopardizing your data, see Alternative Syncing Solutions.

(Technically, there is one exception to the above: if you only ever use Zotero on one computer and never even set up another instance of Zotero to point to the same cloud storage folder, storing your data directory in cloud storage should be relatively safe. However, people regularly encounter odd problems resulting from cloud storage folders not behaving like normal filesystem folders, and we’re not able to provide any support for such problems. And, of course, if you ever accidentally point another copy of Zotero at the same cloud storage folder in the future, you’re likely to corrupt the database at that time.)

### List of WebDAV services {#list-of-webdav-services}

This page contains a list of WebDAV services which provide a free plan and which users have reported success using with Zotero. Free plans may have some limitations. All service providers also offer larger and less limited paid plans. The list is not exhaustive. Service providers not listed here may still work with Zotero.

**This list is user-generated and does not entail an endorsement of any service by Zotero. Zotero works with correctly-specified WebDAV servers and can provide only minimal support for problems with 3rd party WebDAV providers.**

- Service: [4shared](http://www.4shared.com/features.jsp); Free space: 15 GB; WebDAV URL: 0; Notes and Limitations: (unofficial) Maximum file size 2 GB. Maximum 5,000 files per folder (≈2,500 Zotero attachments).
- Service: [CloudMe](https://www.cloudme.com/en/pricing); Free space: 3 GB; WebDAV URL: 0; Notes and Limitations: Maximum file size 150 MB. Users have reported sync errors with CloudMe WebDAV
- Service: [DriveOnWeb](https://www.driveonweb.de/fuer-privatanwender/produktbeschreibung); Free space: 3 GB; WebDAV URL: 0; Notes and Limitations: Documentation is in German.
- Service: [Google Drive](https://drive.google.com); Free space: 3 GB; WebDAV URL: Requires setting up a personal [WebDAV bridge](https://github.com/mikea/gdrive-webdav) or using a third-party service to provide WebDAV access to Google Drive, such as [DAV-Pocket](https://dav-pocket.appspot.com/webdav_access_to_google_docs) (may no longer be working). WebDAV URLs vary by service. Some services will not work on all operating systems.; Notes and Limitations: Consider using linked files with Zotfile instead of WebDAV syncing with Google Drive
- Service: [HiDrive](https://www.free-hidrive.com/product/hidrive-free.html); Free space: 5 GB; WebDAV URL: 0
- Service: [iDriveSync](https://www.idrivesync.com/pricing); Free space: 10 GB; WebDAV URL: 0; Notes and Limitations: Some users have had difficulties with this service and their documentation [recommends against extensive use](https://www.idrivesync.com/webdav).
- Service: [Koofr](https://koofr.eu/blog/posts/koofr-with-zotero-via-webdav); Free space: 10 GB; WebDAV URL: 0; Notes and Limitations: See setup instructions [here](https://koofr.eu/blog/posts/koofr-with-zotero-via-webdav).
- Service: [Storegate](http://www.storegate.com/en/home-user-usd/plans/); Free space: 2 GB; WebDAV URL: 0
- Service: [Yandex Disk](https://disk.yandex.ru); Free space: 10 GB; WebDAV URL: 0; Notes and Limitations: Documentation is in Russian.


Note: The “/zotero” part at the end of the WebDAV URL is added automatically by Zotero and should not be included when entering the WebDAV URL into Zotero preferences.

### Sync {#sync}

For information on how to set up and use Zotero’s syncing features, see Syncing. This page describes the settings in the Sync pane of the Zotero preferences.

The Sync pane has two tabs: “Settings” and “Reset”

#### Settings#

To set up Zotero syncing, you first need to set up data syncing (for item metadata, notes, and the full-text content) using your zotero.org username and password. After you link your Zotero account with the Zotero client program, you will see settings for managing data syncing and file syncing.

##### Data Syncing#

-   **Unlink Account…:** Disconnect the client from your Zotero account. This will prevent syncing. You will be given an option to remove the local Zotero data or keep it. If you later link this Zotero client with another username, the local library will be replaced with the new username’s library from zotero.org.
-   **Choose Libraries…:** This option lets you choose which of the Group libraries you are a member of to sync automatically with the Zotero servers. If you uncheck a library that is present in your local Zotero client, you can still manually sync the library by right-clicking on it in the left Zotero pane and choosing “Sync Library”. It is not currently possible to remove an unsynced library from the local Zotero client.
-   **Sync automatically:** When check, Zotero will start a sync every time you make a change to your library. You can manually start a sync by clicking the sync button (circular green arrow) in the upper-right corner of the Zotero window.
-   **Sync full-text content:** When checked, Zotero will sync the extracted text contents of your PDFs and other files, allowing you to perform searches across devices regardless of whether files have been downloaded to a particular device. This also allows for full-text searches in the web library.

See Data Syncing for more information.

##### File Syncing#

-   **Sync attachment files in My Library using:** Enable/disable file syncing for your personal Zotero library.
    -   **Zotero:**
        -   Sync file attachments using Zotero File Storage.
    -   **WebDAV:**
        -   Sync file attachments using WebDAV storage.
        -   Enter the URL for your WebDAV server (note that 0 is added to the end of the URL automatically), your username, and your password.
        -   Click “Verify Server” to check whether Zotero can connect with the server for file syncing.


-   **Sync attachment files in group libraries using Zotero storage:**
    -   Enable/disable file syncing for your group libraries.
    -   Only Zotero File Storage is supported for group libraries.
-   **Download files:**
    -   **At sync time:** Download all attachment files not already in your local Zotero file storage on your computer each time Zotero syncs.
    -   **As needed:** Only download attachment files when the user attempts to open the file. Useful for reducing the amount of hard disk space Zotero uses for attachments.

See File Syncing for more information.

#### Reset#

The options in this tab allow you to reset Zotero’s file sync history with zotero.org. These options are not intended for regular troubleshooting and should not be used unless directed to on the Zotero forums. For more information, see Sync Reset.

### Syncing {#syncing}

While Zotero stores all data locally on your computer by default, Zotero’s sync functionality allows you to access your Zotero library on any computer with internet access. Zotero syncing has two parts: data syncing and file syncing.

#### Data Syncing#

Data syncing merges library items, notes, links, tags, etc. — everything except attachment files — between your local computer and the Zotero servers, allowing you to work with your data from any computer with Zotero installed and to view your library online on zotero.org. Data syncing is free and unlimited, and it can be used without file syncing.

The first step to syncing your Zotero library is to create a Zotero account (which is also used for the Zotero Forums). Then, open the Sync pane of the Zotero preferences and enter your login information in the Data Syncing section.

By default, Zotero will sync your local data with the Zotero servers whenever changes are made. To disable automatic syncing, uncheck the “Sync automatically” checkbox in this section. You can sync manually at any time by clicking the “Sync with zotero.org” button on the right-hand side of the Zotero toolbar.

When Zotero syncs, it automatically applies changes in both directions — any changes you make in one place will be applied to all other synced computers. If an item has changed in multiple places in conflicting ways between syncs, you’ll receive a conflict resolution dialog asking which version you’d like to keep. If you find yourself using a new computer, you can simply set up syncing and Zotero will automatically download all data from your online library.

#### File Syncing#

Data syncing syncs library items, but doesn’t sync attached files (PDFs, audio and video files, images, etc.). To sync these files, you can set up file syncing to accompany data syncing, using either Zotero Storage or WebDAV.

##### Zotero Storage#

Zotero Storage is the recommended file sync option. It has several advantages over WebDAV syncing, including syncing of files in group libraries, web-based access to PDFs and other attachments, easier setup, guaranteed compatibility, and improved upload performance for certain files. Each Zotero user is given 300 MB of free Zotero Storage for attached files, with larger storage plans available for purchase.

See the Zotero Storage documentation for more information.

##### WebDAV#

WebDAV is a standard protocol for transferring files over the web, and it can be used to sync files in your personal library. (Group libraries cannot use WebDAV.) Your employer or research institution may be able to provide WebDAV storage. Otherwise, there are many third-party options, both free and paid (see WebDAV providers known to work with Zotero).

Once you have your WebDAV account info, enter the URL provided by the service, your username, and your password in the Sync preferences tab. Be sure to select ‘http’ or ‘https’ as appropriate — if you’re not sure, try ‘https’ first. After entering the information, click “Verify Server”. If Zotero successfully verifies the WebDAV account, you’re all set to use file syncing via WebDAV.

Zotero file sync should work with any correctly functioning WebDAV server. Zotero developers cannot provide support for third-party WebDAV servers.

#### Syncing In Practice#

If Zotero is set to sync automatically, changes will be synced within a few seconds of being made. Otherwise, you can start a manual sync by clicking the sync button on the right-hand side of the Zotero toolbar.

If you enter the same login information into the Sync preferences on multiple computers, Zotero will sync everything transparently. Your attention should only be needed if the same item or file is edited in conflicting ways in two different places before Zotero has a chance to sync them. If that happens, you’ll be presented with a conflict resolution window, where you can decide which changes to accept.

If you sync from only one computer, you can still view your online library at zotero.org from any computer. Should something happen to your computer or should you want to start using Zotero on another computer, simply set up your account info on the new computer. Zotero will pull down your entire library from the server.

#### Alternative Syncing Solutions#

If, for whatever reason, you are unable to use Zotero’s syncing features, there are some alternative ways to sync your data, though there are significant risk and limitations depending on the approach that you choose.

**Storing the Zotero data directory directly in a cloud storage folder is extremely likely to corrupt your Zotero database and should not be done.**

If you want to avoid syncing any data to Zotero servers:

-   You can close Zotero, manually copy your entire Zotero data directory to a synced folder on one computer, and then restore it — again with Zotero closed — on another computer, as if you were performing a backup and restore of your data.

If you want to use Zotero data syncing but use an external service to sync just your Zotero attachment files:

-   You can use linked files, rather than stored copies of files, with only your attachment files in the externally synced folder.

### Why am I getting “The attached file could not be found” when I try to open a file in Zotero? {#why-am-i-getting-the-attached-file-could-not-be-found-when-i-try-to-open-a-file-}
#### Why am I getting “The attached file could not be found” when I try to open a file in Zotero?#
##### Short Answer#

Sync the device **where you added the file**. Unless you deleted the file outside of Zotero, you will still be able to open it on that device. After syncing, check for a sync error in the Zotero toolbar.

You may have just reached your Zotero Storage quota, preventing additional files from being synced. You’ll need to add a storage subscription or delete some files. If you added a subscription recently, syncing the device where you added the file will allow the file to be uploaded and make it available to other devices.

If the file can’t be opened in the online library (or isn’t on your WebDAV server, if you’re using WebDAV), **any device where you can’t open the file is irrelevant**.

##### Long Answer#

If you’re unable to open a file in Zotero, it was almost certainly never uploaded to Zotero servers. File syncing may not be set up or working properly on one of your devices, or you may have reached your online file storage quota. You may see one of the following errors:

-   Desktop: *“The attached file could not be found at the following path. It may have been moved or deleted outside of Zotero, or, if the file was added on another computer, it may not yet have been synced to or from zotero.org.”*
-   iOS/Android: *“The attached file could not be found. Please check that the file has synced on the device where it was added.”*

Follow these steps to diagnose and fix the problem:

1.  Make sure the attachment file can be opened from Zotero on at least one of your devices. We’ll refer to the device that has the file as Device A and the device that doesn’t as Device B.
2.  If Device A is a computer and this is a personal library, make sure the attachment is stored within your Zotero data directory, not linked to a location elsewhere on your disk. Zotero syncs stored files; linked files need to be synced outside of Zotero. Linked files are indicated by a chain in the attachment icon. You can use Tools → Manage Attachments → Convert Linked Files to Stored Files to make the files syncable. (All files in a group are stored files.)
3.  Make sure the device is syncing properly:
    -   If Device A is a computer, check that file syncing is enabled in the Sync pane of the settings, and check whether you’re getting a sync error, indicated by an error icon to the left of the green sync button in the Zotero toolbar. If you’re getting a sync error, click the sync error icon and follow the instructions. If you need help, please post to the Zotero Forums with a Report ID and the message you’re seeing in the sync error popup.
    -   If Device A is an iOS/Android device, pull down on the items list and check whether you’re getting a sync error at the bottom of the screen. Tap on the sync error for more information. If you need help, please post to the Zotero Forums with a Debug ID for pulling down on the items list.
4.  If you’re not getting a sync error on Device A, verify that the file has been uploaded to the server. The steps vary depending on whether you’re using Zotero Storage or WebDAV:
    -   **Zotero Storage:** Check to see if the attachment file can be opened from your library on zotero.org. Attachment files will be directly viewable if they have been uploaded — the presence of an attachment item isn’t an indication that the file itself has been uploaded. If you’re using a group, the group must have File Editing enabled in its settings on zotero.org, which isn’t possible for Public, Open Membership groups.
    -   **WebDAV:** If Device A is a computer, right-click on the item and choose Show File to reveal the file in the OS file manager (Finder on macOS, Explorer on Windows, etc.). Look at the name of the file’s parent directory, which should be a string of characters such as ‘F81VWFP2’. Load your WebDAV URL in your browser or another WebDAV client and look for corresponding .prop and .zip files (e.g., F81VWFP2.prop and F81VWFP2.zip) on the server.
5.  If the file hasn’t been uploaded:
    -   If Device A is a computer, go to the Sync → Reset pane of the Zotero settings on Computer A and select Reset File Sync History for the library in question. (**Do not** use any other options in the Reset pane.) Resetting file sync history shouldn’t be necessary under normal usage, but it will cause Zotero to check every attachment to make sure that it has been uploaded to the server. If the file becomes available online, continue to the next step. If not, generate a Debug ID for the first sync after performing Reset File Sync History, followed by a successful opening of the file locally, and post it to the Zotero Forums. For Zotero Storage, include the URL from the web library where you’re not able to access the file.
    -   If Device A is an iOS/Android device, generate a Debug ID for adding a file and the sync that immediately follows. If the file isn’t uploaded, post the Debug ID to the Zotero Forums. For Zotero Storage, include the URL from the web library where you’re not able to access the file.
6.  If the file has been uploaded, check to see if it is available on Device B. If not, there may be a problem on that device. Check the file sync settings on that device to confirm that they match the settings on Device A. For example, if one device is set to use Zotero Storage and another is set to use WebDAV, or if they’re configured with different Zotero accounts (for personal library files) or different WebDAV URLs, files won’t transfer between the two.
7.  Sync Device B and check if there’s a sync error. If so, you’ll need to resolve that before file syncing will work.
8.  If you’ve performed all these steps, you’ve confirmed that the file is available on the server, and you still can’t access the file on Device B:
    -   If Device B is a computer, go to the Sync → Reset pane of the Zotero settings, choose Reset File Sync History, generate a Debug ID for the next sync attempt and an attempt to open the file, and post the Debug ID to the Zotero Forums. **Do not** use any other options in the Reset pane.
    -   If Device B is an iOS/Android device, generate a Debug ID for an attempt to open the file and post it to the Zotero Forums.

### Why aren’t changes I make syncing between multiple devices and/or zotero.org? {#why-arent-changes-i-make-syncing-between-multiple-devices-andor-zoteroorg}
#### Why aren’t changes I make syncing between multiple devices and/or zotero.org?#

*This page covers problems with data syncing. For problems with file syncing, see Files Not Syncing.*

##### 1) Looking in the wrong place#

In Zotero libraries, all existing items are shown in either the library root or the trash. Make sure you’re not looking in a collection that contains only some of your items. If you don’t see the left pane in the Zotero desktop app, click the bar at the left edge of the window to show it, or go to View → Layout and make sure Collections Pane is checked. If you don’t see your collections listed in the left-hand pane, make sure the library is expanded by clicking the arrow or plus sign next to the library. If you see a different number of items in a collection, make sure you’re not just using a different setting for View → Show Items from Subcollections.

##### 2) The right device hasn’t fully synced#

If you’re sure you’re looking in the right place, check the web library on this site to see whether the data in question has synced. If you don’t see the data in the web library, the problem is on the device where the data originated, and any other device is irrelevant. If the data appears in the web library, the problem is solely on the device where you don’t see the data. If the problematic device is a computer, check the sync icon in the Zotero app. If it’s still spinning, the sync process hasn’t yet completed. Hover over the sync icon to see the current status.

##### 3) You’ve received a sync error#

If you see a red error icon to the left of the sync icon (desktop app) or an error at the bottom of the screen (iOS/Android), an error has occurred during the sync. Click the icon or tap the message to view the error. It may give you enough information to fix the problem yourself, or you can post to the Zotero Forums for further assistance. Be sure to include the message you’re receiving and a Report ID in your forum post.

##### 4) The library isn’t set to sync#

In the Sync pane pane of the settings in the Zotero desktop app, click “Choose Libraries…” and make sure that the library you’re trying to sync has a checkmark next to it. If not, click in the Sync column to enable syncing for that library. All libraries are set to sync by default.

##### 5) You’re syncing with the wrong account.#

Check the Sync pane of the app settings to make sure you’re syncing with the same account you’re using to log in to zotero.org and on each of your devices. If the problem is with a group, make sure the account is a member of that group on zotero.org.

##### 6) Further troubleshooting#

If you’re still having trouble, post to the Zotero Forums (in your existing thread, if you have one) with a Debug ID for the first sync after making a change that doesn’t sync (e.g., making a change to an item in the online library that doesn’t appear in Zotero, or vice versa), along with a description of what’s not transferring, including the zotero.org URL after selecting the item in the online library if it appears there. Start debug output logging before making the change and keep it going until the sync has ended.

### Why do I keep getting file sync errors while syncing? {#why-do-i-keep-getting-file-sync-errors-while-syncing}
#### Why do I keep getting file sync errors while syncing?#

When attempting to sync many files to or from the Zotero servers — for example, when syncing a library on a computer for the first time — it’s not uncommon to get intermittent file sync errors. A file sync can involve thousands of network requests, and network glitches or server load can result in a small percentage of those failing. While Zotero retries many such failures automatically, Zotero may occasionally stop trying to sync a file and ask for user input.

For intermittent file sync errors, the solution is simply to press the Sync button again or, if you’re using auto-sync, to simply use Zotero normally and let it auto-sync in the background. You can hover over the sync icon to view file sync progress and see how many files are syncing each time.

The one exception is if you’re getting a file sync error at the very beginning of every sync attempt and don’t see additional progress being made each time when you hover over the sync icon. In that case, security software or a proxy server on your network may be interfering with Zotero. If you’re running security software, try temporarily disabling it. Check your system proxy settings to make sure they’re either disabled (if you’re not using a proxy) or correct (if you are). Then restart Zotero and try again.

If you’re getting immediate file sync errors and haven’t been able to fix them using the above steps, generate a Debug ID for a sync attempt that produces the error and post it to the Zotero forums.

### Why does Zotero keep asking me to reconcile the same conflicts whenever I sync? {#why-does-zotero-keep-asking-me-to-reconcile-the-same-conflicts-whenever-i-sync}
#### Why does Zotero keep asking me to reconcile the same conflicts whenever I sync?#

First, be sure you’re not cancelling the conflict resolution process: you need to click Next/Finish in the bottom-right corner of the conflict resolution window after making each change. If you don’t see a Next/Finish button, you may need to enlarge the window. (This should only be an issue on computers with low-resolution screens.)

If you’re correctly accepting changes but are still seeing the same conflicts over and over, you’re likely receiving a sync error, which should be indicated by an error icon to the left of the sync icon in the Zotero toolbar. If a sync of an item fails due to an error, any selections you’ve made in its conflict resolution window won’t be saved, and Zotero will display the same conflict again the next time you sync. To address the root problem, you’ll need to examine the error message. If you need further help, post a Debug ID for the sync attempt through receiving the error to the Zotero Forums.

### Why is Zotero still saying that my storage is full after I upgraded my storage plan or deleted files? {#why-is-zotero-still-saying-that-my-storage-is-full-after-i-upgraded-my-storage-p}
#### Why is Zotero still saying that my storage is full after I upgraded my storage plan or deleted files?#

If you’re trying to free up Zotero Storage space, first make sure you’ve emptied the trash in Zotero. Files still count against your storage quota until they’ve been permanently removed from the trash.

Beyond that, keep in mind that your storage quota, displayed on your storage settings page, applies solely to files you’ve uploaded to Zotero servers, and doesn’t affect your local usage of Zotero. You can always store as much as you wish locally — you’ll never be prevented from saving to the Zotero app due to being at your online storage quota.

This means that, if you haven’t been syncing files, or if you’ve been getting a warning for a while that you’re at your quota, you might have far more files stored in your local Zotero library than your online storage quota allows. For example, you could have 5 GB (5000 MB) of files stored locally, and your storage settings page would still show you as using the free 300 MB quota. If you then upgraded to a 2 GB storage plan, Zotero would upload 1.7 GB of additional files, hit the 2 GB quota, and show you the same warning immediately.

Similarly, if, rather than adding a storage plan, you deleted 3 GB of files from your local library, emptied the trash, and synced, your storage settings online would still correctly show you as using 300 MB, because any files you deleted from the 300 MB previously uploaded would immediately be offset by additional local files getting uploaded and filling the remaining space. You would need to delete enough so that your total local usage was less than your online quota before you would see your usage drop below your quota, which in this case would mean deleting 4.7 GB of local files.

If you’d prefer to keep all your files, you can upgrade your storage plan from the storage settings page. This would allow you to sync all your files and make them accessible from other devices and the web library, as well as to restore your files from the online library if something happened to your computer.

### Zotero Storage FAQ {#zotero-storage-faq}
#### Zotero Storage FAQ#
##### What is Zotero Storage?#

Zotero Storage provides online storage space for your Zotero files, allowing you to synchronize PDFs, images, web snapshots, and other files among all your computers, share your Zotero attachments in group libraries, and access files via your online library on zotero.org.

You can always save an unlimited number of files to your local Zotero library, with or without Zotero Storage.

##### How do I sign up for Zotero Storage?#

Simply visit your Zotero account profile and select a storage plan.

##### How can my entire lab or university or company sign up for Zotero Storage?#

Zotero Lab and Zotero Institution provide members of your organization with unlimited personal and group cloud storage.

##### How do I pay for Zotero Storage?#

Various payment methods are supported, including all major credit and debit cards, and multiple wallet and bank transfer options depending on where you’re located. Zotero Lab and Zotero Institution plans can also be paid for via bank transfer.

##### What if I don’t have a credit/debit card or my card doesn’t work?#

In addition to credit and debit cards, there are several other payment methods common in specific regions that can be used. If purchasing from the EU, you can choose to make your payment in Euro in the payment dialog which will enable several EUR specific payment methods.

If you can’t use any of the supported payment methods or you encounter problems after trying to pay via your storage settings page, please send an email to storage@zotero.org and describe the issue you’re having.

For individual storage purchases we’re unlikely to be able to make immediate changes, but we can consider supporting additional payment methods in the future.

##### Where can I find a receipt or invoice for my payment?#

You will automatically be emailed a payment notification from our payment processor (Stripe) when a charge has been made.

Detailed receipts/paid invoices for all recent payments for an individual storage subscription are available from your storage settings page.

As of August 2024, paid invoices for Lab subscription payments are available on the Lab’s management page.

##### Will files stored in my groups be freely available to other members of those groups?#

Yes, any groups that you own will draw from your storage subscription. Any members of those groups may freely access any stored files.

##### How do I allocate storage between my personal library and my groups?#

Your personal library and any groups you own automatically draw their storage from your subscription. Groups not owned by you draw their storage from the group owner’s subscription.

##### How can I change my current storage plan?#

You can switch the storage plan you’re currently subscribed to at any time from your storage settings.

When you switch to a new plan, you aren’t charged for the new plan immediately — your expiration date is simply adjusted based on the prorated balance of your previous plan. The first charge at your new subscription level will be made at your new expiration date.

*Example: It’s February 2022, and you have 6 months left on a $20 2 GB subscription, with an expiration date in August 2022. (6 months ÷ 12 months) x $20 = $10 remaining value. If you upgrade to the $60 6 GB plan, we apply that $10 to your new subscription level to set a new expiration date. ($10 ÷ $60) \* 12 months = 2 months, so your new expiration date will be two months from now, in April. You still only paid the original $20 last August, and you’ll be charged $60 for the first time in April.*

(The one exception is if you have less than 2 weeks remaining in your subscription. In that case, the unused balance will still be applied as time to your new subscription, but we’ll charge you for the new plan right away.)

##### Will my subscription renew automatically?#

When you purchase a storage subscription, you may choose to renew or not to renew automatically each year. You can cancel automatic renewal at any time from your storage settings. You’ll also receive a reminder email when your subscription is expiring or about to renew automatically.

##### How do I cancel my subscription?#

If you opted to renew your subscription automatically, you can cancel your next renewal by visiting your storage settings. As stated in the Terms of Service, paid accounts are non-refundable. We implement this policy because costs incurred on our end fluctuate dramatically according to individual usage patterns.

##### Why are there sales or VAT taxes applied?#

Zotero storage subscriptions are subject to consumption-based taxes in many jurisdictions. Depending on your location, this may be Value Added Tax (VAT), Goods & Services Tax (GST), or Consumption Tax (CT).

When you purchase a Zotero subscription, you are charged the relevant tax based on your billing address.

The tax will be displayed as an independent line item on the checkout screen when you purchase, upgrade, or renew your plan.

##### How do I apply my business tax exemption to my purchase?#

Non-US tax-registered businesses may be exempt from paying consumption-based taxes. As a tax-registered business, please check the “Add tax ID” box to input a valid VAT or GST id and avoid being charged tax on your purchases. You will be charged tax if you have not entered tax details into our system.

##### My organization is exempt from sales tax in our state. How do I apply that to my subscription?#

For Lab or Institution subscriptions, you can email us at storage@zotero.org with a copy of your sales tax exemption certificate or other relevant documentation. Note that this must be done before a purchase is made.

Individual Zotero storage subscriptions are intended for sales to individuals. We’re not able to manage sales tax exemptions for individual storage subscriptions.

### Zotero Sync Reset Options {#zotero-sync-reset-options}

This page documents the special sync operations available from the Sync → Reset pane of the Zotero preferences.

**Please note:** These operations are for use only in rare, specific situations and are not necessary during normal usage or for general troubleshooting. In many cases, resetting will cause additional problems. If you’re not sure what these options do, please ask for help on the Zotero Forums before using them.

Before using any options on this page, make sure to first back up your Zotero library.

#### Replace Online Library#

“Replace Online Library” allows you to overwrite an online Zotero library with data from your local Zotero database. This can be useful if you’ve made unwanted changes to a Zotero library locally and those changes have already synced to the online library, or if unwanted changes were made on another computer and uploaded to the online library but those changes haven’t yet been synced to your current computer.

Note that “Replace Online Library” is only necessary when you want to undo changes already applied to the online library. It isn’t necessary if you’ve simply made changes locally that you want to sync. For example, if the online library is empty and you add many items locally, those items will automatically be uploaded — the local items won’t be deleted simply because they don’t exist in the online library. Similarly, if you delete many items locally, those deletions will automatically be synced to the online library without your needing to take any special action.

##### If both the online library and your local database contain unwanted changes#

1.  Temporarily disable auto-sync in the Sync pane of the Zotero preferences.
2.  Restore your local data from either an external backup you made or one of the automatic backups in the Zotero data directory.
3.  Use “Replace Online Library” to upload the local version of your library.

See Restoring Your Zotero Library from a Backup for specific instructions for your situation.

##### If the online library contains unwanted changes that haven’t yet synced to your current computer#

1.  If Zotero isn’t yet open and you want to prevent it from syncing, temporarily disable your computer’s network connection (e.g., by disabling wifi), open Zotero, and then ensure that auto-sync is disabled in the Sync pane of the Zotero preferences.
2.  Make a backup of your Zotero data directory.
3.  Ensure that other computers are fully in sync with the online library. (The specific data being synced doesn’t matter, as you’ll be overwriting online library with the local version, but for the restore to apply to other computers without potential conflicts they need to already be in sync. It may be a good idea to make a backup of the Zotero data directory on any other computers before performing the restore.)
4.  Use “Replace Online Library” to upload the local version of your library. Be sure to choose the correct library from the drop-down. If you need to overwrite more than one online library, perform the restore separately for each library.

If the restore was successful, you can re-enable auto-sync. Keep a backup of your Zotero data directory until all applicable computers have had a chance to sync the restored version.

#### Reset File Syncing History#

If changes made to attachment files are not being synced (e.g., edits, annotations, deleting an attachment, adding a new attachment), this option will reset the file syncing history between your local Zotero database and your storage service (either the Zotero servers or your WebDAV provider). This will cause Zotero to compare all attachment files on your local computer with the ones on your storage service, making the most recent changes to files.

Resetting file sync history should not be necessary, so if you find that files aren’t syncing correctly, see Files Not Syncing for help troubleshooting and reporting the issue.

### “Error connecting to server. Check your Internet connection.” {#error-connecting-to-server-check-your-internet-connection}
#### “Error connecting to server. Check your Internet connection.”#

When Zotero can’t access the network, it’s usually caused by proxy settings or security software on your computer.

By default, Zotero uses any manually entered proxy servers or a proxy auto-config (PAC) URL in your system proxy settings. It does not automatically use Web Proxy Auto-Discovery (WPAD), or “Auto Proxy Discovery” on macOS, even when enabled in the system settings.

If you don’t use a proxy to connect to the internet, you should disable all proxies in your system proxy settings. If you do need to connect via a proxy, you should verify that the system settings are correct. Note that other software on your computer may not be using the system proxy settings, which is why other programs may still be able to connect to the internet when Zotero cannot.

If you need to configure the Zotero proxy settings differently from the system settings, you can access the Config Editor from the Advanced pane of the Zotero preferences, apply the [same settings that you would in Firefox](http://kb.mozillazine.org/Network.proxy.type) (on which Zotero is based), and restart Zotero, but the default setting (network.proxy.type = 5, to use the system proxy settings) is recommended.

If using a PAC file, either automatically or with network.proxy.type = 2, and your proxy requires HTTP authentication, ensure that most or all of the hosts in 0 are being handled by your PAC file. Zotero will test a random subset on each startup to trigger a proxy authentication prompt if necessary.

To use WPAD, you must set network.proxy.type to 4.

If changing proxy settings doesn’t help, try temporarily disabling any security/firewall software on your system.

Some connection errors can also be caused by certificate issues on your network.

See also Zotero and Firewalls.

## Zotero on iPhone and iPad {#zotero-on-iphone-and-ipad}
### Zotero for iOS {#zotero-for-ios}

[Zotero for iOS](https://apps.apple.com/us/app/zotero/id1513554812) is the best way to work with your Zotero library on an iPad or iPhone.

Zotero for iOS lets you work with your Zotero data no matter where you are:

-   Sync your personal and group libraries
-   View and edit item details
-   Organize items into collections
-   Generate citations and bibliographies in any of 10,000+ citation styles (APA, Chicago, MLA, etc.)
-   Take notes on your research
-   Read PDFs and add highlight, note, image, and ink annotations
-   Easily save items and PDFs from the web to Zotero via the Share button in Safari or other apps (other browsers, Twitter, etc.)
-   Quickly add physical books and articles to your Zotero library by scanning book barcodes or article DOIs with your iPhone or iPad camera
-   Avoid straining your eyes with full Dark Mode support, including while reading PDFs

Back on your computer, you can add PDF annotations you’ve made on your iPad or iPhone to Zotero notes and insert those notes into your word processor document with active Zotero citations or export them to Markdown.

Please post to the Zotero Forums with any bug reports or feature requests. Be sure to mention “iOS”, “iPad”, or “iPhone” in your thread title.

[Get Zotero for iOS on the App Store](https://apps.apple.com/us/app/zotero/id1513554812)

### Zotero for Mobile {#zotero-for-mobile}

Official apps are available for iOS and Android:

-   [Zotero for iOS](https://apps.apple.com/us/app/zotero/id1513554812)
-   [Zotero for Android](https://play.google.com/store/apps/details?id=org.zotero.android)

The mobile version of the Zotero web library also allows you to access and edit your Zotero library on your tablet or mobile phone.

You can also save items to your Zotero account using zotero.org/save.
