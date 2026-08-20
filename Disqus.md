---
title: "Disqus Help Guide"
subtitle: "Disqus's official help centre, gathered into one screen-reader-friendly document"
date: 2026-07-31
lang: en
---

# Disqus Help Guide

This guide gathers the official help Disqus publishes for its commenting system, taken from Disqus's own help centre on 31 July 2026. The words are Disqus's; nothing has been summarised or rewritten.

Disqus is the comment system that sits under articles on a great many websites, so most people meet it without ever visiting disqus.com. Both sides of it are covered here: the commenter's, and the site owner's who installs and moderates it. It runs in a web browser; there is no Windows program, no app of its own for iPhone or iPad, and no Alexa skill.

The sections are **Disqus's own**: the collections its help centre files each article under, taken from the trail at the top of every page, in alphabetical order with Miscellaneous last. Two collections that held a single article each are folded into Miscellaneous rather than standing alone.

**27 articles from the Developer and API collections were left out**, in keeping with this collection's rule that programming documentation belongs elsewhere. Nothing else was excluded.

One section is larger than you might expect. Terms and Policies holds 27 articles — Disqus files its terms of service, privacy policy, advertising rules and content guidelines as help articles like any other. The how-to material lives under Installation, Commenting, Moderation and Troubleshooting.

Illustrations do not survive as pictures; where one carried a description it is kept as a short italic note. Disqus's own site navigation, its search box, and the "Did this answer your question?" strip at the foot of every article have been removed.

Headings run six deep at most. Level one is this title, level two is a collection, level three is an article, and anything below that came from inside the article itself.

## Contents {#contents}

- [Ads](#cat-ads)
    - [Ads FAQ](#ads-faq)
    - [Ads.txt FAQ](#ads-txt-faq)
    - [Ads.txt Implementation Guide](#ads-txt-implementation-guide)
    - [Getting started with Disqus Advertising](#getting-started-with-disqus-advertising)
    - [In-thread ads FAQ](#in-thread-ads-faq)
    - [Receiving payments from Disqus](#receiving-payments-from-disqus)
    - [Updating Your Advertising Settings](#updating-your-advertising-settings)
- [Analytics](#cat-analytics)
    - [Capture Disqus commenting activity via callbacks](#capture-disqus-commenting-activity-via-callbacks)
    - [Disqus Ads Analytics](#disqus-ads-analytics)
    - [Disqus Basic Analytics](#disqus-basic-analytics)
    - [Understanding Earnings and Engagement](#understanding-earnings-and-engagement)
    - [Why am I seeing traffic from disqus.com/embed/comments in my analytics?](#why-am-i-seeing-traffic-from-disqus-com-embed-comments-in-my)
- [Commenting](#cat-commenting)
    - [Accessibility on Disqus](#accessibility-on-disqus)
    - [Adding Images and Videos](#adding-images-and-videos)
    - [Browser plugin/extension conflicts](#browser-plugin-extension-conflicts)
    - [Comment Text Formatting](#comment-text-formatting)
    - [Commenting 101](#commenting-101)
    - [Delete account or access account data](#delete-account-or-access-account-data)
    - [Disqus Digests](#disqus-digests)
    - [Disqus Web Notifications](#disqus-web-notifications)
    - [Featured Comment FAQ](#featured-comment-faq)
    - [Flagging comments](#flagging-comments)
    - [Following other users](#following-other-users)
    - [Guest Commenting](#guest-commenting)
    - [Mentions](#mentions)
    - [Remove and Edit Your Comments](#remove-and-edit-your-comments)
    - [Subscribe/Unsubscribe from Notifications](#subscribe-unsubscribe-from-notifications)
    - [Two-Factor Authentication (2FA)](#two-factor-authentication-2fa)
    - [User Blocking](#user-blocking)
    - [Voting](#voting)
    - [Who deleted/removed my comment or blocked my account?](#who-deleted-removed-my-comment-or-blocked-my-account)
    - [Why are my comments marked as spam or removed?](#why-are-my-comments-marked-as-spam-or-removed)
    - [Why can't I comment?](#why-can-t-i-comment)
    - [Windows Phone app help](#windows-phone-app-help)
- [Community Tips](#cat-community-tips)
    - [Best Practices for Moderating Sites](#best-practices-for-moderating-sites)
    - [Build Your Community](#build-your-community)
    - [Comment Policy](#comment-policy)
    - [Connect Your Content to Your Community](#connect-your-content-to-your-community)
    - [Create and Enforce Moderation Guidelines](#create-and-enforce-moderation-guidelines)
    - [Default avatars for communities](#default-avatars-for-communities)
    - [Embedding a comment on your website or blog](#embedding-a-comment-on-your-website-or-blog)
    - [Guiding your Community from Good to Great](#guiding-your-community-from-good-to-great)
    - [How to Increase Reader Engagement and Retention](#how-to-increase-reader-engagement-and-retention)
    - [Personalize Your Site](#personalize-your-site)
    - [Recommendations](#recommendations)
    - [Sample Community Guidelines](#sample-community-guidelines)
    - [Site identity information](#site-identity-information)
    - [Site Profiles](#site-profiles)
- [Disqus Polls](#cat-disqus-polls)
    - [Disqus Polls](#disqus-polls)
    - [Disqus Polls - Pricing and Plans](#disqus-polls-pricing-and-plans)
    - [Polls](#polls)
- [Disqus Pro](#cat-disqus-pro)
    - [Badges](#badges)
    - [Disqus Appearance Customizations](#disqus-appearance-customizations)
    - [Disqus Pro Analytics](#disqus-pro-analytics)
    - [Email Subscriptions](#email-subscriptions)
    - [Shadow banning](#shadow-banning)
    - [Timeouts](#timeouts)
- [Import, Export, and Syncing](#cat-import-export-and-syncing)
    - [Can I import comments from Facebook Comments?](#can-i-import-comments-from-facebook-comments)
    - [Domain Migration Tool](#domain-migration-tool)
    - [How to download, edit, and upload a URL Map CSV](#how-to-download-edit-and-upload-a-url-map-csv)
    - [Importing & Exporting](#importing-exporting)
    - [Importing comments from JS-Kit](#importing-comments-from-js-kit)
    - [Importing comments from WordPress](#importing-comments-from-wordpress)
    - [Importing Typepad comments](#importing-typepad-comments)
    - [Migration Tools](#migration-tools)
    - [Moving from Blogger to WordPress](#moving-from-blogger-to-wordpress)
    - [Redirect Crawler](#redirect-crawler)
    - [Syncing with WordPress](#syncing-with-wordpress)
    - [URL Mapper](#url-mapper)
- [Installation](#cat-installation)
    - [Add Disqus to Static Pages in Blogger](#add-disqus-to-static-pages-in-blogger)
    - [Adding Disqus to your site](#adding-disqus-to-your-site)
    - [Can Disqus be used on WordPress.com sites?](#can-disqus-be-used-on-wordpress-com-sites)
    - [Configuring Disqus on your site](#configuring-disqus-on-your-site)
    - [How to use trusted domains](#how-to-use-trusted-domains)
    - [Manually adding a Disqus gadget to Blogger](#manually-adding-a-disqus-gadget-to-blogger)
    - [Manually install Disqus on WordPress](#manually-install-disqus-on-wordpress)
    - [Moderation Profiles](#moderation-profiles)
    - [Multi-lingual websites](#multi-lingual-websites)
    - [Publisher Quick Start Guide](#publisher-quick-start-guide)
    - [Translating Disqus](#translating-disqus)
    - [Universal Embed Code](#universal-embed-code)
    - [Updating a Blogger template to support all versions of Internet Explorer](#updating-a-blogger-template-to-support-all-versions-of-inter)
    - [What's a shortname?](#what-s-a-shortname)
    - [Will I lose comments if I uninstall Disqus?](#will-i-lose-comments-if-i-uninstall-disqus)
- [Known Issues](#cat-known-issues)
    - [Blogger Syncing and Importing](#blogger-syncing-and-importing)
    - [Follower Notification Emails - Known Issue](#follower-notification-emails-known-issue)
    - [Known Issue: Disqus not loading via Spanish ISPs: Movistar/Telefonica de Espana](#known-issue-disqus-not-loading-via-spanish-isps-movistar-tel)
    - [Security Best Practices](#security-best-practices)
    - [Twenty Thirteen/Fourteen/Fifteen (theme) Conflict in WordPress - Known Issue](#twenty-thirteen-fourteen-fifteen-theme-conflict-in-wordpress)
- [Moderation](#cat-moderation)
    - [Advanced Moderation](#advanced-moderation)
    - [Dealing with spam](#dealing-with-spam)
    - [How do I change the ownership of a Site or Organization?](#how-do-i-change-the-ownership-of-a-site-or-organization)
    - [How do I Close and Open comment threads?](#how-do-i-close-and-open-comment-threads)
    - [How to add Admins and Moderators to your Organization](#how-to-add-admins-and-moderators-to-your-organization)
    - [Moderating 101](#moderating-101)
    - [Moderation Rules](#moderation-rules)
    - [Moderation Settings](#moderation-settings)
    - [Reactions](#reactions)
    - [Site Moderators](#site-moderators)
    - [Toxic Mod Filter](#toxic-mod-filter)
    - [User Reporting](#user-reporting)
    - [User Reputation](#user-reputation)
    - [Using the "Ban User" and "Trust User" controls](#using-the-ban-user-and-trust-user-controls)
- [Other Integrations](#cat-other-integrations)
    - [Sitefinity Installation Instructions](#sitefinity-installation-instructions)
    - [Tumblr Manual Installation Instructions](#tumblr-manual-installation-instructions)
    - [Use Zapier to connect other apps](#use-zapier-to-connect-other-apps)
- [Terms and Policies](#cat-terms-and-policies)
    - [Abusive Behavior Policy](#abusive-behavior-policy)
    - [Ads-Free Subscription & Payments FAQ](#ads-free-subscription-payments-faq)
    - [Amendment to Disqus Terms of Service Applicable to U.S. Federal Government Users](#amendment-to-disqus-terms-of-service-applicable-to-u-s-feder)
    - [Basic Rules for Disqus](#basic-rules-for-disqus)
    - [Basic Rules for Disqus-powered Sites](#basic-rules-for-disqus-powered-sites)
    - [Comments Pricing and Plans](#comments-pricing-and-plans)
    - [Contacting Disqus about a deceased user](#contacting-disqus-about-a-deceased-user)
    - [Cookies and Data Recipients](#cookies-and-data-recipients)
    - [Data Processing Agreement for Publishers](#data-processing-agreement-for-publishers)
    - [Disqus - Publisher Terms of Service Agreement for Ad Management Solutions](#disqus-publisher-terms-of-service-agreement-for-ad-managemen)
    - [Disqus Informativa Sulla Riservatezza](#disqus-informativa-sulla-riservatezza)
    - [Disqus Privacy Policy](#disqus-privacy-policy)
    - [Disqus-Datenschutzrichtlinie](#disqus-datenschutzrichtlinie)
    - [General Security Tips](#general-security-tips)
    - [How to Edit Your Data Sharing Settings](#how-to-edit-your-data-sharing-settings)
    - [How to Report Abuse](#how-to-report-abuse)
    - [How to report threats of suicide or self-harm](#how-to-report-threats-of-suicide-or-self-harm)
    - [Is someone else posting using my account?](#is-someone-else-posting-using-my-account)
    - [Parody Accounts](#parody-accounts)
    - [Politique de Confidentialité de Disqus](#politique-de-confidentialite-de-disqus)
    - [Política de Privacidad de Disqus](#politica-de-privacidad-de-disqus)
    - [Política de Privacidade do Disqus](#politica-de-privacidade-do-disqus)
    - [Privacy FAQ](#privacy-faq)
    - [Spam](#spam)
    - [Targeted harassment or encouraging others to do so](#targeted-harassment-or-encouraging-others-to-do-so)
    - [Terms of Service](#terms-of-service)
    - [Trademark Policy](#trademark-policy)
- [Troubleshooting](#cat-troubleshooting)
    - [Adding Disqus to static Wordpress Pages](#adding-disqus-to-static-wordpress-pages)
    - [Blogger Troubleshooting](#blogger-troubleshooting)
    - [I'm receiving the message "We were unable to load Disqus."](#i-m-receiving-the-message-we-were-unable-to-load-disqus)
    - [Installation Troubleshooting](#installation-troubleshooting)
    - [Introducing the Discussions Editor and FAQ](#introducing-the-discussions-editor-and-faq)
    - [Troubleshooting 101](#troubleshooting-101)
    - [Troubleshooting Common Error Messages](#troubleshooting-common-error-messages)
    - [Troubleshooting Disqus in Internet Explorer 8/9/10](#troubleshooting-disqus-in-internet-explorer-8-9-10)
    - [Troubleshooting Imports](#troubleshooting-imports)
    - [Tumblr Troubleshooting](#tumblr-troubleshooting)
    - [Use Configuration Variables to Avoid Split Threads and Missing Comments](#use-configuration-variables-to-avoid-split-threads-and-missi)
    - [Why are comments being posted to Blogger instead of Disqus?](#why-are-comments-being-posted-to-blogger-instead-of-disqus)
    - [Why are comments posted to other sites showing up in my Disqus admin?](#why-are-comments-posted-to-other-sites-showing-up-in-my-disq)
    - [Why are comments visible in the Disqus admin but not on my site?](#why-are-comments-visible-in-the-disqus-admin-but-not-on-my-s)
    - [Why are the same comments showing up on multiple pages?](#why-are-the-same-comments-showing-up-on-multiple-pages)
    - [Why are the wrong URLs detected for my discussions?](#why-are-the-wrong-urls-detected-for-my-discussions)
    - [Why isn't the comment box loading?](#why-isn-t-the-comment-box-loading)
    - [WordPress Troubleshooting and FAQ](#wordpress-troubleshooting-and-faq)
- [User Profile](#cat-user-profile)
    - [Improving Your User Profile](#improving-your-user-profile)
    - [Log into Disqus with a Social Media Account](#log-into-disqus-with-a-social-media-account)
    - [Logging into Disqus](#logging-into-disqus)
    - [Making Your Activity Private](#making-your-activity-private)
    - [Registering a commenter account](#registering-a-commenter-account)
    - [Site-Specific Profiles](#site-specific-profiles)
    - [Updating your Account Settings](#updating-your-account-settings)
    - [Use of cookies](#use-of-cookies)
    - [User Profiles 101](#user-profiles-101)
    - [What is the difference between my Username and my Display Name?](#what-is-the-difference-between-my-username-and-my-display-na)
    - [Your Homepage on Disqus](#your-homepage-on-disqus)
- [What is Disqus?](#cat-what-is-disqus)
    - [Common Questions About Disqus](#common-questions-about-disqus)
    - [Disqus Glossary](#disqus-glossary)
    - [How does Disqus work?](#how-does-disqus-work)
    - [What is Disqus?](#what-is-disqus)
- [Miscellaneous](#cat-miscellaneous)
    - [Channel Help](#channel-help)
    - [Disqus Advertising Content Guidelines](#disqus-advertising-content-guidelines)

## Ads {#cat-ads}

### Ads FAQ {#ads-faq}

***\*\*Note: Ads analytics are only available to publishers eligible for Disqus Ads. More information on Disqus Ads can be found**here**.***

Curious about Disqus Ads? We've answered the most common questions below to help you get started.

After signing up for Disqus, publishers will have the opportunity to apply for Ads by running ads for a full week and then reaching out to us as described here.

#### What does 'below-the-fold' mean?

'Below-the-fold' is the area of a webpage that is only visible after a reader scrolls down the page.

#### What ad types will I see on my site?

You're in control of what ad types appear on your site. Quality **native content ads (**Sponsored Story and Sponsored Links**)** are just a click away at Disqus Admin > Revenue > Settings.

#### How do I know how much I'm earning?

An earnings report is available in the Disqus Admin > Revenue > Settings.

#### How do I know when I'll be paid?

Your upcoming payments and unpaid earnings are available in the Disqus admin panel Revenue > Analytics > Payments.

For more important payment details, please visit: Receiving payments from Disqus

#### Where are the settings for Disqus Ads?

Ads settings are in the Disqus Admin > Revenue > Settings and can only be accessed by the forum founder and moderators of the forum who have "can change settings" permissions. For more information on Disqus Ads settings, please visit: Updating Your Advertising Settings.

#### How do Disqus ads impact page load time?

Page load time isn't impacted by ads. Disqus is non-blocking and waits to load both the discussion and the ads until all other page elements have loaded.

#### More information is available here: Does Disqus impact page load time?  What if I want to remove Ads?

If you'd like to remove Disqus Ads from your integration, you may purchase and ads-free subscription from your Subscription and Billing page. More information on Disqus ads-free subscriptions may be found [here](#comments-pricing-and-plans).

### Ads.txt FAQ {#ads-txt-faq}

#### What is ads.txt?

Ads.txt stands for Authorized Digital Sellers and is a simple, flexible, and secure method for publishers and distributors to declare who is authorized to sell their inventory, improving transparency for programmatic buyers. By creating a public record of Authorized Digital Sellers, ads.txt will create greater transparency in the inventory supply chain and give publishers more control over their inventory in the market, making it harder for bad actors to profit from selling counterfeit inventory across the ecosystem. As publishers adopt ads.txt, buyers will be able to more easily identify the Authorized Digital Sellers for a participating publisher, allowing brands to have confidence they are buying authentic publisher inventory.
​
For more background see:
IAB ads.txt FAQ here: 0
How to create an ads.txt file for publishers

Adding an Ads.txt file on Wordpress

#### What problems is this intended to address?

Adoption of ads.txt can help prevent domain spoofing and the sale of unauthorized inventory by providing a simple way for demand-side platforms (DSPs) to check if inventory from a particular source has been authorized for sale by the domain owner.

**What are the Publisher IDs for my domain to put in the ad call that you send to buyers?**

If you are already registered with Disqus, you may view your Ads.txt lines here. If you are not yet registered, you may view the lines at 0.

#### What is a Publisher ID?

Disqus has a unique Publisher ID for each ad provider that we work with; this helps to identify Disqus specifically as the provider of this ad inventory. A Publisher ID is the ID associated with a Disqus’ account on an exchange or SSP platform. As a best practice this ID is transmitted as part of the ad call as the Publisher.ID along with the Publisher.Domain in the Publisher object.

#### Where is the latest ads.txt file from Disqus?

When logged in with your admin account, you can locate your latest file here. We recommend that you update this file on a monthly basis.

#### How do I implement an ads.txt file for my site?

Please see our [implementation guide](#ads-txt-implementation-guide) for further instructions.

#### Will you send me the Publisher ID that corresponds to our integrations, for all sites?

Yes, see the actual ads.txt file itself.

#### Are there any other Publisher ID’s that are connected to our account?

No, there is only one Publisher ID and it will remain the same.

#### Can I have consent to expose these IDs on our sites via ads.txt?

Yes, you can expose any information that is in Disqus’ ads.txt file via your own implementation of ads.txt.

#### Why should I support ads.txt?

By publishing an ads.txt file, you are indicating which parties are authorized to sell your inventory. As enforcement becomes widespread, this will make it harder for unauthorized parties to spoof your inventory leading to more spend on your legitimate inventory.

#### Where can I find further information and documentation?

0

0
​
​
Feel free to reach out with any other questions about implementing Ads.txt

### Ads.txt Implementation Guide {#ads-txt-implementation-guide}

The following is a step-by-step guide for implementing our latest [ads.txt](#ads-txt-faq) file:

1.  If you already have an ads.txt file implemented for your site, you can simply copy the most up-to-date Disqus ads.txt file from your ad settings page page and paste it into your existing ads.txt file.
    ​

2.  If you do not yet have an ads.txt file implemented:

    1.  Copy the Disqus ads.txt file from your ad settings page and paste it into a new document.

    2.  Save your document as a .txt file and name it ads.txt.

    3.  Upload the ads.txt file to the root directory of your domain's server. The root directory of a site is the directory or folder following the top-level domain (example.com/ads.txt). The file can be accessible via HTTP or HTTPS but must be located under a standard relative path: "ads.txt". The HTTP request header must contain "Content-Type: text/plain"

To ensure you’re maximizing your ad revenue, continue to update your file on a monthly basis.

More information on Ads.txt may be found [here](#ads-txt-faq).

### Getting started with Disqus Advertising {#getting-started-with-disqus-advertising}

Disqus’ free-to-use, Basic plan is supported by advertising. Publishers using this version of Disqus will see advertising within the Disqus embed on their sites. Publishers using other plans can choose to enable or disable Disqus advertising on their sites. Qualifying publishers can participate in the Disqus Ads program to earn revenue with Disqus. Disqus advertising is highly configurable, allowing publishers to choose from several ad positions and types.
​
*To participate in the Disqus Ads program, please refer to the* *Revenue Qualification* *page for more details.*

-   Advertising Settings

-   Getting Paid

-   Earnings and Performance Reports

-   Upcoming Payments

-   Getting Help and Giving Feedback

#### Advertising Settings

To see and configure your settings, go to your Ads Settings. Ads settings are available to organization owners, organization admins, and moderators with 'can edit settings' permissions. How to disable advertising.

#### Choose your ad types

Whether you choose to show ads above, below, or in-thread with your comments, quality native content and premium ads are just a click away in your Ads Settings page.

#### Payment information

Complete payment information is required in order to receive payments from Disqus.

#### Getting Paid

#### Requirements for getting paid

Earnings for a single shortname must exceed 100.00 USD (after fees) and payment information must be complete.

#### Payment timing

Payments over 100.00 USD are sent once a month 90 days after the revenue was earned. Payments are sent by month end.

#### Missing payments

The most common reason for missing payments is incomplete payment information.

#### Earnings and Performance Reports

#### Your site's revenue

A report of daily earnings is available to all Ad-enabled publishers in the Admin > Revenue > Analyze revenue. In this report you'll find:

-   **Ad Revenue:** Amount you earned from Disqus advertising

-   **Viewable Impressions:** How often users scroll down below the fold and see Disqus ads

-   **Viewability %:** How often users scroll below the fold and see Disqus ads compared to overall page impressions

-   **RPMv:** Revenue earned per 1,000 viewable impressions

#### Upcoming Payments

Your site's upcoming payments is available in Payments.

#### Upcoming payments

Upcoming payments are shown in green along with an estimated delivery date. Unpaid earnings are shown in yellow and the money earned during each period is paid out 90 days after that earnings period completes.

#### Getting Help and Giving Feedback

#### Contact the publisher success team

Send your Ads questions and feedback to our Publisher Support Team and we'll be happy to assist you.

-   Advertising FAQ

-   Updating Your Advertising Settings

-   Receiving payments from Disqus

### In-thread ads FAQ {#in-thread-ads-faq}

The in-thread ad position supports the Sponsored Story and IAB Display ad types.. They appear in between comments.

Sponsored Story ads will adapt to the color scheme of your site. On light colored sites, in-thread ads will have a white background, as seen in the image above. On dark colored sites, in-thread ads will have transparent backgrounds.

#### When will it display?

If you have the in-thread ad position and the sponsored story or IAB display ad type\* selected, an in-thread ad will display in all threads that have 4 or more parent level comments (not replies). The first ad will display after the 3th parent level comment. The 4 comment minimum is so that there will always be at least one comment below the in-thread ad Additional in-thread ads will display for every 6 additional parent level comments.

***\*Sponsored story and IAB Display are the only ad types currently supported by the in-thread position. At least one of these two ad types needs to be selected for in-thread to display.***

#### How can I turn them on and off?

You can turn in-thread ads on and off, at any time from your Ads Settings page.

#### Can I control which threads it displays on?

Not at this time. If the in-thread ad position is active for your site, in-thread ads will display on all pages that have at least 4 parent level comments.

#### Is it on mobile?

Yes, the in-thread ad position will display on mobile as long as all of the above criteria are met.

#### Can I run the in-thread position alone, without any other positions?

Currently, we support the in-thread only configuration exclusively for select publishers who are approved by our Publisher Success team. To learn more or discuss your site’s settings, please reach out to our Publisher Success team.

### Receiving payments from Disqus {#receiving-payments-from-disqus}

**The following conditions must be met before a payment can be issued:**

-   Must be eligible for Disqus Ads.

-   Earnings must exceed the amount of 100.00 USD (after fees). Earnings and payments are on a per forum basis – earnings from multiple forums cannot be consolidated into one payment.

-   Payment information must be complete. If you have more than one forum with Ads enabled, payment information must be completed for each forum **individually**. Only moderators with 'Can change settings' permissions can access payment information.

Payments are sent once a month **by month end**. Disqus payment terms are net 90, so payments earned through July are paid at the end of October, earnings through August at the end of November, etc. Earnings must be due and over \$100 to be paid.

#### Payment Information

You can locate and complete payment information using the following steps:

1\. Navigate to your ads settings.

2\. Click "Set up or update your payment method" under "Payment information".

3\. Complete the following forms:

-   Contact information

-   Method of Payment

-   Tax form

We'll send you a confirmation email once your payment information is complete and for each payment we make to you.

#### What Fees are charged?

Fees may vary depending on your country. Our payment service provider charges nominal fees depending on your form of payment, the fees are generally as follows:

-   PayPal - \$1 (+ normal PayPal fees according to your individual settings)

-   ACH - \$1

-   Wire Transfer - \$15

-   eCheck - \$5

-   Check - \$6

#### International Payments

Can I receive payments in EUR or USD if I’m in a country not in the European Union or the US? Yes, if your chosen payment method supports it. For example, if you are in Germany and choose "Wire Transfer" as a payment method, the option to be paid in Euros, Pounds, or Dollars is available:
​

#### Still have questions?

Contact our support team with the shortname of the forum you have questions about and any other important details.

### Updating Your Advertising Settings {#updating-your-advertising-settings}

The Basic version of Disqus is ads supported, but you can easily control which ad types appear on your site by going to your ads settings. Click Settings from the top nav, and then select Ads on the left side of the page.

You'll find a variety of settings that can be configured to fit the needs and appearance of your site.
​

If you're logged into your moderator account, visit your Ads Settings page. If you have more than one forum, you may be prompted to select which forum you wish to see settings for.

#### Enabled Positions

Here you can select the positions where ads display. By default, ads are enabled to display Above the comments. You can also choose to display ads In-thread (within the comments) and below the comments.

-   ***Above Comments***: With this position, an ad will always display above the comments, whether it's a Sponsored Story (our default ad type) or Sponsored Links.

-   ***In-thread Position***: With this position, an ad will display within the comments on every page that has 4 or more parent level comments. Sponsored Story is the only ad type that this position supports, if Sponsored Story is not selected, In-thread ads will not display.

-   ***Below Comments***: With this position, an ad will always display below the comments whether it's a Sponsored Story (our default ad type) or Sponsored Links with Thumbnails.

#### Enabled Positions Preview

In the Enabled Positions Preview, you can see a preview of where ads will appear on your site based on your enabled positions selections.

#### Appearance

Disqus ads will automatically adapt to your site's color scheme and typography. If you'd like to adjust these settings yourself, you can manually update your site's appearance settings from the General Settings page.

#### Category

Remember to select a category for your site. This helps us to serve the more relevant ads to your audience based on your site’s content.

#### Earnings Potential (for Ads eligible publishers, only)

The meter indicates your site's earnings potential based on your enabled positions and allowed ad types. Select all ad types and positions to ensure that the best ads are served to all of your readers and to maximize your earnings potential. The Above position and Sponsored Story ad type have the biggest impact, disabling these options will significantly decrease your earnings potential.

#### Payment Information

Remember to fill out your payment information in order to receive payments from Disqus. Learn more about getting paid here.

#### Disabling Advertising

Advertising is a default part of the free version of Disqus. For sites where advertising is not a good fit, we have subscription options that include the ability to remove ads.
​
If you're experiencing issues with quality, content, relevance, or any other ad-related issue, please contact our publisher success team.
​

## Analytics {#cat-analytics}

### Capture Disqus commenting activity via callbacks {#capture-disqus-commenting-activity-via-callbacks}

If you would like to track new comments and replies via your own analytics service, such as Omniture or Google Analytics, you can do so via the following callback function.

The following callback can be added to the 0 function which already exists within the universal embed code:

    var disqus_config = function () {
        this.callbacks.onNewComment = [function() { trackComment(); }];
    };

Make sure to replace 0 with the script you wish to track via your analytics service.

**NOTE:** If you're using the Wordpress plugin, you'll need to edit the plugin and add your tracker to the existing 0 function.

This callback function accepts one parameter 0 which is a JavaScript object with comment ID and text. For example you can find the unique comment ID and comment text for further analysis the following way:

    var disqus_config = function () {
        this.callbacks.onNewComment = [function(comment) {
          alert(comment.id);
          alert(comment.text);
        }];
    }

#### Example within Full Embed Code

        /**
         *  RECOMMENDED CONFIGURATION VARIABLES: EDIT AND UNCOMMENT THE SECTION BELOW TO INSERT DYNAMIC VALUES FROM YOUR PLATFORM OR CMS.
         *  LEARN WHY DEFINING THESE VARIABLES IS IMPORTANT: 0
         */

        var disqus_config = function () {
            this.page.url = PAGE_URL;  // Replace PAGE_URL with your page's canonical URL variable
            this.page.identifier = PAGE_IDENTIFIER; // Replace PAGE_IDENTIFIER with your page's unique identifier variable
            this.callbacks.onNewComment = [function(comment) {
                  alert(comment.id);
                  alert(comment.text);
            }];
        };

        (function() {  // DON'T EDIT BELOW THIS LINE
            var d = document, s = d.createElement('script');

            s.src = '//EXAMPLE.disqus.com/embed.js';

            s.setAttribute('data-timestamp', +new Date());
            (d.head || d.body).appendChild(s);
        })();

    Please enable JavaScript to view the comments powered by Disqus.

### Disqus Ads Analytics {#disqus-ads-analytics}

***\*\*Note: Ads analytics are only available to publishers eligible for Disqus Ads. More information on Disqus Ads can be found here.***
​
***For commenting analytics, please see our articles [here](#disqus-basic-analytics) and [here](#disqus-pro-analytics).***

Once you've gotten confirmation that you've been set up to earn revenue with Disqus Ads, an Ads tab will appear at the top of your moderation panel. Clicking into that tab will give options to monitor both earned revenue and payments.

#### Revenue

You'll find a line chart and table that details the following performance metrics for your forum. This data can be filtered by any date range you'd like using the provided drop down.

-   **Ad Revenue:** Amount you earned from Disqus advertising

-   **Viewable Impressions:** How often users scroll down below the fold and see Disqus ads

-   **Viewability %:** How often users scroll below the fold and see Disqus ads compared to overall page impressions

-   **RPMv:** Revenue earned per 1,000 viewable impressions

#### Payments

Your Payments page will provide additional rows detailing upcoming payments.

Revenue becomes available to be paid out 90 days from the month in which it was earned.
​
The Upcoming Payment section will display the amount in green that has passed the 90 day mark and is ready to be paid out. The date of the next payment window will appear immediately below the Upcoming Payment amount.
​
The Unpaid Earnings section will contain all revenue that has been earned but has not yet passed the 90 day mark. Each month's earnings will move from Unpaid Earnings to Upcoming Payment as they they pass the 90 day mark.

### Disqus Basic Analytics {#disqus-basic-analytics}

**This article discusses analytics that are available with Disqus Basic and Plus Subscriptions. Information about advanced analytics that come with a Disqus Pro subscription can be found [here](#disqus-pro-analytics). Information about Ads and Revenue analytics can be found [here](#disqus-ads-analytics).**

Building a community is hard. Buckets of blood, sweat, and tears go into it. For over 3.5 million sites, Disqus is a part of that community building process. But sometimes it can feel like guesswork. Commenting and Ads analytics are designed to make the arduous process of increasing readership and tracking your earnings performance over time even easier.

These analytics cover three key areas:

-   Comments: number of comments, number of votes, and top comments

-   Revenue: revenue, viewable impressions, viewability %, RPMv, and clicks from using Disqus Ads

-   Payments: earnings history and payment schedule

Note that the Revenue and Payments pages will only display if your forum has Disqus Ads enabled. See Updating Your Advertising Settings for more information.

**Comment Activity**

Total comment and vote activity is available for your site in both daily and monthly view. You can select a custom date range or select a shortcut date range to track commenting activity over time in your community. You can also export a CSV of commenting activity for additional processing and reporting.
​

**Top Comments**

These are the comments from the last 7 days that have received the most positive voting ratios. This section can help you identify new and active users in your forum.

If you are interested in using your own analytics tool to capture commenting activity, see our help documentation here.

#### Reactions Analytics

**Total Reactions**

By default, the top graph of the Reactions Analytics page will show the total number of reactions left on the site by date. Since Reactions can be clicked by anyone, the data is broken out into Anonymous (logged out reaction clicks) and Authenticated (logged in reaction clicks). You may adjust the Date Range selector in the top right corner of the graph to change which dates appear.

Alternately, you may change the "Daily" option in the top graph to "Threads". The graph will then display the reaction count by thread for the 30 most recent threads.

**Thread-Specific Reactions**

While the top graph on the page will allow you to see total reactions by thread, the bottom graph will allow you drill deeper, seeing the breakdown of specific reactions for a given thread.
​
Selecting a date range in the top right corner of the bottom graph will populate the Threads section in the top of that graph, showing up to 30 threads published within the selected date range. Choosing a given thread will then show how many of each reaction were left on the thread. The reactions images used will appear above the graph for reference.
​

### Understanding Earnings and Engagement {#understanding-earnings-and-engagement}

Your Disqus revenue is closely tied to reader engagement. That means that the more readers comment on your site, the more viewable your ads will be, and the more you’ll earn. As you begin monetizing your site, remember to keep to track of both engagement and revenue to optimize your earnings.

You might already be tracking pageviews and subscribers, but what does that really tell you about your audience? Experts agree that engagement is a much better indicator of your site’s success than vanity metrics like views and followers.

1\. What should you track?

-   Disqus’s analytics tools let you track your site’s engagement through votes and comments. These metrics indicate the value, not just volume of visitors to your site.

-   Go to the “Analyze Engagement” page to view votes and comments by date. You can break down your engagement metrics by month or day, and select a specific date-range to see how engagement change over a given period of time.

2\. How to use engagement metrics

-   Look at your posts with the highest engagement. How do they differ from low-engagement posts? Perhaps you wrote about a controversial topic, used a provocative title, or posed a question to your audience. Pay attention to how your audience engagement changes for every post and use that information to develop more engaging content in the future.

#### Ads 101: Understanding revenue metrics

Disqus’s revenue analytics show you key ad performance metrics like RPMv, viewable impressions, viewability Percent, and total earnings. The closer you can get to 100% viewability, the higher your earnings will be.

1\. What should you track?

-   Take a look at your revenue metrics in the “Analyze Revenue” page. You’ll see your earnings, viewable impressions, viewability percent and RPMv—that’s revenue per one thousand viewable impressions broken down by date. As with the engagement analytics, Disqus’s revenue analytics page lets you select a specific date-range to see how your earnings change over a given period of time.

2\. How to maximize earnings.

-   Viewability percent gives you the most control over your Disqus earnings. The higher your viewability percent, the more you’ll earn. Since Disqus loads asynchronously, you can maximize viewability percent by removing any slow-to-load elements on your site. And remember, if there’s anything separating your comments section from the article (like a click-to-view-comments button or large blocks of other ads) your Disqus earnings will decrease significantly. Keep your earnings high by including the Disqus embed immediately after the article.

#### Get paid

Fill out your payment info at the bottom of the settings page. Disqus payments are made on a 90-day trailing basis once a minimum \$100 balance is met. That means your January earnings are paid in April, February earnings are paid in May, etc. You can view your upcoming payments and estimated pay date in the payments tab.

### Why am I seeing traffic from disqus.com/embed/comments in my analytics? {#why-am-i-seeing-traffic-from-disqus-com-embed-comments-in-my}

Disqus is fully housed in an iframe, loading from the disqus.com domain (0). This iframe is the element in which Disqus loads on any Disqus-enabled webpage.

As Disqus is an interconnected platform, this iframe effectively allows you to see both internally-recirculated and external referral traffic from Disqus embeds on your and other sites.

**What counts as a referral?**

Any time a user clicks a link in the Disqus embed and is taken to a new page, whether on the same site or another site, that counts as a referral. This includes:

-   Links in a discussion (i.e., the Discussion tab)

-   Links in the Community and My Disqus tabs

-   Links in profiles

-   Links in the [Recommendations](#recommendations) widget

**Do I have to navigate to another URL to generate a referral from Disqus?**

Yes. Note that this URL can be on the same domain as the referring discussion and it will still count as a referral.

**Can I see exactly which pages are referring traffic?**

Not currently; all referral traffic will show as coming from 0. This is something we're working to improve.

**If I click on a link pointing back to the same discussion, is that a referral?**

No. Clicking a link which stays inside the same, single Disqus embed instance will not generate a referral. For example: clicking on a profile and then a link in the profile which points to the same discussion being viewed.

## Commenting {#cat-commenting}

### Accessibility on Disqus {#accessibility-on-disqus}

-   We use WAI-ARIA tags/attributes.

-   We use best practices with keyboard navigation. Any modal box (profile modals) launched from the embed is marked properly as a modal and gets the keyboard focus. The whole embed is navigable via keyboard and tabs with proper alt/title text for screen readers (except for the image upload button for simplicity).

-   We hide some "decorative" elements to clean-up voice-guided keyboard navigation.

#### Screenreaders we recommend for the best experience:

-   VoiceOver

-   NVDA

**Please note that although we strive to make Disqus work well with screenreaders in all the most popular modern browsers, Disqus typically works best with screenreader technology in Firefox on Windows.**

#### What to do if you have encountered vision related accessibility issues while using Disqus (including those related to low-vision or visual processing disorders)

Please let us know by sending a report to accessibility@disqus.com

#### When sending a report, please include the following information:

1.  Your browser and version (please note that we may not be able to support all screenreader issues in IE8 at this time).

2.  You operating system and version.

3.  Your screenreader and version (if applicable).

4.  A summary of what issues you're encountering.

5.  Specific examples of issues related to features of Disqus (if applicable).

### Adding Images and Videos {#adding-images-and-videos}

#### Direct Uploading

Click the image icon at the bottom of any comment box and choose which file you'd like to upload. You can also drag-and-drop an image file directly into the comment box.

-   Max file size is 2 MB

-   Supported image formats: JPG, JPEG, PNG, GIF

#### Adding Gifs

To add gifs to your comment, you can click on the *GIF* icon at the end of the row of formatting options. This will allow you to search for a gif to add to your comment.

#### Rich Media Linking

Within your comment, paste the URL of the direct image file you’d like to include (e.g. 0). You should see a thumbnail appear if the image was detected. When you post the comment, the image is embedded in your comment.
​
We support embedding rich media links from the following services:

-   Youtube

-   Vimeo

-   X/Twitter (tweets)

-   Facebook (status, video, photo)

-   Instagram (photo only)

-   Bluesky Social

-   Giphy

-   Imgur

-   Google Maps

-   Soundcloud

-   Vine

#### How does it work?

Type your comment and copy and paste the URL of the media you’d like to include into the compose box. You can still use the image icon to upload still images from your computer.

When you click the post button, we’ll include the media in the post itself. This means that you can watch the videos your friends reference without leaving the comment screen.
​

#### How do I turn it off?

We understand that sometimes you’re not in the mood to check out lots of images and videos, so we’ve added the ability to hide media. To do this, click the gear icon at the top right above the comment stream, and select “hide media.”

This will hide all media items behind a link. Please note that if you go to another site and don’t want to see media, you’ll have to select this setting again.

After you’ve hidden media, your thread will look something like this, and you’ll be able to view media links one at a time by clicking the media link next to the comment.

#### For publishers:

If you allow your users to upload content or include media links in their posts, we encourage you to update your community guidelines. Rich media is a powerful tool for communication, and sometimes it can be abused. Setting clear expectations for commenting behavior can help ensure your readers have a positive commenting experience.

If you’re not seeing rich media at all on your site, you may not have media turned on. To do this, go to your publisher settings, and select the checkbox next to “enable media attachments.”

To disable, uncheck the box next to “enable media attachments.” Please note that it is not possible to disable certain types of media. Disabling and enabling applies to all rich media included in this feature.

#### F.A.Q.

#### What should I do if someone posts inappropriate content (YouTube, Twitter, etc)?

Flag the comment and report it to the site moderator. You can also block the user if you no longer want to see their comments.

#### Removing attachments

Removing attachments is not currently available. Please contact the moderator of the website if you'd like to have your comment removed permanently.

### Browser plugin/extension conflicts {#browser-plugin-extension-conflicts}

Some browser plugins may block Disqus from loading, and the following is a list of known ones. Note that these plugins may not cause a conflict after further updates, and it's best to try it yourself with and without these plugins enabled.

-   **Browser**: Chrome

-   **Reported**: 12/30/2013

-   **Resolution**: Enable "Allow social networks"

#### Avast Online Security

-   **Resolution**: Possible resolution in Windows – turn off "Do Not Track" *Source:* Avast! Forum.

#### Avast Antivirus 2014

-   **Reported**: 1/10/2014

-   **Resolution**: Enable "Allow social networks" – or allow Disqus individually

#### Ghostery

-   **Browser**: Chrome

-   **Reported**: 12/30/2013

-   **Resolution**: Toggle blocking for Disqus, or add an individual site to the Trusted list

#### KB SSL Enforcer

-   **Browser**: Chrome

-   **Reported**: 1/6/2014

-   **Resolution**: Add 0 to the list of ignored domains. Click on the KB icon to get to the options, type in disqus.com, and then click ignore.

#### IE Protected Mode

-   **Browser**: Internet Explorer

-   **Reported**: 2/21/2014

-   **Resolution**: Disable "Protected Mode" in Internet Explorer

#### Not listed?

If you've found any plugins that block Disqus which are not on this list, please contact us to let us know.

### Comment Text Formatting {#comment-text-formatting}

To format text within your comment, you can use the Text Editor buttons, our supported hotkeys, or markdown formatting. These can be applied to text that has already been written, or they can be used before typing new text.

Each of the Text Editor buttons can be clicked when typed text is selected, and it will be applied to that text. Alternately, you may click the button first, and whatever you type next will appear with that formatting applied.
​
Below is an image of how the text will appear with each of the first 8 formatting options applied:

Hovering over linked text will show the URL that the link points to. The grey rectangle is text that has Spoiler tags applied to it. Hovering over text with spoiler tags applied will render the text visible. More information on our Spoiler tags may be found here.

#### Supported Hotkeys

The following hotkeys are currently supported:

*(note: CTRL should be CMD/*⌘ *for Mac)*

-   ⌘/Ctrl + b for Bold

-   ⌘/Ctrl + i for Italics

-   ⌘/Ctrl + u for Underline

-   ⌘/Ctrl + s for Strikethrough

-   ⌘/Ctrl + Shift + 0 for Spoiler

-   ⌘/Ctrl + Shift + m for Code

-   ⌘/Ctrl + Shift + 9 for Blockquote

#### Markdown Formatting

The following markdown is currently supported:

-   \*text\* for Italics

-   \_text\_ for Italics

-   \*\*text\*\* for Bold

-   \_\_text\_\_ for Underline

-   \~\~text\~\~ for Strikethrough

-   \0 for Code

-   > text for Blockquote

#### Syntax Highlighting

Disqus supports automatic syntax highlighting in a number of languages. To use this feature, place your code inside \0 markdown symbols (this character appears on the same key as \~). Using these code markers will ensure that the code formatting is preserved.
​
For example:

    `var foo = 'bar';
        alert('foo');`

By default, Disqus will try to automatically detect the language.
​

#### Supported languages

-   Bash

-   Diff

-   JSON

-   Perl

-   C#

-   HTML/XML (*Note: You must first HTML-encode these tags to display them*)

-   Java

-   Python

-   C++

-   HTTP

-   JavaScript

-   Ruby

-   CSS

-   Ini

-   PHP

-   SQL

-   Spoiler tags

### Commenting 101 {#commenting-101}

-   Registering a commenter account

-   Verifying your Disqus account

-   Login to Disqus

-   Login to Disqus with a social media account

-   User profiles

-   Improving your user profile

-   Site Specific Profiles

-   Making your activity private

-   Subscribe/Unsubscribe from notifications

-   Deleting your account

#### Commenting

-   Adding images and videos

-   Remove and edit your comments

-   What HTML tags are allowed within comments?

-   Guest commenting

-   Mentions

-   Syntax Highlighting

-   How to: Get a direct comment link

#### Community

-   Your homepage on Disqus

-   Voting

-   Following other users

-   Remove a follower

-   User Blocking

-   Disqus web notifications

-   Sorting comments

-   Flagging comments

-   Featured comment FAQ

-   Disqus digests

#### Help

-   Which browsers does Disqus work with?

-   Use of cookies

-   Why is email auto-responder posting reply comments

-   How to get help and send feedback

-   Accessibility on Disqus

-   Windows phone app help

#### Troubleshooting

-   Browser plugin/extension conflicts

-   [Login troubleshooting](#logging-into-disqus)

-   Why can't I comment?

-   I'm seeing a "You are not allowed to perform this operation" error when commenting

-   Why are my comments marked as spam?

-   Troubleshooting email notifications

### Delete account or access account data {#delete-account-or-access-account-data}

In compliance with the General Data Protection Regulation (GDPR), Disqus provides users with options to access and fully delete all of the data associated with their accounts. Use the guide below to understand what type of user you are on Disqus and how to request access or deletion in each circumstance.

**Note**: If your website uses Disqus, visit Update on Privacy and GDPR Compliance for more information.

#### Types of users

#### Registered Accounts

If you own a registered Disqus account, you can use the self-serve deletion feature in Settings > Account > Delete Your Disqus Account. See **How to Delete Your Account** and **How to Request Data Access** below for more instructions.

#### Lost Registered Accounts

If you can no longer log in to a registered account that you suspect you owned in the past, reset your password by entering any email addresses that you believe may be associated with the account.

After successfully resetting the password, you can now use the self-serve deletion feature in Settings > Account > Delete Your Disqus Account. See **How to Delete Your Account** and **How to Request Data Access** below for more instructions.

If you have lost access to both your email account and your Disqus password, you will not be able to log in to Disqus to initiate the deletion yourself and we will do our best to assist you. See **How to Delete Your Account** and **How to Request Data Access** below for more instructions.

#### Users without Registered Accounts

**Guest Commenters:** You can unsubscribe from emails by replying to any notification with the word "unsubscribe" in response. See more ways here. You can also contact website moderators to delete your comments for you, if desired. If you would still like to delete or access your Guest Commenter data, we recommend registering a full Disqus account, then See **How to Delete Your Account** and **How to Request Data Access** below for more instructions.

**Site-specific Profiles**: Some sites have their own login systems which integrate with Disqus. You'll know these profiles because they can only be logged into/used to comment on that particular site. Since these profiles aren't managed by Disqus, you will only be able to delete or access data through the website the profile belongs to.

#### How to Delete Your Account

1.  At 0 in the upper-right, click the gear next to your avatar and then click Settings.

2\. Under the Account tab, scroll down to click the Delete button.

3\. You'll then be brought to a page to confirm your account deletion. Select a reason for account deletion, and then click the "Delete my account" button to confirm deletion and queue your account to be deleted.

#### How to Request Deletion of Your Account or Data Access

To request account deletion or an export of all user data on you stored by Disqus, please fill out the form here. You will be required to complete an email verification after completing the form. Once the form and email verification have been completed, your request will be satisfied, with either full account and data deletion, or a data export as selected.

#### What happens when I initiate account deletion?

-   When you delete your account, your account will be immediately deactivated and a full deletion will be completed within 30 days.

-   During the 30 day deactivation period, your Disqus account, including comments, will not be publicly accessible.

-   After 30 days, your Disqus data is not recoverable.

For more information on GDPR and Privacy, visit:

-   Update on Privacy and GDPR Compliance

-   Data Sharing Settings

-   Privacy FAQ

-   Terms of Service

-   Privacy Policy

### Disqus Digests {#disqus-digests}

The Disqus Digest email helps users stay engaged with the conversations and communities they care about. It’s available in both **daily and weekly** formats, based on user preference.

Each Digest includes:

-   **Personalized Content Recommendations**

    -   Discover new articles and discussions tailored to your activity on Disqus.

-   **Reply Notifications**

    -   A quick look at who’s replied to your recent comments.

-   **Username Mentions**

    -   Stay in the loop when someone tags you in a thread.

-   **Favorited Discussion Updates**

    -   Get notified when there’s new activity on threads you’ve favorites in.

-   **Disqus Stats Snapshot**

    -   See a summary of your Disqus engagement, including total comments, upvotes received, and more.

You will receive a Digest if “Receive Disqus Digest emails” is enabled in your email notification settings. Each new Disqus user is subscribed to the Digest by default upon registration.

#### How do I unsubscribe?

If you’d like to **unsubscribe from Digests only**, simply click on the “unsubscribe” link at the bottom of each Digest email. Also, we’d love to have your feedback. If there’s something we could do that would make you consider re-subscribing to Disqus Digests, please let us know.

If you’d like to **unsubscribe from all Disqus emails** (including Digests and notifications) see Managing Notifications.

If you’d like to receive only Digest emails (but no longer wish to receive regular notification emails), follow these instructions:

-   Navigate to disqus.com. If you’re not logged in, click “Login” at the top right and enter your username and password.

-   Place your mouse over your avatar at the top right to reveal a drop-down menu. Click on “Edit Profile”.

-   You’ll be brought to your Dashboard, and a dialog box will appear with various account settings. Find the “Notifications” tab on the left and click on it.

-   Uncheck the items which you no longer wish to receive regular notifications on. Be sure to leave “Receive Disqus Digest emails” checked if you wish to continue receiving Digests.

-   Once you’ve made your selection, click “Save Changes” to confirm your new settings.

Please note, if you’re receiving both Disqus Digests and Disqus notification emails, you may see the same content appear in both.

#### How can I provide feedback on Disqus Digests?

Whether there’s something you don’t like about Disqus Digests, or you have a feature request, we’d love to hear your thoughts. Please visit the Disqus Digest feedback form and let us know what’s on your mind.

### Disqus Web Notifications {#disqus-web-notifications}

Disqus makes it easy for you to add color to content across 3 million websites around the world. And each day, your colorful contributions have the opportunity to be recognized by millions of others through comments, replies, upvotes and follows. Notifications on Disqus alert you about those events. Here’s a breakdown of what you’ll see, and we’ll be adding new types of notifications in the future.

#### Provide us with your feedback on Discuss Disqus!

#### What can I be notified about?

#### Replies

Whenever another person replies to your comment. You can also reply directly inline without leaving the page.

#### Upvotes

When other people upvote your comments. If there are more than 1, they’ll be grouped together.

#### Follows

When other users on Disqus follow you, which means they can see what you comment on in their Home feed and email digests.

If you would not like others to follow you, you can make your profile private in settings. Please note that when logged in, your profile will still show your comments. To check if the profile has been correctly set as private, try logging out and visiting your profile page.

#### Where can I see notifications?

Besides the normal email notifications you receive for replies, you can also see notifications in the Notifications page on Disqus.com:

#### How do I filter the notifications I see in my Inbox?

There are currently two Inbox views. "Most Recent" includes replies, follower notifications, upvote notifications, and invitations to join discussions. You can filter out upvote notifications and discussion invitations from the "Most Recent" view by clicking the Inbox settings gear and checking the corresponding boxes. The "Replies" view includes just replies.

You can also determine which actions trigger Web Notifications from your account settings page:

#### Brandless Web Notifications

For sites with a Business tier subscription, we can also support brandless web notifications. Instead of linking to the Disqus.com notifications page, a brandless sidebar will slide over the page to display all web notifications for the account.
​

For more information on brandless web notifications, please contact your Disqus Account Manager or request information from our team here***.***
​

### Featured Comment FAQ {#featured-comment-faq}

Featured comments are a way to highlight comments in a conversation. Use it as a way to guide conversations and reward quality contributions.

You can feature any comment within the discussion, even if it’s a reply to someone else. When you feature a comment, it will be displayed prominently at the top of the thread.

#### How do I feature a comment?

The controls for featuring a comment are located in the comment dropdown; this is the same dropdown that you use to moderate comments from within the thread.
​
​

#### How do I stop featuring a comment?

When you’re ready to take your featured comment down, you can stop featuring it by using the same dropdown menu. You can also stop featuring a comment from within the comments section. Locate the comment that is featured, hover over it and use the moderator drop down tools to select ‘stop featuring.”
​

​

#### Can I feature more than one comment?

No. Only one comment can be featured at a time.

#### Will the featured comment also show up in the regular comments section?

Yes, the featured comment is a highlight of the original comment. When you scroll down into the comments section, you will see the original comment, and any replies to it.

#### Are commenters notified when their comment is featured?

Yes, the author of the comment is notified when a comment they posted is featured on a site.

### Flagging comments {#flagging-comments}

Flagging a comment tells a site moderator that a comment requires moderator attention. On most sites you can flag a comment by clicking its flag icon or link.

Flagging a comment is only counted once per person; you won't need to do it multiple times.

Every site has a different commenting policy, be sure to review it before flagging comments.

Generally, comments **should** be flagged for:

-   Spamming

-   Violating a site's commenting policy

-   Being clearly unrelated

-   Attacking other commenters personally

Generally, comments **should not** be flagged for:

-   Disagreeing with the content

-   Disputing with other commenters

***\*\*Moderators may be able to see if specific users are abusing the flagging function***

#### Report Spam to Disqus

We also appreciate reports of any spam comments you find. After flagging a comment for the moderator, feel free to click on their profile and flag the user.

-   Who deleted or removed my comment?

### Following other users {#following-other-users}

Following users is a great way to discover new content, and be part of the most recent conversations happening in your network.

Many comments posted through Disqus are hilarious, insightful, or just fascinating to read. Following other users keeps you up-to-date with the latest news that's relevant to you. After you choose to follow someone, his or her activity will show up in your home feed.

#### Follow people that interest you

When you follow people on Disqus, their conversations and activity will populate your Network tab . You'll see what they comment on most, what they upvote, and where the conversation is happening.

#### How do I follow/unfollow others?

You can follow a user by expanding their profile (by clicking their avatar) and selecting the "Follow" button.
​

#### How do I find people to follow?

You can follow other Disqus users the way you normally would from the discussion or on someone's profile, but we've also added some new ways to follow people.
​
If someone starts following you, or they reply to one of your comments, the you can follow them from within your notification feed by clicking the follow button next to their name.
​

#### How do I block people from following me?

While the functionality to block followers doesn't currently exist within Disqus, you can Remove a Follower and you do have the option of making your activity on Disqus private. Additional guidance on what to do when situations on Disqus become abusive can be found in the Abuse Overview.

### Guest Commenting {#guest-commenting}

Guest commenting is an optional feature in Disqus that allows users to comment without creating a Disqus profile.

Commenting as a guest is different than commenting with a registered Disqus account in a few main ways:

**Please note:** We no longer support Gravatars associated with the email address used when making guest comments.

1.  Neither a Disqus account nor profile will be created using your credentials. This also means all benefits of having a registered account, such as being able to customize your profile, change settings, edit/delete your comments, are not included when commenting as a guest.

2.  Commenting as a guest also means you will not be subscribed to email notifications of new comments, including replies to your comments.

3.  Guest comments will not be approved automatically, and must be manually approved by a site moderator to appear on the page, regardless of whether Pre-Moderation has been enabled.

#### Is guest commenting the same as anonymous commenting?

No. Guest commenting is the ability to comment without registration. You may do so anonymously, with a pseudonym, or with your real identity.

Anonymous or pseudonymous commenting with registration remains available as always.

#### How do I enable guest commenting on my site?

Guest commenting can be enabled at the Disqus admin > Settings > Moderation in the Guest Commenting setting.

#### How do I know if a site has guest commenting enabled?

You'll see a checkbox below the Password field when posting a comment that says "I'd rather post as guest".

#### How do I comment as a guest?

When commenting, simply enable the "I'd rather post as guest" checkbox. Your credentials will not be used to create a Disqus account or profile, nor will you receive any email notifications.

#### Why can't I comment as a guest?

Not all sites in Disqus allow guest commenting. If you'd like more information on why guest commenting isn't available on a site in question, kindly contact that site's moderators.

If you are unable to comment on a site with guest commenting enabled, check out Why can't I comment?

#### What happens to my guest comments if I make a Disqus account?

Comments you make as a guest will remain separate from your Disqus account. Guest comments cannot be claimed.

### Mentions {#mentions}

With \@mentions, you can tag and link to the profile of a user who is following you. It provides a good visual queue for directing your comment at a specific person and it gets their attention by notifying them of your mention.

Just type the @ symbol and then continue typing the name of the person you’d like to mention. As you type, Disqus will show a drop-down list of suggested users. The drop down is smart — it will update with increasingly accurate suggestions the further you type.

#### Who can I mention? Can I mention people not already in the conversation?

You can mention any user that follows you, whether they are in the conversation or not. Mentioning a user that isn't in the conversation will notify them of your mention, and give them an opportunity to join you in the discussion.

#### Can I disable this? How?

Go to your Web Notifications settings.

#### Can I mute mentions from just a single user?

Blocking can achieve this, but it also hides all other commenting activity from that user, in addition to mentions.

#### What if someone tries to abuse this?

You can always block a user who disrupts your experience on Disqus.

#### I don't like that I get duplicate notifications if a reply to myself contains a mention.

This is a known issue right now. We're looking to have this resolved in the future. We appreciate your patience :)

### Remove and Edit Your Comments {#remove-and-edit-your-comments}

#### As a registered commenter

Registered Disqus users can edit their own comments by clicking the Edit link within 7 days after posting. After 7 days, editing is permanently disabled. Disqus does this to help make comments less susceptible to spam abuse.

#### As a guest commenter

Guest commenters should contact the moderator of the website where their comment was originally posted to request that it be edited. Editing is also disabled for moderators for comments older than 7 days. It's up to each site how they manage their discussions, including how to respond to commenter requests.

#### With a site-specific profile

Comments made with a site-specific profile can only be edited on the original page where they were posted. Comments made by site-specific profiles cannot be edited in Disqus Home.

#### Remove comments

#### As a site moderator

See Moderating your community.

#### As a registered commenter

Registered Disqus users can remove their own comments from public discussions and their profile by deleting them.

-   Once a comment has been deleted it cannot be claimed again.

-   Comments are not anonymized when they are deleted and can still be seen by moderators.

**How to delete a comment:** hover your cursor over the comment you want to delete > click the actions dropdown > click Delete.

To **delete a comment appearing in your profile**, click "View in discussion", then delete it from the discussion as described above.

#### As a guest commenter

Guest commenters should contact the moderator of the website where their comment was originally posted to request that it be removed. It's up to each site how they manage their discussions, including how to respond to commenter requests.

#### With a site-specific profile

Comments made with a site-specific profile can only be removed on the original page where they were posted. Comments made by site-specific profiles cannot be removed in Disqus Home.

-   Registering a Commenter Account

-   Guest Commenting

-   Site-Specific Profiles

-   Who deleted or removed my comment?

### Subscribe/Unsubscribe from Notifications {#subscribe-unsubscribe-from-notifications}

By default, you receive notifications for all replies to your comments. You can also subscribe to entire discussions on which you comment under the Email Notifications settings.

Additionally, anyone can subscribe to individual threads via RSS or email in any Disqus embed by clicking one of the subscription links at the **bottom** of the embed. The option will change color to denote that you've been subscribed.

#### Registered commenters

Registered users can turn off new thread subscriptions at Edit Profile > Notifications > Personal Settings

or from specific threads by clicking one of the subscription links at the **bottom** of the embed. The option will change to gray to denote that you've been unsubscribed.

#### Guest commenters

All users (including Guest commenters) can reply to a Disqus notification email with the keyword "**unsubscribe**" to stop notifications for that specific thread or "**unsubscribe all**" to be unsubscribed from all Disqus notifications.

Alternatively, Guest commenters can click the "Stop receiving notifications" link at the bottom of an email notification.

### Two-Factor Authentication (2FA) {#two-factor-authentication-2fa}

Two-Factor Authentication (often abbreviated as 2FA), is a way to add an additional layer of security to your Disqus account. This is accomplished by requiring multiple forms of login to be completed before access to the account is granted.

With 2FA enabled, after you enter your account password, you'll be prompted to complete an additional step for login. This could be entering a code sent to your email address, or opening an authenticator app and entering the code supplied there into Disqus. Once the additional login measure has been completed, you'll be let into your account as usual.

A more in-depth guide to Two-Factor Authentication can be found here.

#### Setting up 2FA in Disqus

To set this up for your account, you'll want to navigate to the Two-Factor Authentication section of your account settings page. There, you can connect the method you'd like to use for your second layer of authentication.
​
If you have an authenticator app that you'd like to use, you can simply scan the QR code present on that page with the authenticator app on your mobile device, and follow the additional prompts to connect the app with Disqus.

#### Which Authenticator Applications are supported?

Disqus 2FA should work with all authenticator apps that send a Time-Based One Time Password (TOTP). As long as your authenticator app of choice supports TOTP as an authentication method, it can be used.

#### Email Authentication

If you'd prefer to authenticate via email instead, this is also supported. After logging in with your password, you'll receive an email with a numeric code and be redirected to a page where you can enter this numeric value. Simply enter this code into the "Code" field in Disqus, and you'll be logged into your account.
​

#### Backup Codes

In addition to the authentication methods, the two-factor authentication section of your account settings page will also provide options to generate backup codes for access to your account.
​
We strongly recommend generating these backup codes, and storing them in a safe place in your local digital system. If for some reason you lose access to the authentication methods attached to your account, you will need to use these codes to regain access to your account.
​
Because 2FA is a security measure adding additional login security to your account, we cannot manually go around this to provide access to an account in cases where the authentication methods have been lost. Backup codes will be the only option for account access in cases where access to the 2FA authentication methods have been lost.

​

#### Frequently Asked Questions

**Can I add both Authenticator App and Email Authentication to my account?**

Yes, you can enable both Authenticator App and Email Authentication to your Disqus account at the same time. When both are enabled, we will default to using the authenticator app.

If you have both installed and would like to use your Email Authentication, you can click the "Verify by email" option. This will send an authentication email and open the correct field for the numeric code sent to the email address on your account.

**Can I add multiple Email addresses to my account?**

No. At this time, we only support one email address on each Disqus account, for login and email authentication.

**What happens if I enable 2FA and lose access to my account?**

If you lose access to your email address but still have access to your authenticator app for 2FA authentication, you can still log in with your old email and password value, to update the email address on your account and verify the new address. If you've forgotten your password, a password reset email may be requested from 0.

However, if you lose access to your 2FA authentication methods, your only option for access will be to use previously generated backup codes for access. Because of this, we strongly recommend generating and storing backup codes immediately after setting 2FA up on your Disqus account. In cases where the 2FA authentication methods have been lost and no backup codes have been generated, we will not be able to provide access to the account, and you will only be able to access your account if you have already generated and stored backup codes for account access.

### User Blocking {#user-blocking}

User Blocking is a feature that allows you to deal with trolls, spammers, and other unwanted content on Disqus. Once you’ve blocked a user, their comments will be hidden from your view, and your comments will no longer be visible to that user. The blocked account will not receive a notification or indication that they have been blocked by you, and they will no longer be able to view any of your comment content.

User blocking may be accessed either from the profile of the account you would like to block, or from the dropdown menu appearing on each of their comments.

-   On a user's profile, the blocking option may be accessed in the menu next to the Follow button for their profile:

-   In a comment thread, the blocking option may be produced by clicking the triangle dropdown menu in the top right corner of their comment:

#### The Impact of Blocking a User

Once a user has been blocked, all of their comments, discussions, and recommends will be hidden from you throughout your Home feed, Inbox notifications, and discussions on external websites. Any posts by the user will be collapsed with a message stating "User is blocked".
​
Additionally, the blocked user will no longer be able to view your comments. They’ll see ”Content unavailable” when viewing your posts or your profile.

#### Manage Your Block List

Any user that you block will be added to your block list which is located in Settings > Blocking. Use the Previous and Next buttons to change pages, and use the Unblock button to remove a specific user that you no longer want to block. The maximum number of users you can block is 1,000.

#### F.A.Q.

**Q: If I block someone, will they be able to see my posts?**
*A: No, a blocked user will not be able to see your posts. If you block a user, this will hide all of your comments from that user’s view.*

**Q: If a moderator blocks someone, will they be able to see their comments in the moderation panel?**
*A: Blocking a user will not affect anything in the moderation panel. All comments will still be visible in the moderation panel, though blocks may change what content is visible in the embed.*

**Q: If I block someone, will I still see if they upvote me?**
*A: Your posts will be hidden from the blocked user, so you won't receive additional upvotes from that account.*

### Voting {#voting}

**Who can vote on comments?**
Guests and logged out users can no longer vote on comments. Login or register a Disqus account to vote.

Below each comment there are upvote and downvote buttons which will add your vote to the tally.
​

#### Why vote on comments?

Voting actively will increase the engagement with your fellow commenters without adding a comment. This will help create an incentive to post content which you approve of.

#### Can I find out who voted on my comments?

To view the people who have voted on a comment, hover your mouse over the vote icon which will reveal a box of users who have upvoted or downvoted the comment. Here's a screenshot of how this looks:

​
You can also view your total number of received upvotes by hovering your mouse over your avatar when it appears with a comment. A tally of the number of upvotes or downvotes per comment is also available next to the voting buttons on any comment.

#### Didn't mean to vote on something?

You can undo any vote by clicking the button again.

#### Who can vote on comments?

Only logged-in users can vote.

#### Why is my upvote count decreasing in my profile?

When a Disqus account is deleted, all comments and votes are also removed. Your upvote count will fluctuate over time if you have votes from users whose accounts have been deleted.

#### Hiding the downvote count and list

Showing the downvote count and list of downvoters for a comment may not be appropriate for all communities. If you’d like to hide the count and list for downvotes for all comments, you may disable this by checking the Downvote Details option at the bottom of your Community Settings page.
​

### Who deleted/removed my comment or blocked my account? {#who-deleted-removed-my-comment-or-blocked-my-account}

Is your comment not appearing in the thread or are you unable to post? The best way to learn more is to check in your comments feed within your profile for the status of the comment or contact the site moderator who will have more information on why your account was blocked.

If you are banned across multiple sites, it is likely that your username, email address, or IP address was globally banned by Disqus, usually as a result of being detected as a spam account. If you suspect that you have been wrongly banned, try the following:

-   use a different email service (disposable email services used commonly by spammers will be blocked).

-   use a different account (usernames used by spammers will be blocked).

-   use Disqus from a different IP address (IPs used by spammers will be blocked).

#### If you are banned on a single site

This message indicates that your account has been banned on a single site by a site moderator. If you are only banned on one site, try the following:

-   contact the site moderator for more information about why you were blocked.

#### If a site moderator has given you a timeout

This message indicates that your account has been placed in timeout by a moderator. Timeouts are temporary, and last for any amount of time as specified by the moderator.

-   Contact the site moderator for more information about why you were put in timeout.

#### How to contact a moderator

Most websites have a contact form or email address listed to get in touch with their moderators or support team. To find a site's contact page more quickly, it can help to google: \[website name\] "contact us".
Please note that sites are not required to provide contact information and in cases where the site doesn’t provide contact information, Disqus won’t be able to put you in contact with that site’s moderator.

#### If you have posted a comment that is not showing

If your comment was posted but isn't appearing in the thread, check your profile > comments feed to determine why the comment is not appearing. Your comment could be...

If you see one of the above messages next to your comment, try reaching out to the moderator of the site where you posted the comment; site moderators are the only ones with the authority to approve a comment or explain its status.

If none of these messages appear next to your comment, there could be a lag in duplicating this data across the Disqus system, so try revisiting the thread at a later time to see if the comment is visible.

#### Reasons a comment may be awaiting pre-moderation

Websites using Disqus have the ability to pre-moderate comments based on certain criteria. If a comment meets these criteria, it may be queued for moderator review before being published publicly on the website. Reasons a comment may be pre-moderated include, but are not limited to:

-   Contains a link.

-   Contains a media attachment, e.g., image or video.

-   User has not verified their email.

-   The current discussion thread is set to pre-moderation.

If your comment is awaiting pre-moderation, it will appear as **pending** in your profile > comments feed.

#### Disqus doesn't moderate individual sites

Disqus does take action across our network on comments, profiles, and discussions that violate our Basic Rules, including spam.
​
However, when it comes to individual sites and communities, Disqus takes no part in moderation decisions (e.g., approving comments, deleting comments, or handling disputes among commenters) nor can we offer an explanation as to why a comment or a user account has been moderated. These decisions are made by websites using the Disqus service (the "site moderators").

Commenting experiences can differ greatly depending on the individual site's moderation practices/policies. Policies and settings can vary widely between Disqus-enabled sites.

-   Site Moderators

-   Removing and Editing Your Comments

-   Why are my comments disappearing?

### Why are my comments marked as spam or removed? {#why-are-my-comments-marked-as-spam-or-removed}

See our Why are comments visible in the Disqus admin but not on my site? documentation.

#### For commenters

#### Comments marked as spam

Comments are removed from public view (or "disappear") when they are filtered as spam, whether manually by a site moderator or automatically by our system.

The following commenting behavior can cause comments to be marked automatically as spam:

-   Including a signature in multiple comments. For example, appending a name or website link to the end of multiple comments. We recommend entering that information at Edit Profile > Profile tab instead.

-   Bad or strange syntax. For example, excessive paragraph breaks, bad punctuation, double-spaced comments.

-   Posting the same comment multiple times to the same page, or across sites;

-   Using multiple links in one or multiple comments.

To prevent this from happening further, contact the moderator(s) of the site(s) in question where your comments are disappearing. They'll want to first edit your comments to remove the relevant text and then approve your comments. If you've posted duplicate versions of the disappearing comment, we'd recommend the moderator(s) approve only one of the duplicate comments and delete the rest in order to keep those comments from being counted against you.

#### Other non-spam issues

#### Your comment is awaiting moderator approval

Please see Who deleted or removed my comment? for more information.

#### Your comment has been removed by the site moderator

Please see Who deleted or removed my comment? for more information.

#### Sort order not set to "Newest"

Comments may seem to disappear if the comment thread has many comments and the sort order is set to something other than "Newest". Set sort order to "Newest" when posting a comment to keep your comment at the top of the comment thread, at least until further comments are made. Please see Sorting Comments for more information.

-   Who deleted or removed my comment?

### Why can't I comment? {#why-can-t-i-comment}

There are a few reasons why you may be unable to comment:

See Why isn't the comment box loading?

#### "We are unable to post your comment..."

See We are unable to post your comment....

#### General Troubleshooting

We'd recommend the following troubleshooting steps for users experiencing issues while trying to comment.

-   Disable all plugins, extensions (Chrome), and add-ons (Firefox), in your browser, as well as any privacy-related software

-   Clear your browser's cache and cookies, as instructed here

-   Enable third-party cookies, following this documentation

-   **Firefox:** If you have selected the "Strict" Privacy setting in your preferences, please ensure that the "Fix Major Site Issues" checkbox is selected at the bottom of that option. If that checkbox is not available, you may need to select the "Standard" Privacy option for Disqus to function correctly

-   **Safari:** In the Privacy settings, "Website Tracking - Prevent Cross-Site Tracking" must be *unchecked* for Disqus to function correctly

    -   If you have issues staying logged in with Disqus on Safari, navigate to your Privacy settings, click the "Advanced Settings" button in the bottom left corner, and ensure that the "Block all cookies" option is *unchecked*. Cookies must be allowed for Disqus to keep you logged in between sessions

#### Still having trouble?

In order for us to best help you, contact us within the following details for what you're experiencing.

-   A brief description of the issue

-   Link to any page where you saw the issue

-   Screenshots that illustrate the problem - How do I take a screenshot?

-   Which browser you used - What's my browser?

### Windows Phone app help {#windows-phone-app-help}

#### Version 1.3

In-progress comment won't disappear when locking the screen; New mobile-friendly edit profile page; Notifications will un-highlight when clicking dismiss all; Bug fixes and new languages.

#### Version 1.2

New inline browser; A link to your recent discussions; Images now show up for threads (when available); Some bug fixes and new translations.

#### Version 1.1

Some bug fixes, added a Dutch translation and improved Italian translations.

#### Version 1.0

This is the first version of the Disqus app! Send us your feedback.

#### Getting Started

#### What is DISQUS?

Disqus is a network of thousands of live discussions across the web on many of the largest websites in the world, and covering any topic imaginable.

#### What does this app do?

The Disqus app will help you keep up with your favorite communities, find new discussions, and manage your identity all on the go. You can pin discussions or custom searches right to your home screen for easy access later as well.

#### Logging in

You can log in to DISQUS using the same credentials you've registered with on disqus.com, or any Disqus-enabled website. If you're not sure if you have an account you can get additional help in our documentation.

#### Registering an account

You can register for an account right in the application, and you can use this same account on the web version of Disqus as well.

#### What's next?

1.  Fill out your profile.

2.  Find awesome communities.

3.  Follow interesting people.

#### Using DISQUS

#### Your profile

Your profile is your own space on Disqus and allows other users to know what type of commenter you are. You can enter your display name (shown next to your comments), your website URL, a location and a brief bio for others to read.

#### Following users

You may follow other commenters who you're interested in keeping up with. Following will populate your network activity with a feed of their latest comments anywhere on the Disqus network.

#### Finding communities

There are thousands of highly active communities on Disqus covering any interest across many different regions. You can find some of our top communities through the explore topics, search, through user profiles and by following more users who share your interests. You'll be considered active on a community by commenting, voting on, or starring discussions on that site.

#### Notifications

You'll receive a notification for announcements from Disqus or when someone replies to your comments. If you've pinned Disqus to your start page, we'll occasionally check for notifications and update the live tile.

#### Discussions

#### What is a Discussion?

Discussions are threads of comments that belong to a certain topic. Around the web you'll find Disqus as a thread on each article or post on a site and is the primary way to interact with Disqus.

#### Leaving a comment

Make your opinion known! If the site allows it, you can leave a comment through the app on a discussion. You can either leave a top-level comment, or reply to another user.

#### Starring threads

Starring discussions are a good way to participate without commenting. This will put the site on your active communities area, and also populate "active threads" in your network feed so you can stay up-to-date on comments

#### Comment voting

Voting on comments help surface the top comments in a discussion. In addition, upvoting a comment will aggregate it to your followers' dashboard feeds. Downvoting shouldn't be used for inappropriate or spam comments, flagging the post would be more appropriate

#### Sharing

You can share any comment or discussion to the social networks you've configured on the phone. This will link people to the web version of Disqus.

#### Commenting

#### What can I comment on?

Commenting is enabled on virtually every site in the network. A select few require that you visit their site in order to create an account with them, but you only need to do this once.

#### Commenting Rules

DISQUS imposes no broad rules on commenting, but many communities have their own rules that can result in your comments being removed, pre-screened or get you banned. Visit each site to get a sense of what the community rules are before participating.

#### Who do I contact about my comments?

The moderator of the website owns the comments, and should be contacted for any questions. DISQUS is here to help with any technical issues you have with the system.

## Community Tips {#cat-community-tips}

### Best Practices for Moderating Sites {#best-practices-for-moderating-sites}

Disqus offers a range of Moderation tools to help communities deal with antagonistic or spammy posters. Typically, adding users to the Banned User list will ensure that they no longer post at your site. However, with more dedicated troll and spam accounts, additional measures may be necessary to curb their destructive actions. Below is a list of actions and tips to help with toxic commenters, repeat impersonators, dedicated trolls, or any other users you don’t want in your community.

#### Ban culprit accounts

When moderating, you'll want to use the Priority Sort filter to surface the most troublesome commenters. This filter uses a special calculation of flags, downvotes, restricted words, reputation, and guest account status to determine which comments will likely need swift action. Priority sort can be found next to the existing “newest” and “oldest” options within the sort menu in the top right navigation of the Moderation panel.

After you have identified any users that are breaking your site’s community guidelines, we recommend taking action on their accounts. As a moderator of the forum, the most important step is usually banning culprit email address since that is the central identifier used for accounts in our system. This prevents the user from creating a new account using the same email address. Similarly, you can consider banning the IP address. This can be helpful when you know that multiple users are using the same IP address, but different email addresses.

#### Turn on pre-moderation and use Trusted User list

If the unwanted content or users continue to be a problem in your community, you may want to turn on the Pre-moderation setting for all users, which will require all posts to be explicitly approved by a moderator before appearing on the page, with the exception of users who have been added to the Trusted User list. If you are subscribed at the Pro level, you have another Pre-moderation setting available to you: New Commenter Pre-moderation. This setting will require moderator approval for all users that are new to your site, without restricting posts from your returning commenters. This means that you can vet users that are new to your site, and you can customize how many days you want them to be in the approval state. This tool is incredibly effective for dealing with Spammers as well.

For more active sites, we recommend adding your most trusted and frequent commenters to the trusted User list, so that moderators can spend more time reviewing comments from new users or more troublesome accounts.

Add accounts to your Trusted User list in Moderation → Banned Users → Trusted. It’s even possible to programmatically add and remove trusted/banned users using the API if you have data like “subscribers” on your site’s end that you’d like to pass to Disqus (see the blacklists and whitelists endpoints).

For ideas on who to add to your trusted user list, view your site’s top commenters in your site’s community profile at 0 or in your Analytics at Community → Top Comments.

After making use of the Trusted User list, you can consider toggling the pre-moderation setting to “All”. This will ensure that all comments will require moderation, with exception to your trusted users’ comments which will become approved immediately.

#### Adjust your forum settings

#### Enable comment flagging

Lower the flagging threshold so that comments which are flagged X number of times are removed from public view. This can help your community moderators by allowing your users to surface and hide comments most in need of moderation. The flag settings may be adjusted at Settings → Moderation → Flagged Comments.

#### Use the Restricted Word List

#### Stopping Spam

If you have identified any recurring spammy URLs or words in these comments, you can add those strings of text to your forum's [Restricted Words](#moderation-settings) list. Any comment containing one of these strings, for example "money.net/click", will be automatically placed in the Pending filter and will require moderator approval before it is public.
​
When comments are posted containing a Restricted Word, that word will be highlighted in the Pending filter.

Learn more about how to use the Restricted Word list effectively.

#### Impersonation

In Disqus, accounts must use unique usernames (which appear in the URL of their profile), but are allowed to use the same display name, so that two users named Bob may have just “Bob” appear on their profiles and could be distinguished by their account avatars.

Impersonation occurs when one user maliciously copies both the display name and avatar of another account, making their copy account indistinguishable in order to post false content.

In general, we recommend that Impersonating accounts be added to the Banned User list to prevent them from further posting to the site. If they continue to create new impersonating accounts, the Restricted word function can be used in conjunction with the Trusted User list to combat repeated impersonation.

The Trusted User list will override the Restricted word filter, a Moderator may add the original community member to the Trusted User list, to ensure that their comments will appear successfully on the site. If the display of the account is also added to the Restricted Word filter, this will cause comments from all other accounts besides the original account to be set to Pending, requiring explicit moderator approval before showing publicly. This will allow moderators to stop all impersonating comments from appearing publicly on the site.

To implement this, you may add the display name being copied to your Restricted Word filter, and add the original user that is being impersonated to your Trusted User list. This will allow the original user to post successfully without moderation, and will set all comments from other accounts using that exact name as Pending until they've been explicitly approved by a moderator.

#### Add more moderator users to help moderate

Adding more moderators will make moderation faster and more comprehensive across your community. Encourage moderators to be more active in the community by commenting regularly. Having your moderators participate in your own community is one of the best ways to grow engagement, set examples for your community, and have a stronger moderation presence.

Learn more about getting your authors and moderators involved in the comments section.

#### Add community Guidelines

Community guidelines are a place for you to set the tone of your community and lay a foundation for what is acceptable. This can be communicated to commenters by adding visible community guidelines to your site's HTML, directly above the Disqus embed. More information, see our 4 Best Practices for Integrating Community Guidelines in your Site.

#### Automatic Closing of threads

Consider turning on automatic closing. This feature closes threads after a set amount of days which can help your moderation team focus on recent articles only. Recent articles are most likely to be where the highest volume of traffic and engagement is likely to be occurring.

#### Disable “Allow guests to comment”

Consider disabling “Allow guests to comment” if your goal is to tighten control over the quality of comments that you receive. Guest comments sometimes require closer moderation. If you’re more concerned about receiving higher volumes of comments, you can leave the setting as you have it. More information here.

### Build Your Community {#build-your-community}

Add a comment count to articles on your homepage so readers know where discussions are happening. Make it easy to find Disqus on your stories, even if you hide comments behind a click.

#### Ask readers to comment

If you want to have a conversation, you have to start it. End your stories with a question for readers to show that you’re interested in hearing from them.

#### Reply to commenters

Be active and present in the comments by replying to readers, especially first-time commenters. Author participation builds trust and strengthens your community over time.

#### Feature the best comments

Highlighting top comments has been shown to increase overall engagement by 30% according to research by the Engaging News Project. You can feature comments or embed them directly in stories you publish.

#### Get to know your readers

Readers will visit your site for the content, but they’ll keep coming back for the conversation. Analytics helps you measure the growth of your audience on Disqus over time and understand which stories generate the most engagement and why.

*💡 Looking for more? Check out our* *Ultimate Guide to Increasing Reader Engagement*

1.  Listening to your community

2.  Turn comments into content

3.  Host events

4.  Promote comments outside your website

5.  Recognize top commenters to retain them

6.  Lower the barrier to commenting

**Previously:← Create and Enforce Moderation Rules**

### Comment Policy {#comment-policy}

Your Comment Policy is a place for you to set the tone of your community and lay a foundation for what is acceptable. They also provide a reference for making moderation decisions. If you ever need to ban a member, or remove a comment, you’ll be able to refer those members to the pre-determined community guidelines that they may have failed to follow.

When developing community guidelines it's good to consider the type of community you're trying to cultivate. Guidelines can cover topics like:

-   **Etiquette** - "Be polite and stay on topic”. "No self-promotion". "Don't flag/downvote comments because you disagree with a user."

-   **Expectations** - "Your comment will be removed for reason X, Y, and Z"

-   **Privacy** - “Don’t post personal information"

-   **Moderation Settings** - "Comments containing links are pre-moderated". "Discussions auto-close after 7 days". (Inform users about your forum's moderation settings)

-   Anything else that you’d like members to keep in mind while commenting

####  How to add your Comment Policy

You can add your Comment Policy in your General Settings. The Comment Policy Summary is where you'll provide a quick snapshot of what is accepted/prohibited behavior. You can also add a link to your full policy on your website in the Comment Policy URL field.
​
Once added, your comment policy will appear formatted as below:

#### Get started using our Sample Community Guidelines.

### Connect Your Content to Your Community {#connect-your-content-to-your-community}

There are many factors that go into creating content that generates great discussion –– maybe it’s a provoking photo, some well-timed commentary on current events, or a philosophical forray that makes your readers' head explode.

And we know that the last sound you want to hear after skillfully crafting your content and releasing it into the Universe is:

\*crickets\*

**One way to spark the discussion about your content is to explicitly ask your visitors what they have to say.** It’s a simple practice, and successful bloggers agree that it's important in showing your community that you care. We have seen first-hand how addressing your community directly in your content can help seed a quality discussions.

Here are some fine examples for inspiration:

#### Wendy's Lookbook

You can blur the content-community lines even further by asking people to submit their own artwork/reaction-gifs/poetry in the comments! Check out how contemporary art blog 0 calls upon the masses every week to showcase their photography skills.

**Enough from us though, we absolutely want to hear what you have to say. :)**

-   In what ways do you connect your content with your unique community?

-   For which type of content does this work best?

-   Share a link to an example on your site where you have tried this!

### Create and Enforce Moderation Guidelines {#create-and-enforce-moderation-guidelines}

Set the tone of your community and remind readers about the guidelines for commenting on your site. Enlist the help of your community to report abusive content by flagging comments.

#### 2. Configure Pre-moderation Settings

If comment quality is a concern, you can turn on Pre-moderation for some types (e.g. comments containing links) or all comments. This sends comments to Pending and will only be published once a moderator has approved it.

#### 3. Automatically close discussion threads

It might make sense to automatically close discussions after a set period of time, especially if moderating comments on older posts takes too much time. Generally, we recommended closing threads after 14 or 30 days.

#### 4. Create a Restricted Words list

Any time a comment or name contains a word you've specified in this filter, it will automatically be flagged for moderation review before it can be published. You can use the default sample list provided or customize with your own words.

#### 5. Add Moderators to Your Team

You can’t just add comments and expect a community to emerge. Active moderators are essential to cultivating a healthy environment for diverse opinions. Add members from your team or recruit readers from your community to help moderate discussions.

**Previously:** ← **Personalize Your SiteNext up:Build Your Community** →

### Default avatars for communities {#default-avatars-for-communities}

As a site administrator you can add a default avatar for commenters who don't have their own yet. For any Guest comments or comments by profiles without an avatar set, your custom image will show as the avatar for these users. This adds a nice touch to the Disqus embed to make it feel like a part of the community.
​
To override the default Disqus commenter avatar, go to your Disqus admin, click the **Settings** page and then the **General** sub-tab. The option will be under **Default commenter avatar**. Click the Plus button to add a new image, ensure that it is bordered in blue, and click Save.
​

Tips

-   Keep it simple. Use an image that anyone in the community can associate with, and that doesn't distract from their comment.

-   Avoid showing anything that represents a specific person or brand other than your own. Users reading through the comments might confuse the default avatar with something that the commenter chose.

-   Don't put a picture of yourself. To change your own avatar, see [Updating Account Settings](#updating-your-account-settings)

If you prefer the original default avatar, you may download the one below and upload it in your Disqus admin again using the same instructions above.
​
​

​

### Embedding a comment on your website or blog {#embedding-a-comment-on-your-website-or-blog}

Embedded comments let you bring the best comments in your community directly into the content you publish on any website. No longer will you need to take screenshots of comments to share in a blog post. Just use the auto-generated embed code and paste it directly into any HTML web page.

#### This is what it looks like:

#### To embed a comment, you just have to:

-   Get the direct link to the comment

-   Visit 0

-   Copy the embed code into a blog post or website and publish

Embedding comments is super easy. With its flexibility, there are a number of ways you can use it. Here are some ideas we recommend:

1.  ***Promote the top commenters in your community*** - Recognition is a huge motivator in communities that can increase engagement from your commenters. Over at xoJane, they publish a “Comment of the Week” series where they recap the most interesting comments from their community. Promoting comments also lets you invite readers to join the discussion.

2.  ***Ask your readers a question*** - We encourage authors to be an active presence in the comments whether that’s replying to readers’ comments or asking them a question. A best practice to getting a new discussion going is to pose a question directly to your readers in a comment that you feature. Embed that question directly in your article so that readers are more likely to see it and share their thoughts.

3.  ***Create more compelling stories*** - When creating content, publishers care not only about it attracting an interested audience but one that leads to meaningful engagement. Engagement that sparks discussion, provide new perspectives, and connect readers closer to the story.

#### Questions?

#### Q: What happens to an embedded comment if the comment is deleted/pending?

The embed turns into a Guest author and the comment text is removed.

#### Q: Can I embed comments from a private profile?

Yes.

#### Q: How many comments does this show?

Embedding a single comment will only show that comment, unless it is a reply, then it will show the parent from to which it's replying. Like this:

With embedded comments, you can integrate the best content from your community to unlock more ways to create better content.

### Guiding your Community from Good to Great {#guiding-your-community-from-good-to-great}

Welcome to the discussion! Jump to this week’s questions.

The internet is like a giant delicious pie that’s filled with all kinds of cool junk, and you are building a community for a very specific slice of that sweet internet-pie. Pie metaphors aside, you probably are working to build a unique community filled with people who interact with each other in distinct ways. As a community leader, it is important for you to invest in your community and encourage these conversations and interactions on your site.

That’s where community guidelines come in. They’re one useful tool for strengthening your community by setting straightforward expectations and “rules of engagement” for your members. Check out these communities for some great examples:

#### The Dissolve

The Dissolve links directly to their guidelines below the comments for better visibility.

#### Politico

Providing an accessible contact link can help community moderators do a better job.

#### What community guidelines are right for my site?

If your community is still taking shape, we have some sample guidelines to get you started. But you’ll also want to think about type of community you want to build and the behavior you want to encourage. Maybe your members communicate primarily in cat-memes and puns, or perhaps you hope to generate in-depth conversations and civil debate. Think about the community that you want to build, and how your guidelines can help steer the discussions in that direction.

#### Guidelines can empower great discussions

Logic may tell you that you should keep your site as “open” as possible, so that you don’t accidentally stifle discussion. However, if you want to encourage quality discussions, providing clear expectations for your community can help focus and curate great conversations.

Guidelines can set the groundwork for what is allowed, and what isn't. By setting expectations of what is acceptable on your site, you'll allow more people to feel welcome to engage.

-   What kinds of guidelines do you set for your forum? Do you have an example?

-   What’s unique about your community and how does that affect your guidelines?

-   Describe a situation where community guidelines had an impact on someone's behavior on your site

We welcome relevant, respectful comments. Please read our Community Guidelines.

### How to Increase Reader Engagement and Retention {#how-to-increase-reader-engagement-and-retention}

*💡 Looking for more? Download our free ebook* *The Ultimate Guide to Increasing Reader Engagement*

1.  Listening to your community

2.  Turn comments into content

3.  Host events

4.  Promote comments outside your website

5.  Recognize top commenters to retain them

6.  Lower the barrier to commenting

An engaged audience is more likely to spend more time on your site—and then come back and do it again. More importantly, engaged readers help ensure higher publisher earnings for sites with ads enabled. But how do you get your readers to start commenting in the first place? And how do you keep them coming back for more? Follow these steps for increasing reader engagement.

Clear community guidelines are a useful tool for strengthening your commenting community. Create a safe and supportive space for readers to share their thoughts by setting straightforward expectations and rules of engagement. Let your readers know that, while you welcome heated debates, there are moments when moderators may need to step in. Use our sample guidelines to help you get started on your own community guidelines.

#### 2. Reciprocity is key

Once you’ve outlined a few rules and regulations for your community, you’ll need to designate a moderator. Forum moderators should not only filter any harmful or unwanted content, they should also respond to comments on your page to let readers know you’re listening. If you have other writers and content contributors, go a step further and encourage them to chime in as well. Keeping your site’s voice active in the comments section is a great way to build longer, more meaningful discussions.

#### 3. Always ask questions

Asking your audience a question is perhaps one of the oldest tricks in the book for increasing reader engagement, but remember, consistency is key. Make a point to include a question in every article, or add a call for comments to your author bio. Reminding readers that you want to hear from them is a great way to promote engagement.

#### 4. Jump to comments

Posing a question to your readers can help boost reader participation—but don’t stop there. Provide your audience with multiple opportunities to chime in by including a link that jumps to the comments section. Doing so will encourage your readers to voice their opinions. Consider adding a “jump to comments” link at the top of the article and after every question you ask your readers, or include a comment count after article titles.
​
*Tip: Getting more readers to visit your comments section could improve your ad viewability and earnings.*

#### 5. Highlight the best comments

If you have a comment that sparks a great discussion on your site (or just a comment that you think should get more attention), you can feature it at the top of your discussion thread. This will make the comment more visible to readers scrolling through your discussion section so that it won’t get overlooked. Get the how-to for featuring a comment.

### Personalize Your Site {#personalize-your-site}

Add a default site avatar for commenters who don’t have one yet. This is a nice way to make readers feel like a part of the community.

#### 2. Create a custom moderator tag

Moderators on your site have a special tag next to their comments verifying who they are. By default this reads "Mod" but you can easily customize this.

#### 3. Set the default sort order

If a commenter hasn't chosen a preferred sort order, the default you set is used instead. The default option of "Best" is recommended for most sites.

#### 4. Update the color scheme and typeface

Disqus automatically checks your site's font and background color and adapts to either a light or dark color scheme, along with a serif or sans-serif font. If these are detected incorrectly, you can override them here.

**Next up:Create and Enforce Moderation Rules →**

### Recommendations {#recommendations}

Recommendations is an engine that helps recirculate traffic to pages with Disqus installed.

To customize or disable recommended links within your site, go to your Disqus Admin, click Settings -> "Recommendations".

From this page you can customize:

-   **Recommendations Placement:** this can be show at the top or the bottom of the embed.

-   **Content Descriptions:** you can choose to show or hide brief descriptions of the content on that page.

-   **Publish Date:** you can choose to show or hide the date when the thread was created.

-   **Comment Count:** you can show or hide the comment totals of the recommended pages.

-   **Date Threshold: y**ou can set how current the threads that appear in Recommendations are, selecting between threads created in the last week, month, 6 months, or year.

-   **Placement:** you can choose whether you'd like Recommendations to appear at the Top or the Bottom of the Disqus embed.

#### Standalone Recommendations

In addition to being shown at the top or the bottom of the comments section, the Recommendations unit can be inserted elsewhere on the page (apart from the Disqus comment section), or on pages without Disqus comments present.
​
First, you'll want to open your site's code editor, and insert the following script code into the body of the page, ensuring that the EXAMPLE shortname is replaced with your site's [shortname](#what-s-a-shortname).

    <script>
    (function() { // REQUIRED CONFIGURATION VARIABLE: EDIT THE SHORTNAME BELOW
    var d = document, s = d.createElement('script'); // IMPORTANT: Replace EXAMPLE with your forum shortname!
    s.src = '0'; s.setAttribute('data-timestamp', +new Date());
    (d.head || d.body).appendChild(s);
    })();
    </script>
    <noscript>
    Please enable JavaScript to view the
    <a href="`https://disqus.com/?ref_noscript`" rel="nofollow">
    comments powered by Disqus.
    </a>
    </noscript>

You'll then want to insert the following Div into the body of your site. The div's placement will determine where we show the Recommendations unit, so you'll want to make sure that this line is placed in the location that corresponds with where you'd like the unit to appear:

    <div id="disqus_recommendations"></div>

If your site is using the DISQUS.reset method for infinite scroll, this can also be used with the Standalone Recommendations unit. To implement with this method, you'll want to use the following reset line:

    window.DISQUS_RECOMMENDATIONS.reset()

#### F.A.Q.

#### Recommendations isn't working - what can I do?

Recommendations require Disqus servers to access your pages. If Recommendations is not populating as expected, you can try adding Disqus' public IP addresses to your site's allowlist. A list of our public IPs can be found here.

#### How does it choose which links to recommend?

We will look for recent discussions within the same site and show ones with recent activity. If none were found, Recommendations won't show until there are a few recent discussions with activity. Note that Recommendations will only show discussions using the same shortname.

#### How can I change a post's title appearing in Recommendations?

Discussion titles are set automatically using the title assigned to the page where the thread was created. If the title is incorrect, it can be updated using the threads/update API call.

#### Why are there no recommendations showing?

Because article recommendations are based on comment content, not article content, a discussion must have at least a few recent comments to be shown in Recommendations. Recommendations works best for sites with some consistent commenting activity.

#### How do I block certain links?

Close comments for a discussion to prevent it from appearing in Recommendations. However, if there are recommendations showing that should not be accessible to the public, we recommend closing the comments for that discussion and registering a new shortname specifically for your future test discussions. Check out Best practices for staging, development and preview sites for more information.

#### I have ads enabled on my site, can I still use Recommendations?

Yes, if you have Ads and Recommendations enabled, we'll show both the Ad and the Recommendations unit side by side:

####  How do I manually set the images for my Recommendations?

You may want to do this if an image is not appearing automatically, or if you'd like a different image to appear in the window for a given thread. To manually set an image on the page for us to scrape, you'll want to add this code (with the appropriate thread image) to the 0 of the page:

    <meta property="og:image" content="IMAGE_URL">

### Sample Community Guidelines {#sample-community-guidelines}

Every site on Disqus over time should develop their own community guidelines that considers the community they are trying to cultivate and reflects the types of discussions they want to encourage. Here are some examples active communities have created for their members:
​
**Sites**

-   The Mary Sue: The Mary Sue’s Comment Policy

-   Destructoid: Be chill. Be critical. Be you. But don't be a dick

-   Grist: Comment Policy

-   Screenage Wasteland: Discuss. Debate. Agree and disagree. But stick to the topic

We’ve put together a set of universal guidelines that we believe every community can benefit from and moderators can adopt. Our goal is to ensure that these set of guidelines are easy to understand and widely applicable to any community on Disqus. It’s not meant to be comprehensive, rather focusing on the most commonly shared principles we believe are critical for setting a foundation in a community to enable it to thrive.

#### Keep it civil aka don’t be a jerk

We’re going to get into the thick of a lot of heated discussions and that’s okay. These discussions often entail topics that we all personally care a lot about and will passionately defend. But in order for discussions to thrive here, we need to remember to **criticize ideas, not people**.
​
So, remember to avoid:

-   name-calling

-   ad hominem attacks

-   Responding to a post’s tone instead of its actual content.

-   knee-jerk contradiction

Comments that we find to be hateful, inflammatory, or harassing may be removed. If you don’t have something nice to say about another user, don't say it. Treat others the way you’d like to be treated.

#### Always strive to add value to every interaction and discussion you participate in

There are a lot of discussions that happen every day on Disqus. Before joining in a discussion, browse through some of the most recent and active discussions happening in the community, especially if you’re new there.
​
If you are not sure your post adds to the conversation, think over what you want to say and try again later.

#### Keep it tidy

Help make moderators’ lives easier by taking a moment to ensure that what you’re about to post is in the right place. That means:

-   don’t post off-topic comments

-   don’t cross-post the same thing multiple times

-   review any specific posting guidelines for the community: some communities such as a movies community may have specific rules regarding spoilers.

#### If you see something, say something

Moderators are at the forefront of combatting spam, mediating disputes and enforcing community guidelines and, so are you.
​
If you see an issue, contact the moderators if possible or flag any comments for review. If you believe someone has violated the Basic Rules, report it to Disqus by flagging the user's profile.

#### Related: Guide to building community guidelines

### Site identity information {#site-identity-information}

In your site settings there are a few options available to provide your site identity. These include:

**Website name**
Simply the name of your site. This will show up in Disqus feeds, Email notifications, and your Community tab.

**Website URL**
The starting URL of your site. Note that this is for display purposes and your site won't be bound to only this domain.

**Category**
Your site's general category. Choose "Other" if none of the choices are relevant.

**Description**
Usually one or two engaging sentences about your site to draw people in.

**Language**
The default language to be displayed in the Disqus embed. This can be overridden with the javascript language override.

**Site Image (Favicon)**
Images are pulled automatically, more information below.

The default language to be displayed in the Disqus embed. This can be overridden with the javascript language override.

Note: your forum shortname cannot be changed. The shortname is not user-facing and does not need to be updated if your Website name changes

#### Other identity information

In addition to the above, we also pull in images and content from the articles. It's best practice to include proper Open Graph tags on your pages in order to represent your content the way you'd like to.

Favicons are automatically pulled in from the URL provided in the site identity settings, so make sure an accessible address is provided in order for us to receive it.

**Is Disqus showing an old favicon for your website?**

Click the blue chat icon on this this page in the bottom right corner and please provide us the following so we can reset the image:

-   Website URL

-   Disqus shortname

#### Where is this information shown?

This information will be aggregated throughout several places:

-   The Site link at the top of comments on the User Profile

-   My Disqus feeds

-   Email notifications and digests

#### Also see

-   What's a shortname?

### Site Profiles {#site-profiles}

In addition to the comment threads on their pages, each site using Disqus will also have a Site Profile. The site profile is where users may:

-   Follow the site

-   View active discussions created by that Site

-   View the top commenters of the community

#### **Viewing a Site's Profile**

When on a given site, you can visit their site profile by clicking the Community link in the dropdown menu.

If you've commented on a site, a link to their Site Profile will also appear at the top of any comment on your user profile.
​

#### **Viewing a Site's Active Discussions**

Once on the Site Profile, you may view a list of active Discussions from that forum. This section will show all threads that have had at least one comment posted to them within a week of thread creation.

#### **Following and interacting with a site**

Follow a community or a site by going to their page and clicking on the follow button in the top banner. Once you start following the site, new content will start flowing into your home feed. You may also view the site's Top Commenters by clicking the Top Commenters option below the Follow button.
​

## Disqus Polls {#cat-disqus-polls}

### Disqus Polls {#disqus-polls}

Welcome to the newest Disqus product, **Disqus Polls**. Disqus Polls will allow you to create fully customizable polls to be installed anywhere on your website, complete with reporting and analytics so that you can monitor the additional engagement from your site visitors.

*Please note this is a separate product from Disqus Comments, and will require a separate subscription, or specific approval to join the Ads-Supported version of Disqus Polls. For more information on the plans that provide access to Disqus Polls, please check our “Gaining Access to Polls” section* here*.*

Once you’ve subscribed to the correct Polls plan, new polls may be created from your Polls Editor Page. You can also access this page by clicking the “Polls” option in the header of your admin and then navigating to the “Editor” option appearing in the left sidebar.
​

The first step of creating your new poll is adding the poll name and date range. The date range may be entered manually, or selected from a calendar view by clicking the calendar icon in the right side of the field. The poll name will not appear externally, and all of these header elements may be edited later if needed.

After adding the name and date range, you'll want to select any of the following options if they are applicable to this poll:
***Require login to submit response***

This will prompt login with a Disqus account before the user can submit the poll.
​
***Allow user to take poll again***

When checked, you’ll be given options to allow a user to take the poll again after 1, 7, or 30 days. This option will work even if users are not logged in.

***Exclude from Universal Tag***

The Universal Tag is a single set of code to be installed on your site that will allow the cycling through of all active polls that do not have “Exclude from Universal Tag” checked. Checking this option will ensure that this poll does not appear anywhere that the Universal Tag is installed.

***Add Call-To-Action (CTA) to results screen***

This allows you to insert a link with custom link text on the Poll Results page, to capture and direct your audience's traffic.

#### Adding a Call-to-Action Link

When the "Add Call-to-action (CTA) to results screen" option is checked for a poll, you may also add a call-to-action link to your poll unit. You can customize the link text and add a link value of your choosing.
​

The link button will appear as an overlay on the results screen after the user has completed all questions of this poll. The CTA button will inherit the same link color value from your site that is used for the Submit button and a selected answer.
​

####   Adding Questions and Answers

**When creating new polls, we recommend always *saving* your poll entries. Once they are Published, the user-facing elements cannot be edited in any way, and the poll will need to be recreated entirely if you wish to change any of the question or answer elements.**
​

For each individual poll you can include a maximum of 3 questions, and each question can have up to 8 possible answers. If multiple questions are included in a single poll, the respondent will only see the poll results once they have selected answers for all questions and have clicked Submit.
​
For each question you can also choose how the answers will display and behave. The display and selection options will appear in the dropdown menu to the right of each question.

-   ***Single Choice*** - Answers will appear in the same order every time, and only one answer may be chosen

-   ***Multi-select*** - Answers will appear in the same order every time, but the user will be able to select multiple answers at once for this question

-   ***Single Choice Random*** - Only a single answer may be selected, but the answer options will appear in a different order each time the poll appears

-   ***Single Choice Random + Anchor*** - Only a single answer may be selected, and the answers will appear in a different order each time, aside from the last answer option, which will always appear at the bottom (and can be used for an “All of the above” answer or something similar)

####  Publishing Your Poll

Again, we recommend ***saving and reviewing polls before publishing them***. Once they are published, all question and answer text and options will become fixed and uneditable.
​
You may also click Preview at any point to see how the poll will appear on a separate page. If you are on an Ads-Supported plan, an ad creative will appear next to the poll once it is set live, but will not be displayed on the Preview page.

Once you’ve double-checked all aspects of your poll and are ready to set it live, click the Publish button. This will lock the Poll and set it live, and the Publish button will change into a “Tags” button which can be used to copy the code block for this specific Poll for installation on your site.

####  Installing Polls to your Site

Once you've clicked publish and copied the tags for your poll, this will add the entire code block needed for poll installation to your clipboard. You can then go into your site code, and paste this code block within the 0 where you'd like the poll to appear. The Disqus Polls unit is responsive, and will give itself the best layout based on the space provided that it is installed within.
​
When installing the poll unit, place the code block in a location viewable on both desktop and mobile to generate the highest engagement. For example, when placing the poll in a side-rail, or multi-column layout on desktop, ensure that portion of the site responsively comes into view on mobile as well.
​
We recommend not editing this code block, as it could affect the poll's functionality.
Tags for any poll can also be copied from your Polls Editor page by clicking the computer code icon under the Actions section:

If you are looking to have multiple Polls appearing simultaneously on a single page, we recommend using only one instance of the script portion of the Polls code and installing the script at the end of the body of your site, preferably just before the </body> line. With the single script in that fixed position, you'll still be able to dictate the placement of each poll on the page by moving the div code for each poll.
​
The div code for a single poll will look something like this:
0

#### Using the Universal Tags

In addition to the poll-specific tags provided on Publish and on the Polls Editor page, we also offer a Universal Tags option. This allows you to install a single block of code to your site, which will then cycle through all active polls you have set up. All polls will be shown by installing this code, aside from polls which have expired or polls which have the *"Exclude from Universal Tags"* option selected.

To implement the Universal Tags, simply click the Copy Universal Tags button at the top of the Polls Editor page, and paste this code into the 0 of your site where you'd like the Polls unit to appear.

#### Polls Appearance

In a similar fashion to Disqus Comments, site styling will be inherited automatically. This includes the background color, and the link color used for selection and the Submit button. For Ads-Supported Polls sites, the Poll unit will automatically determine whether to display the ad creative next to or below the poll, depending on the width of the container it is installed within. The module will adapt to vertical or horizontal orientation automatically, and will format itself depending on the sie of the container it is installed within. If it is installed in a container that is too short to show all of the Answer options, a scrollbar will appear automatically.
​
**Editing the Link Color**

We'll automatically pull the link color from your site's theme, and use this color for the poll Answers options and the Submit button.
​
To manually change what color is used for this, you may add a version of the following to your CSS sheet:
​
0

0

0

**Language Overrides**
Using the 0 JavaScript variable, you can dynamically load the Disqus poll unit in different languages on a per-page basis. For example, to load the poll in French:
​

    var disqus_polls_config = function () { this.language = "fr"; };

For a full list of language codes, see Languages on Transifex. Note however that Disqus does not support all languages offered in Transifex. For a full list of supported languages, see the Disqus project on Transifex.

#### Polls Analytics

By default, the Polls Analytics page will show an aggregate of all Polls activity on your site. To drill deeper into this data, you may select a specific poll in the dropdown at the top of the page. The graph will then chart Respondents, Impressions, and Response Rate for that Poll, so that you can see changes in engagement over time. You may group data by day or month, and then hover over any area of the graph to see the numbers for the selected period. A custom date range can be selected by clicking into the date range selector in the top right corner, and the results in your graph may be exported to CSV by clicking the Export Data button.
​

In addition to the engagement chart, a breakdown of the responses to each question will appear in pie charts below the graph, which will allow you to see total votes for each answer.

#### Gaining Access to Polls

Polls requires a separate subscription from Disqus Comments. To purchase a Polls plan as a new customer, you may select a subscription plan based on your site's monthly pageviews at our pricing page here.
​
If you're already using Disqus Comments, you may add a Polls subscription to your plan from your Subscription and Billing page.
​

Available plans are listed below.
​

#### Polls Pro ***(ads-free)***

-   0-250K monthly pageviews - **\$9/mo** (when billed annually, otherwise, \$10/mo)

-   250k-1.5M monthly pageviews - **\$22/mo** (when billed annually, otherwise, \$25/mo)

-   1.5M-3M monthly pageviews - **\$45/mo** (when billed annually, otherwise, \$50/mo)

-   3M-10M+ monthly pageviews - **\$90/mo** (when billed annually, otherwise, \$100/mo)

#### Polls Business ***(ads-free)***

The Polls Business plan offers SSO integrations, dedicated Account Management, Custom Integrations, and allows for the removal of Disqus branding. Pricing for this plan will vary by use case, so get in touch with us here for more information.

#### Polls Ads-Supported

This plan allows you access to the Polls feature free of charge, but will require an ads unit to be run with your poll. All sites wishing to use this plan must be reviewed and approved by our team. To submit a request for approval, please fill out our form here.
​

#### FAQ

**If I set a certain date for my poll to go live, when exactly on that date will it appear?**

Polls will start at 12am on the start date and end at 11:59pm on the end date (local time to the user who created or last edited the poll).

### Disqus Polls - Pricing and Plans {#disqus-polls-pricing-and-plans}

*Disqus Polls does not include access to Disqus Comments, which require a separate subscription. Our Comments Pricing and Plans may be found* [here](#comments-pricing-and-plans)*.*

Each Polls Pro plan has limits of eligibility based on site traffic. To purchase a Polls plan as a new customer, you may select a subscription plan based on your site's monthly pageviews at our pricing page here.
​
If you're already using Disqus Comments, you may add a Polls subscription to your plan from your Organization's Subscription and Billing page.
​

#### Polls Pro ***(ads-free)***

-   0-250K monthly pageviews - **\$9/mo** (when billed annually, otherwise, \$10/mo)

-   250k-1.5M monthly pageviews - **\$22/mo** (when billed annually, otherwise, \$25/mo)

-   1.5M-3M monthly pageviews - **\$45/mo** (when billed annually, otherwise, \$50/mo)

-   3M-10M+ monthly pageviews - **\$90/mo** (when billed annually, otherwise, \$100/mo)

#### Polls Business ***(ads-free)***

The Polls Business plan offers SSO integrations, dedicated Account Management, Custom Integrations, and allows for the removal of Disqus branding. Pricing for this plan will vary by use case, so get in touch with us here for more information.

#### Polls Ads-Supported

This plan allows you access to the Polls feature free of charge, but will require an ads unit to be run with your poll. All sites wishing to use this plan must be reviewed and approved by our team. To submit a request for approval, please fill out our form here.

### Polls {#polls}

*This article pertains to our Polls offering for select Ads-Supported publishers. To see if you qualify for this product, please write to us at publisher-success@disqus.com. For Custom Polling options with the ability to create your own questions and answers, please see our Disqus Polls documentation [here](#disqus-polls).*

For qualifying Ads-Supported publishers, Disqus is now offering an additional tool for audience engagement - Auto-Generated Polls.

When enabled, this will create Polls that appear above your Comments embed, with auto-generated questions and answers based on your forum category. There is no need to input any content, or do any poll management. You'll also have the option to enable Contextual Polls, which will generate Polls directly based on your site's article content, rather than your forum category.
​

Auto-Generated Polls are only available to qualifying Ads-Supported publishers, and will appear with an adjacent ad placement above the Disqus Comments. To see if you qualify for this product, please write to us at publisher-success@disqus.com.
​
Once you've been approved, you will be able to enable the Polls from the Polls Settings Page in your admin.

You'll also have the option here to enable **Contextual Polls**. This option prioritizes creating Polls based on the content of your site's articles, and falls back to category-based when necessary. Contextual Polls will give you the Polls that are most relevant to your site content and community. The variety and volume of available Contextual Polls is based on how frequently new article content is posted.
​
If you choose to run auto-generated Polls on the default category setting, you'll want to ensure that your forum category is up to date, as this is what those Polls will base their content on. You may update your forum category from the General Settings page in your admin.

Once Polls have been enabled by checking one or both checkboxes on this page, new auto-generated Polls will begin appearing above your Comments embeds, typically within 30 minutes. To disable Polls, simply ensure that both boxes on the Polls Settings Page are unchecked.
​

####  Polls Placement and Styling

When enabled, Polls will appear with an ad unit above the Comments embed. The Polls unit is responsive to the width of the Disqus embed, and will widen or narrow as necessary to remain centered above the comments with the size of the joining ad unit.
​

If Reactions are also enabled, Polls will appear above the Reactions options. Polls can only appear above the Comments in this location, so if Recommendations is also enabled in the top position along with Polls, then Recommendations will automatically be moved below the Comments embed. It is not currently possible to reverse these placements, or show auto-generated Polls below the comments or in the middle of a comments thread.

#### Editing the Poll Color

We'll automatically pull the link color from your site's theme, and use this color for the poll Answers options and the Previous and Next buttons.
​
To manually change what color is used for this, you may add a version of the following to your CSS sheet:
​
0

0

0

#### Taking Polls

When a user is taking Polls, all they need to do is click their response to the question, and the unit will submit this as their answer automatically. They will then be shown the results of that Poll for 5 seconds before the next Poll loads. They may click the Next button to proceed to next poll immediately, or click the Previous button to view the results of all the Polls they've taken in this unit.

#### FAQ

##### **Do I have to pay for Polls?**

No, Polls is a free-to-use product for sites that are running ads and are qualified for this unit. To see if you are qualified to run this Polls unit, please email us at publisher-success@disqus.com.
​
The paid Polls plan on the Subscription and Billing page only applies to the ability to run [Disqus Polls](#disqus-polls), our Custom Polling unit that allows the publisher to set the questions and answers manually.

##### **Why are my Recommendations appearing below when Polls are also enabled?**

Polls is only configured to appear above the Comment embed.
When Recommendations is enabled in the top position **and** Polls is enabled, Polls will appear above the comments, and Recommendations will automatically adjust to appear below the comments. It is not possible to customize these placements at this time.

#####  **Does Polls use AI? Will my content be used to train an AI model?**

Polls are generated using Artificial Intelligence (AI). When contextual polls are enabled, Disqus uses the page content solely to create a poll relevant to that content. This process is handled by a **private AI model instance** dedicated to poll generation. Importantly, there is:

**No External Sharing:** Your data is never sent to any public or third-party AI service.

**No Model Training:** The content used for poll generation is not retained or used to train AI models.

##### Polls doesn't seem to be working on my site - what can I do?

Polls require Disqus servers to access your pages. If Polls is not populating as expected, you can try adding Disqus' public IP addresses to your site's allowlist. A list of our public IPs can be found here.

## Disqus Pro {#cat-disqus-pro}

### Badges {#badges}

Badges are a great way to distinguish contributing community members, and can be applied based on a number of different criteria. They will appear next to a user's name on the awarding site and will be visible universally on the user's profile, incentivizing engagement from the community.

Sites will first set up their badges from the badges settings tab in the admin panel, determining what kinds of recognition are most appropriate for their community. There are automated and manual badges, and each badge is designed to be set up with an image and a custom badge name. Up to 8 badges may be added per site.

Badges can be set based on the following criteria:

-   **Number of upvotes on a comment** - a user will get this badge automatically when a comment of theirs receives the target number of upvotes

-   **Number of comments** - a user will get this badge automatically once they have posted the target number of comments to the site

-   **Number of featured comments** - a user will get this badge automatically once they have posted the target number of comments to be featured by a moderator of the site

-   **Manual** - a user will receive this badge type when a moderator adds the badge to their account

#### Adding Images

Once the names, criteria, and target numbers have been set, you'll want to add images to each badge. Badge images must be in the .jpg or .png formats, must be under 25kb in size, and should be roughly 64 x 64 pixels.
​
As long as the above criteria are met, any image may be uploaded and used as the badge image. If you'd prefer to use pre-created images for your badges, we have some available to download here.
​

#### Applying Badges to Users

Automated badges will be applied automatically to user accounts, though the counter will start once the badge is created. For example, if you create a badge to be set when a user has posted 100 comments, it will only be applied to accounts who have posted 100 comments **after** the badge was created. It will not be retroactively applied to users who have posted 100 comments to the site prior to badge creation. If a badge is deleted and re-added it will start the counter anew, but if the badge is simply edited it will remain applied to all users who have received it, regardless of the new criteria.

To apply a Manual badge to an account, moderators will want to locate a comment by the user in the embed, and click the dropdown menu on the right side of the comment, selecting "Manage Badges"

This will open the window within the comment embed to award an existing badge to the user, create a new badge, or remove a badge from the user's account.

Badges may also be applied through the moderation panel. Simply click on a comment, and then click the "Manage user badges" option that appears in the right sidebar

Once a badge has been awarded to a user, they'll be notified in their Disqus Notifications:

####  Where will badges appear?

Once a badge has been applied to a user's account, it will appear next to their name for all comments left on that site. Up to 3 badges will appear by default, and users may click to view additional badges in the comment embed.

Additionally, all badges awarded to the user (from all sites) may be viewed in the Badges section of the user's profile:

#### User Controls

Once they have received a badge, commenters can determine which badges they want to show next to their comments and on their profiles. They can click the slider below each badge to retain the badge, but hide it from their profiles and comments.
​
Hovering over each badge will produce a trash can icon. Clicking this will allow the commenter to remove this badge from their profile entirely. They may still re-earn the badge later, either by it being manually awarded again, or again passing the milestone needed to obtain that badge.
​

#### Email Notifications

Users can opt into Badges notifications from the Email Notifications section of their Account Settings page. When checked, users will receive emails whenever they receive a new badge.

#### If Badges Are Not Appearing For Your Site

If Badges are not appearing as an option for the sites in your organization, this likely means that at least one of the sites in your organization is not qualified to run Disqus Ads. Subscribing to a paid plan via your Subscription & Billing page will make the Badges option available for all sites in the Organization.

### Disqus Appearance Customizations {#disqus-appearance-customizations}

The following customizations require a Pro or above subscription to enable:

-   Disable Disqus Branding

-   Disable Social Share Icons

-   Disable or customize Voting

-   Use Custom Fonts

-   Customize discussion prompts

-   Customize default thread length

Color scheme and link color edits may be applied by all sites.

####  Disqus Branding

If you have a Pro subscription plan, you can remove the Disqus logo from your footer. You can update this by going to the General Settings page in your Disqus Admin.

Unchecking the Disqus Branding option will remove the Disqus logo from the right side of the footer:

For those with a Business subscription, additional branding removal can be enacted when Single Sign-On is employed as the only login method. With SSO-only in place, you can additionally remove:

-   The remainder of the Footer options

-   The thread "Favorite" heart icon

-   Thread-level social-sharing link

-   User profile links that would point to Disqus.com

For more information on branding removal with the Business plan, please contact contact your Disqus Account Manager or request information from our team here***.***

​

#### Social Sharing Icons

This will allow you to disable the Facebook and Twitter social sharing buttons at both the comment and thread level. When on the Pro plan, you may disable social sharing here. Please note that even with Social Sharing icons disabled, a share option will still appear at the comment level, but will only contain a direct link to the comment, with the previous Facebook and Twitter comment share options removed.
​
Before:

After:

#### Voting Customizations

Voting needs may vary from site to site. When on the Pro plan, you can now set the voting functionality that will be best for your site. This includes:

-   Upvotes and Downvotes enabled (default)

-   Only Upvotes enabled

-   Upvotes and Downvotes disabled

Please note that the ability to hide downvote details (such as count and voting users) is still available to sites on all plans.

#### Custom Fonts

For sites with a Business subscription, a variety of fonts will be available for use. The desired font may be selected from the Typeface section of the General Settings page.
​
Business Publishers may also request for new fonts to be added to the Disqus system. To be assessed for addition, the font must meet the following requirements:

-   Be available through Google Fonts

-   Have the following styles available: 400 weight regular, 400 regular italic, 700 bold, 700 bold italic, and either 500 medium or 600 semi-bold

#### Customize Discussion Prompts

This will allow you to customize the text that appears in your postbox where your users type their comment. Currently, this text will read "Start the discussion" for a thread with 0 comments, and "Join the discussion" for an active thread. With this feature, you can use any text you'd like with a 45 character limit. This change will be applied to all threads on your site.

#### Customize Embed Length on Load

This feature allows you to customize how many comments are shown in the embed when Disqus first loads on the page. This can be used to have Disqus load a smaller window when it first appears on the page.
​
By default, 50 comments will load initially on desktop, and 20 comments will load initially on mobile. This feature can be used to show any number of comments below 50 on initial embed load.
​
Please note that even with this customization enabled, clicking "Load More Comments" will load an additional 50 comments on desktop, and an additional 20 comments on mobile.

The following settings and tweaks can be applied to all sites on Disqus.
​
Although it's not currently possible to apply custom CSS to the Disqus iFrame, the appearance can still be tweaked in a few different ways.

#### Light vs. dark color scheme

A light or dark color scheme is automatically selected based on your site's stylesheets.
​

#### How is the color scheme determined?

-   The light scheme is loaded when the text color Disqus inherits from your site has >= 50% gray contrast: between 0 and 1

-   The dark scheme is loaded in all other instances.

#### Overriding the color scheme

The color scheme can be overridden in two ways:

*In the Disqus admin*

1.  At the Disqus Admin > Setup > Appearance page, locate the Color Scheme option.

2.  Choose the appropriate color scheme, or allow Disqus to choose for you by selecting "Auto".

*In your site's stylesheets directly*

1\. Locate 0

2\. Insert a 0 tag into disqus_thread. This requires editing HTML in the web inspector.
3. Inspect the 0 tag > expand the Computed Style dropdown > observe the 'color' parameter. This is the color that is being inherited.

4\. Expand the dropdown to the left of the 'color' parameter to expose which specific stylesheet rule is setting this color. Change the color being passed via this rule based on the "How is the color scheme determined?" section above.
​

#### Link color

Disqus inherits your site's link colors, and will use this color for all links, the Load More Comments button, and the upvote/downvote buttons when clicked. In order for this to work you'll need to make sure the relevant CSS rules are inheritable by Disqus.

To update the color used by Disqus for these elements, you may insert the following code in your CSS sheet, and then update the color line accordingly:
​
0

0

0
​

#### Width, margin and padding

Disqus is set to fill 100% of the width of its parent HTML element and has no margin or padding set on its 0. This means Disqus often looks best when its parent container gives it some margin or padding. Additionally, the width of the 1 ID can be adjusted using CSS.
​

#### Elements that can't be edited or removed

-   Font size

-   "# Comments" text above the embed (For home page comment counts, see Adding comment count links to your home page)

### Disqus Pro Analytics {#disqus-pro-analytics}

The analytics dashboard features top-line metrics and content analysis that help publishers better understand audience engagement occurring directly on their site. This analytics suite is designed to provide actionable insights for the unique engagement you’re capturing through Disqus.
​
To visit the advanced analytics dashboard, go to the Disqus admin > Analytics.

***\*\*Access to Pro Analytics is currently available with a**Pro subscription**. If you would like to subscribe to Pro, you can do so in your**Subscription & Billing**.***

For information about our basic community metrics page, which will remain available to all sites, see Commenting and Ads Metrics.

#### Overview Analytics

Track these top-line metrics over time to easily identify how your site is performing with your audience. Month-over-month change indicators will display in green and red so that you can quickly gauge the health of your engagement metrics, especially as your community grows over time.

-   ***Article reads***: The number of times people view an article where Disqus is installed.

-   ***Comment reads***: The number of times people view the comment section on an article.

-   ***Total engagements***: The number of times people have commented or voted on an articleTraffic

-   ***Overview***: A multi-axis graph of Article Reads, Embed Reads, Engagements over time so that you can easily track the positive relationship between Disqus engagement and your readership.

Use the Traffic Overview graph to dig into top-line metrics for specific days or to zoom out for spotting larger trends. Click on specific metrics to toggle their visibility in the graph to adjust your focus when tracking trends over time.

#### Content Analytics

Content Analysis provides a summary of your most recent site content, ranked by total engagement. The insights generated on this page are designed to inform your unique content strategy as a publisher so that you can optimize for attracting and retaining loyal readers.

-   ***Total Engagements***: Total number of comments and votes for a given article.

-   ***Comments***: Total comments for a given article.

-   ***Commenters***: Total unique commenters for a given article.

Click on any of the engagement column headers to sort the data, or use the date picker to narrow your view to a specific date range.

#### Audience Analytics

The Audience Analytics provide insights into your readers’ engagement behavior and informs the actions you can take to further optimize your community.

#### Overview:

Use the Overview to understand how your community is growing and to measure the overall impact of your current engagement strategy.

-   **Comment readers:** total users who read the comments on your site

-   **Engaged readers:** total users who have either commented or voted a comment

-   **Subscribers:** total users who have opted in to your email list via the Email Subscriptions feature.

#### Snapshot of Engaged Readers:

Drill down further for a breakdown of total New, Returning, and Recovered users to see how each segment is represented in your community and their respective rate of growth.

-   **New:** Engaged for the first time on your site in the past 30 days

-   **Returning:** Engaged on your site in the past 30 days as well as the prior 30 days

-   **Recovered:** Engaged on your site in the past 30 days but not in the prior 30 days

#### Community Members:

Use the Community Members table to view a profile of each user, including their total engagements (comments + votes), the date of their first and most recent engagements, and their engagement status.
​
​

To sort any of these columns using a column header to identify interesting segments of your community, for example "recently acquired users" or "most active commenters".

Click **Export CSV** to receive a CSV email attachment. Downloading the data from Disqus can be useful for performing a deeper analysis, especially if you want to combine Disqus user attributes with other audeince data that you have.

#### How far back can I look up metrics in Audience Analytics?

Data is available as far as a year ago so you can find metrics up to January 2017.

#### I did not receive the Exported CSV email, what should I do?

While you are logged in, go to your account settings to make sure your email address is correct. Note that the time it takes to receive the email may vary based on the size of your audience. Contact Support if you still have trouble receiving the email.

### Email Subscriptions {#email-subscriptions}

***\*\*Access to Email Subscriptions is currently available with a Pro or Business subscription. If you would like to subscribe to Pro or Business, you can do so in your**Subscription & Billing**.***

The email subscription prompt appears directly within the comment embed and encourages engaged users to subscribe to the publisher's email list. This feature is designed to help you grow your subscription list and own the relationships with your readers so that you can share content more effectively.

#### Enable the Prompt

Visit the Admin > Settings > Email Subscriptions page and click "Enable Email Subscription Prompt" to enable the feature with the default copy that is provided.
​

To ensure the best experience for your readers, use form fields to customize Title, Description, and Confirmation copy. We recommend that you are transparent in this copy about the types and frequency of emails that readers will receive if they opt-in. Use the toggle view logged-in and logged-out user states and click "Subscribe" to preview the confirmation message.
​

Below is an example of how the Email Subscription form will look. The colors for the Subscribe button and "hide this message" text will be automatically pulled from your site's CSS to match your theme.

#### Export New Subscribers

At Admin > Settings > Email Subscriptions page to export all of your subscribed users as a pre-formatted CSV. This format can be imported into other popular email marketing providers such as Mailchimp and Constant Contact.
​

#### Reader Experience

For commenters, the email subscription prompt will appear below the comment box and enables readers to seamlessly opt-in to emails from publishers. For readers that already have Disqus accounts or are logged-in using one of our social sign in options, subscribing is an easy, one-click experience.

Logged out users will see a form field where they can enter an email address. To verify these email addresses and avoid spam and other fraudulent entries, Disqus will send a confirmation email to users who subscribe.

#### Disable the Prompt

Visit the Admin > Settings > Email Subscriptions page and click "Turn email prompt off".
​

#### How do I export my list of email subscribers?

In the Email Subscriptions settings page, you can either export a CSV of all subscribers who have ever opted-in or the newest subscribers since your last export.

#### Does this support an email marketing service like Mailchimp or Constant Contact for automatically importing new subscribers?

Using the "**New Email Subscriber**" trigger of the Disqus integration on Zapier, you can automatically migrate new email subscribers to nearly 40 popular email marketing services including Mailchimp, Constant Contact, Campaign Monitor, and Aweber. To get started, use any of the following Zap templates:

Check out Zapier's guide to learn how to create a Zap.

If your email service is not currently supported, Zapier allows you to set up webhooks to send information to an API.

#### What information about subscribers are available in addition to their email address?

Display name, username, user ID, IP address, signup URL, signup date, longitude/latitude

#### How do you verify email addresses?

If user already has Disqus account, their account is already verified. If the user does not have a Disqus account, the user can enter their email address which we then send a confirmation email to in order to verify.

#### What customization options are available?

You can customize the copy for the title, short description, and confirmation message shown in the email subscription prompt.

#### How do my readers unsubscribe?

Publishers are expected to manage the ability for subscribers to opt-out or unsubscribe from future email communications. Publishers should provide readers with relevant information about the frequency and type of emails that they will receive if they opt-in.

#### What are the guidelines for using this feature?

When using the Email Subscription feature, we expect publishers to abide by all relevant email and spam laws. This includes, but is not limited to, giving users who subscribe to email the ability to opt-out and honoring opt-outs appropriately. As previously stated, publishers should provide readers with relevant information about the frequency and type of emails that they will receive if they opt-in. In cases of abuse, we reserve the right to revoke access of this feature.

### Shadow banning {#shadow-banning}

Shadow banning is way of discreetly banning users, without their knowledge, in order to avoid instances of troublesome users coming back with new accounts. Shadow-banned users will be able to continue posting normally; however, their posts will not be visible to any other readers.

***\*\*Note: Shadow banning only affects a user's future comments, and does not work retroactively to hide their previous posts.***

How to shadow ban a user from your moderation panel:

1.  Go to your moderation panel and click on a comment from a user you wish to shadow ban.

2.  Select **Ban User** from the side bar on the right.

3.  A window will open that will allow you to select Shadow Ban, if you are subscribed to ***Pro***

When a shadow banned user posts new comments, their comments will appear in the Moderation Panel with a "Shadow Banned" tag.
​
​

You can find the list of recently shadow banned users in the Banned Users settings page:
​

### Timeouts {#timeouts}

***\*\*Access to Timeouts is currently available with a Pro subscription. If you would like to subscribe to Pro, you can do so in your**Subscription & Billing**.*Timeouts** is a method for temporarily restricting commenting privileges to a disruptive user for a length of time. Using this tool, you'll be able to allow a heated exchange to cool off for a predetermined amount of time, at your discretion.

#### How to give a timeout from your moderation panel:

1.  Go to your moderation panel and click on a comment from a user you wish to shadow ban.

2.  Select **Ban User** from the side bar on the right.

3.  A window will open that will allow you to select ***Timeout***, if you are subscribed to ***Pro***

4.  Here you will be able to choose how long you wish to prevent the user from commenting.

This message is displayed to a user who has been given a timeout. Timeouts are temporary, and last for any amount of time specified by the moderator.

#### Important note for moderators using this feature:

-   "Reason for banning" is an internal note for moderators, and is not displayed to the user who has received the timeout.

## Import, Export, and Syncing {#cat-import-export-and-syncing}

### Can I import comments from Facebook Comments? {#can-i-import-comments-from-facebook-comments}

This process requires re-formatting of JSON and XML data and is recommended for developers only.

Facebook Comments does not currently offer a feature to export comments all at once, however comments can be programmatically accessed from Facebook one page at a time and then re-formatted for import into Disqus.

Per Facebook's Can I get comments for a URL via an API? documentation for their Comments plugin:

The comments for every URL can be accessed via the Graph API. Simply make an HTTP GET request to:

    0}
    &id=<YOUR_URL>
    &access_token=<YOUR_TOKEN>

Note that, as this is publicly-visible data, sensitive personal information like email and IP addresses are not included. Hence comments will be imported as guest comments into Disqus and cannot be claimed by registered Disqus accounts.

#### Re-format for import into Disqus

Using our custom XML import format. This should all end up in one .xml file.

#### Import into Disqus

At the Disqus admin > Tools > Import > Generic (WXR) page.

### Domain Migration Tool {#domain-migration-tool}

The migration tool is meant for simple domain changes, such as from 0 to 1.

Note: The Migration Tool is only designed to change the Domain (what appears between http:// and .com in a URL). To update slugs ( /blog/09/2014/exampletitle.html), you’ll want to use the URL Mapper tool instead.

Once the tool starts, it will automatically detect what domain your comments are linked to. If it detects the wrong one, you can change it to the domain you're trying to migrate from.

1\. After selecting the Domain Migration Tool as your migration option, verify that the base domain which you'd like to migrate **from** is correct.

If the domain specified isn't that which you'd like to migrate **from**, enter the correct domain manually by clicking "Manually override this." *Note: If your site contains a www. in the URL, include it without the 0, e.g.* 1

2\. Enter in the new domain that you'll be migrating your threads **to**.

3\. Verify that the domains you're migrating **from** and **to** are correct and then confirm the migration.

*Migrations can take up to 24 hours to complete, so it is best to check back every few hours.*
​

#### Running the URL Mapper alongside the Domain Migration or Redirect Crawler

If you've run the Domain Migration Tool or Redirect Crawler and your threads have not migrated after 24 hours, it is possible that the threads are not set up to properly migrate with that tool. Try using the URL Mapper to input the exact thread URLs that must be changed. You can run the URL Mapper concurrently with the Domain Migration Tool or Redirect Crawler, as they will not conflict with one another.

-   URL Mapper - if the slug of your URLs have changed

-   Redirect Crawler - if you have setup 301 redirects

### How to download, edit, and upload a URL Map CSV {#how-to-download-edit-and-upload-a-url-map-csv}

If you're using the URL Mapper Tool, you'll need to download the CSV (Comma-Separated Value) file from Disqus, edit it on your computer, and upload it back on Disqus. Here's how:

Go your forum's admin panel > Discussions > Tools and click the Start URL Mapper button. **Download** the CSV file containing a list of thread URLs which currently belong to your forum by clicking the download link.

#### 2. Edit the file

The CSV will be downloaded as a compressed gzip (extension .gz) which you will need to first uncompress. If you're not sure how to uncompress the file, see this guide for both Mac and Windows.

Open the CSV file using a spreadsheet or text editor. You can use the following software to edit a CSV file:

-   Microsoft Excel

-   Open Office

-   Google Docs Spreadsheet

See the URL Mapper Tool documentation for detailed instructions on how to properly add URLs to your CSV file.

#### 3. Upload the edited file to Disqus

**Upload** the edited CSV file at your forum's admin panel > Discussions > Tools > Start URL Mapper > Upload a URL mapping.

-   URL Mapper Tool

### Importing & Exporting {#importing-exporting}

Disqus supports several XML formats natively for comments. If your comments are not in any of these formats, it'll need to be adapted to our Custom XML import schema.

#### Considerations

-   Make sure the files you're importing are valid XML. You can use the W3C xml validating tool to check.

-   Compressed files (e.g. .zip, .gzip) can't be read by default, so make sure you've decompressed these before uploading them.

-   The importer will filter out duplicate comments *unless* you've changed some of the comment data.

-   Email addresses are unique identifiers in Disqus, so make sure each unique user has their own email address before importing. Otherwise all comments will appear from the same user.

-   Imported comments **can't** be permanently deleted. Consider following our guidelines for development sites to make sure the data you're importing is correct.

#### Importing comments into Disqus

#### WordPress

#### JS-Kit

#### Custom XML imports

#### Exporting from Disqus

Disqus provides an export of all comments on your site in a g-zipped file. This is found in your Moderation panel at Disqus Admin > Setup > Export. The export will be sent into a queue and then emailed to the address associated with your account once it's ready.

Note that this file can't be re-imported into Disqus as-is.

Please note: exports may not be available for all sites, particularly those of a large size. If you've requested an export file more than twice and still have not received a download link from us, it's likely that an export for your site is currently unavailable.

#### Export format (XML)

#### Other methods

These are other ways you may be able to import/export/sync comments, but are **not supported** by us.

#### API data synchronization

#### Importing from Facebook

#### Importing from Typepad

#### Troubleshooting

#### How long will my import take to finish?

#### Troubleshooting guide

#### Syncing with Wordpress

#### Moving from Blogger to Wordpress

#### Migrating Disqus Threads

#### How can I update discussion urls?

### Importing comments from JS-Kit {#importing-comments-from-js-kit}

Users using Echo comments on WordPress or Blogger should be able to sync their comments back to either system. Afterwards, they can follow the corresponding import directions found on Disqus.

Note that this will require you to have an existing shortname registered. Additional information on registering a shortname with Disqus can be found in our Quick Start Guide.

#### Exporting your JS-Kit or Haloscan Comments

Use the following steps to export your comments from JS-Kit:

1.  Log in at 0.

2.  Visit your Moderation page.

3.  Click the Sites I Manage link under General Settings.

This will give you a list of commenting domains you manage. From here use the **Export Comments** link to queue an export of the comments for the corresponding domain:

Afterwards you'll be presented with a link to download your comments into the JS-Kit XML format.

Note that the time it takes to export your comments will depend on the number of comments within the corresponding domain.

#### Importing into Disqus

Once your comments have been successfully exported from JS-Kit, you'll be able to import them into Disqus. Use the Disqus JS-Kit importer on the Disqus imports page.

Select the forum to which you'd like to upload your comments. Choose the XML file to upload. Designate **JS-Kit (Echo)** as the importer option to be used.

After importing you'll want to verify that your thread URLs are correct. If comments aren't successfully displaying on your website, you may still need to migrate your threads to the correct URL using our URL mapper.

#### Conditions:

*These conditions are important, and if not met, can prevent your comments from being imported properly — please note each requirement carefully.*

-   Comments without an author's name aren't able to be imported.

-   Comments can't be imported as threaded conversations.

-   Guest comments without an assigned email can't be claimed by a registered Disqus account.

-   The permalinks of each comment within a comment thread must match.

-   Media attachments are not currently able to be imported. We'd suggest saving these separately.

-   Thread titles must be set correctly within the XML export before being imported. If not then they may appear as the thread's URL in your moderation panel.

### Importing comments from WordPress {#importing-comments-from-wordpress}

To bring your old comments into Disqus, they must be exported from WordPress and imported into Disqus. There are two ways to do this:

-   Automatic Import

-   Manual Import

-   Considerations

#### Automatic Import

Use the **Import Comments** button located in the **Syncing** tab in the plugin admin settings. This feature requires your API app credentials in the Site Configuration section before the import can be started. This will automate the export (from WordPress) and import (to Disqus) process. Note that imports can take up to 24 hours to complete.

If the Automatic Import method does not work for your site, use the Manual Import option below.

#### Manual Import

If the export comments function isn't working on your site, you have the alternate option of exporting your entire site from WordPress and importing manually. This can be done with the following steps:

1.  **Deactivate** all plugins except Disqus at your WP Dashboard > Plugins page.

2.  **Export** your WordPress site in WXR format at your WP Dashboard > Tools > Export.

*Note: Sometimes the WXR file cuts off prematurely, resulting in not all comments being exported. This can be caused by Wordpress not allocating enough resources. In that case, we recommend exporting your blog in date segments to keep the file size down. Click the Posts option and select a date range:*

3\. **Import** your exported WXR file.

Troubleshoot your import file if you run into any errors. Note that imports can take up to 24 hours to complete.

-   Make sure the files you're importing are valid XML. You can use the W3C xml validating tool to check.

-   Compressed files (e.g. .zip, .gz) can't be read by default, so make sure you've decompressed these before uploading them.

-   The importer will filter out duplicate comments *unless* you've changed some of the comment data.

-   Email addresses are unique identifiers in Disqus, so make sure each unique user has their own email address before importing. Otherwise all comments will appear from the same user.

-   Imported comments **can't** be permanently deleted. Consider following our guidelines for development sites to make sure the data you're importing is correct. You can register a new forum if you have imported the wrong comments.

-   **I'm not seeing custom avatars attached to my comments** We no longer support Gravatars associated with the email address used when importing guest comments.

-   Troubleshooting Imports

-   Data synchronization

-   Importing & Exporting

-   How long will my import take to finish?

### Importing Typepad comments {#importing-typepad-comments}

While Disqus has no official Typepad importer, we do support standard WordPress imports. This article details how to import your Typepad blog into WordPress so that you can then export it from WordPress and into Disqus.

Note that while this method should work, it isn't fully supported by Disqus and we can't help troubleshoot any unforeseen quirks.

#### Instructions

#### Step 1.

#### Export your Typepad Blog

You'll first need to get a file containing all of your blog contents. See Typepad's documentation for importing and exporting.

#### Step 2.

#### Create a self-hosted WordPress site

The next step involves setting up a WordPress site that you'll use to import the blog content. Don't worry if you have access to a server to do this, you can always install Wordpress locally using these instructions.

#### Step 3.

#### Import your blog into WordPress

Now you need to get all that content (including comments) into WordPress. Importing from Typepad is the same as Movable Type, so you can use these instructions. Pay attention to the "Forcing WordPress to Use the Movable Type Permalink Structure" section to make sure your comments are imported with the right links.

#### Step 4.

#### Export from WordPress and import to Disqus

All that's left is to get an export file from WordPress in your Dashboard > Tools > Export. Depending on the size of your blog, you may need to export this in multiple files based on month/year. Then import to Disqus (see: Manual Import) using those file(s).

### Migration Tools {#migration-tools}

The migration tools let you move discussion threads on your site to a new thread. Common scenarios when you might use this are:

-   You change your domain. Example: You change your web address from 0 to new-awesome-url.com.

-   You change your CMS/blogging system and, as a result, the article URL structure is no longer the same. Example: You move from Blogger to Wordpress.

-   Two different discussions need to be merged into one.

-   Once two threads are merged together, the merge **cannot** be undone.

-   The migration tools will **not** migrate comments to a new Disqus shortname, so be sure to use the same shortname on the new site. This will ensure that your migrated threads are tied to the correct Disqus account.

-   Migrations can take up to 24 hours to finish.

-   You can check on the migration progress during that time by looking for the updated comment threads on a page where they did not previously exist.

-   If you're using a custom *disqus_identifier* this will still play a role.

-   If two threads are merged into the same URL and have different identifiers as well, the new thread will contain **both** identifiers.

#### Where to find these tools

Go to your Disqus Admin and click Tools > Migration Tools (towards the bottom of the navigation panel on the left).
​

#### Migration options available

-   Domain Migration Wizard - If the base domain of your website has recently changed, you'll need to migrate the URLs of your commenting threads to use the new base domain (ex: 0 -> new-awesome-url.com).

-   URL Mapper - If the URL slug of your site's posts has recently changed, you'll need to use our URL Mapper to migrate your commenting threads from the old URL to the new one (ex: yourdomain.com/p=23 -> yourdomain.com/new-page-slug).

-   Redirect Crawler - If you've already set up 301 redirects for your site's pages, then use our Redirect Crawler to crawl your site's pages and migrate the URLs of your commenting threads automatically. Recommended for advanced users only.

### Moving from Blogger to WordPress {#moving-from-blogger-to-wordpress}

Moving from Blogger to WordPress can seem daunting at first, but with our migration tools and this step-by-step guide you'll be on your way in no time!

1.  Register a Disqus forum for your comments if you haven't already.

2.  Import your Blogger comments into Disqus. **The following step is absolutely necessary as Blogger and WordPress URL structures are not the same. This is the most important part of the moving process.** See: *Differences between Blogger and WordPress URL structures* below.

3.  Upload a URL map to migrate your threads' URLs from their old Blogger URL structure to their new WordPress URL structure.

4.  Install and setup the Disqus WordPress plugin on your WordPress installation.

Sample Blogger URL: 0 Sample self-hosted WordPress URL: 1

#### Notable differences:

-   No .blogspot in the base domain of the WordPress URL

-   No .html appended to the end of the WordPress URL

These are just examples. Your Blogger and WordPress URL structures may have additional differences. Read more about the URL mapper.

### Redirect Crawler {#redirect-crawler}

This migration tool will follow 301 redirects and migrate comments automatically. Each old permalink **must** redirect to the proper new permalink rather than the root domain.

For example, 0 needs to redirect to 1

-   0
-   0
-   0
-   0
-   0
-   0

After these 301 redirects are functioning properly, you can start our 301 redirect crawler by clicking the Start Crawler button on the Tools > Migrate Threads page.

*Note: Migrations can take up to 24 hours to finish, and we suggest checking its progress occasionally during that time (by looking for the updated comment threads on a page where they did not previously exist).*

#### Running the URL Mapper alongside the Domain Migration or Redirect Crawler

If you've run the Domain Migration Wizard or Redirect Crawler and your threads have not migrated after 24 hours, it is possible that the threads are not set up to properly migrate with that tool. Try using the URL Mapper to input the exact thread URLs that must be changed. You can run the URL Mapper concurrently with the Domain Migration Wizard or Redirect Crawler, as they will not conflict with one another.

-   URL Mapper - if the slugs of the URLs have changed

-   Domain Migration Wizard - if only the base domain has changed

### Syncing with WordPress {#syncing-with-wordpress}

Syncing may be enabled between WordPress and Disqus using the WordPress plugin. This will copy comments posted in Disqus to the WordPress native comment system, so that they appear in both locations. Comments are synced to your WordPress database using a webhook method starting in version 3.0 instead of wp-cron which is both more reliable and secure for your website.
​
As Disqus replaces the native WP comments system, this will not change what appears on your page when Disqus loads. However, the comments synced from Disqus back to the native WP comment system will appear if Disqus is no longer installed, and will be included in the source code of the page they were posted to.

Syncing is not enabled by default in the WordPress plugin. To enable syncing, you will need to set it up either via **Automatic Installation** in the plugin or manually by setting up an API application and entering the credentials in your **Site Configuration** (see instructions below). Once you've entered your API credentials, enable Comment Syncing in the Syncing tab of the plugin.
​
​

#### How to manually set up Comment Syncing

When creating an API application to sync comments, make sure your API credentials are entered in the **Site Configuration** tab. Go to the **Syncing** tab to enable comment syncing.

1.  Create an application (or use an existing application you own) and make note of the following which you will need to copy and paste: ***public key, secret key, & access token***

2.  Change your application "Default Access:" drop-down setting to "Read, write, and manage forums" so that your application can create syncing webhooks for your forum

3.  Paste the keys and token from the step above into the correct fields in ***Site Configuration*** tab the your WordPress plugin settings.

4.  Click on the ***Syncing*** tab, and select "***Manually Sync Comments"***.

5.  Select a date range and click "***Run Manual Sync"***

#### How to pause Comment Syncing

To pause Comment Syncing, simply click **Pause Auto Syncing** in the Syncing tab of the WordPress plugin.

#### Which comments will be synced?

When a comment is posted or edited, this triggers the comment to be synced with WordPress. It supports all comment states including approved, pending, spam, and deleted comments.
​

#### Troubleshooting

1\. Check the webhook link here: 0

If that link is not returning anything or not loading, something is blocking outside sources like Disqus from making requests to your site.

If the REST API is blocked for Disqus or your entire site, you’ll want to ensure that it is enabled. You can follow instructions here to enable REST API and troubleshoot why it may not be enabled:
0

More general information on the REST API may be found here: 0

2\. This may also be caused by a conflict with ones of your current plugins (often a security plugin). You can test this by temporarily disabling all plugins except Disqus. (These are not permanent recommendations; they are simply to help isolate the issue.) Once all themes (or the suspected themes) have been disabled, try to click the Syncing option within the Disqus plugin to sync your comments.

Specific conflicting themes that we've identified so far:

-   Disable REST API

-   iThemes Security

3\. Additionally, you may want to ensure that there isn't a firewall on your site that could be blocking auto syncing.

4\. If auto syncing is still not working, we recommend manually syncing comments in the Disqus plugin.

5\. If Manual syncing is not working for you, be sure that you have your trusted domain set in the settings of your API application.

#### Exporting Issues

Alternatively, if you're having issues with copying comments from WordPress to Disqus, see Exporting comments from WordPress to Disqus.

#### Other Issues

#### Comments synced to WordPress have the wrong status

To make sure that comments have the most-up-to-date status, we'd recommend using the **manual sync option** within the settings of the Disqus WP plugin. You'll want to overwrite the previously synced comments in order to make sure the correct comments are synced.

To prevent comments from being synced automatically in the future, you can disable this feature within the Syncing settings of the WordPress plugin.

### URL Mapper {#url-mapper}

This manual way of migrating comments is required when more than just the base domain of a thread's URL has changed. This method requires you to use a CSV (comma-separated values) file, which you can download on the Discussions > Tools > Migrate Threads > URL Mapper page.

If needed, see our additional guide on How to download, edit, and upload a URL Map CSV file (including a list of software you can use).

This URL Mapper tool is **not** required for sites that have moved from HTTP to HTTPS because Disqus will automatically load the correct discussion thread regardless of protocol.

#### Usage:

1\. **Download** a CSV file containing a list of thread URLs which currently belong to your forum by clicking the download link.

2\. After you've received this file via email, download and open it as a spreadsheet on your computer. You should see a single column of URLs.

3\. **Input new URLs**, which you're looking to map the old URL **to**, in the right column.

***A note on which thread URLs to include:*** You are not required to include all thread URLs in the CSV and we recommend removing any thread URLs that do not require migration. Only thread URLs listed in the CSV will be migrated; any thread URL not included in the CSV will be left unchanged.

4\. **Upload** the CSV file at Migrate Threads > Upload URL Map.

-   Migrations can take up to 24 hours to complete, so it is best to check back every few hours.

-   **A note regarding very large files:** If your migration includes tens of thousands of URLs, we recommend breaking your CSV into a few smaller files and uploading them individually. CSV files uploaded to the URL Mapper should not exceed 2.5mb in size.

-   It is not possible to migrate a thread from one shortname to another

-   You can merge two existing URLs by mapping the first URL to the second URL, and all comments will be combined in the resulting thread.

-   When threads are combined via the URL Mapper, Column B's threads will take on the title details of the corresponding threads from Column A.

-   How to download, edit, and upload a URL Map CSV

-   Domain Migration Wizard - if only the base domain has changed

-   Redirect Crawler - if you have setup 301 redirects

## Installation {#cat-installation}

### Add Disqus to Static Pages in Blogger {#add-disqus-to-static-pages-in-blogger}

There might be times when you would like comments to appear on pages other than articles. This tutorial will guide you through the process of adding Disqus to static pages in Blogger.

#### Step 1.

Go to the admin page for the Blogger site you would like to change and choose, "Template".

#### Step 2.

Click "Edit HTML".

#### Step 3.

To locate the Disqus Widget code, click inside the text box that contains the template code and press command+F. This will open a search box; enter "Disqus" into the search box and press enter. Alternatively, scroll all the way down to the bottom of the template and locate, "title='Disqus'".

#### Step 4.

Click the arrow shown below to expand the template code:

Click the subsequent arrow to expand the template code again:

It should look like this:

#### Step 5.

Delete the following tags from the template:

#### Step 6.

Click the "Save template" button.

#### Step 7.

Check your static pages to bask in the glory of your accomplishment.

### Adding Disqus to your site {#adding-disqus-to-your-site}

Disqus has many integrations available which make installation easy. The most popular integrations also come with built-in importing and syncing tools so you can bring in your old comments as well. Before installing, make sure you've registered a Disqus shortname, and this will be used to reference all of your comments and settings.

#### WordPress

#### Installation

-   WordPress plugin installation

-   Manually install Disqus on WordPress

-   Setting up SSO on WordPress

#### Importing and syncing comments

You can export your old comments from WordPress and import them into Disqus, or sync comments made in Disqus back to WordPress

-   Importing comments from WordPress

-   Syncing with WordPres

#### Troubleshooting

-   Troubleshooting WordPress

-   WordPress forums

The WordPress community is a great place for answers for things Disqus is unable to assist with.

#### More topics

-   Can Disqus be used on WordPress.com?

####  Blogger

#### Installation

*Because of a change made on Blogger's end, the automatic installation method in this video is no longer supported. To install Disqus on a Blogger site, please use our "Manually Install on Blogger" instructions in the list below*
​

-   Blogger widget installation

-   Manually install on Blogger

-   Add Disqus to Static Pages in Blogger

-   Updating Blogger templates to support Internet Explorer

-   Loading Disqus on mobile templates

####  Importing and syncing

We're currently unable to support import and syncing functionality for Blogger. We recognize that syncing and import functionality is an important feature for many Blogger blogs and we apologize for any pains this disruption in service may cause. We want publishers to have a great experience and hope to restore import and syncing services to Blogger in the future.

Learn more about exporting for backup purposes

#### Troubleshooting

-   Blogger troubleshooting

-   Blogger product forums

####  Tumblr

#### Installation

#### Tumblr theme-supported installations

If the theme author included Disqus integration, this is how you'd enable it.

#### Tumblr manual installation

Not every theme includes Disqus or integrated it correctly. This is a guide on manual installation.

#### Troubleshooting

#### Tumblr Troubleshooting

####  Typepad

#### Installation

#### Typepad installation instructions

#### Importing

#### Importing Typepad comments

####  Movable Type

#### Installation

#### MovableType plugin integration

####  Drupal

#### Installation

#### Drupal plugin integration

####  More integrations

#### Joomla, Squarespace, and other managed integrations

Integrations that are managed for their platforms as plugins or widgets.

#### All available integrations

####  Developer integrations

How to customize your Disqus integration on the web, add to mobile apps and more.

#### Developer integration guide

### Can Disqus be used on WordPress.com sites? {#can-disqus-be-used-on-wordpress-com-sites}

Wordpress has two offerings: a self-hosted option available at wordpress.org, and a version that is hosted by Wordpress at wordpress.com.

If you'd like to use Wordpress for free, you'll want to choose the self-hosted option which is available from wordpress.org. You'll need to set up your site locally and host it yourself. A free hosting service may be found here.

If you'd prefer to have Wordpress host your site, you can choose the option available at wordpress.com. Using Disqus on a 0 site requires a WordPress.com Business plan which supports third-party plugins like Disqus.

To install a plugin from the Plugin Directory, head to My Site → Plugins, and click the Add Plugin button. Search for "Disqus" and select the plugin authored by Disqus. Click "Install" and the plugin will be added to your site.

For more information on adding Disqus to your WordPress.com site, check out WordPress' documentation here.

### Configuring Disqus on your site {#configuring-disqus-on-your-site}

Engage by Disqus is the easiest-to-use and feature-rich commenting and community platform for publishers.

Engage offers robust administrative and moderation tools to help you keep your community vibrant and welcoming and to manage the data stored in Disqus.

Each section below contains best practices to consider when configuring your Engage forum. We recommend configuring your forum settings and planning your moderation practices to best suit your goals and the needs of your unique site/community. See the Ads Launch Pad for monetizing content.

-   Basic settings

-   Filters and community rules

-   Advanced configuration

-   Moderating comments

-   Managing discussions

#### Basic settings

The basic configuration details for a site include appearance settings, site information, and basic rules that apply to commenters on your site.

#### Appearance settings

Learn more about updating your background color and other appearance tweaks.

#### Site identity information

What gets seen around the network.

#### Multi-lingual websites

Configure Disqus to appear in your community's preferred language.

#### Filters and community rules

Setting up automated rules that single out individuals or classes of users is a good way to make moderation easier. These guides provide an overview of how to use these features.

#### [Community rules](#moderation-rules) & [adding moderators to your site](#how-to-add-admins-and-moderators-to-your-organization)

Set rules, add additional moderators/admins to your site, and restrict certain words.

#### Using the "Ban User" and "Trust User" controls

Manage who is blocked from commenting and who can be immune from existing moderation rules.

#### Advanced configuration

Options targeted at advanced users or certain use cases.

#### Adding a default commenter avatar

How to upload a custom avatar for users who don't have one.

#### Using categories

How to set categories and filter data with the API.

#### Configure trusted domains

Set a trusted domain to keep your Disqus shortname from loading on unwanted sites.

#### Moderating comments

How to use the moderation tools.

#### Moderation panel overview

Moderate, get context, search and edit comments on your site through the moderation panel.

#### User reputation

Take advantage of user reputation to help you moderate.

#### Inline and email moderation

Moderate comments within the public discussion thread or from your inbox.

#### Managing discussions

Disqus provides a number of tools to help you manage discussions on your site. This includes migrating, closing, importing/exporting or changing basic details about each.

#### Introduction to the discussions editor

Manage discussions individually by updating titles, URLs, authors and closing threads.

#### Using the migration tools

Change your site to a new address and comments went missing? Update them using any of the migration tools.

#### Importing and Exporting

Import existing comments and discussions from another system.

### How to use trusted domains {#how-to-use-trusted-domains}

Trusted domains are set by websites to specify which domains are allowed to create and load new threads with the Disqus javascript embed. As a site owner it's highly recommended that you set at least one trusted domain. It's otherwise possible for any website to load your Disqus shortname and contribute comments and threads.
​
Note that it is not possible to retroactively set a trusted domain on existing threads. The domain will only apply to newly created threads.

#### Setting a trusted domain

Go to your **Disqus Admin > Setup >Advanced**. Locate the "Trusted domains" box and enter your domains there.

#### Best practices and troubleshooting

-   Sub-domains are unnecessary, as everything is covered in the base domain. So using 0 will cover 1 and 2 .

-   Don't include 0 in your trusted domain, or else your comments may not load.

-   As part of the Best Practices for staging sites you can use the trusted domain as a check to make sure you don't load accidentally load a production Disqus shortname on a staging site and vice versa.

#### Also see

-   Why are comments posted to other sites showing up in my Disqus admin?

-   I've installed Disqus on my website but it isn't loading. What do I do now?

### Manually adding a Disqus gadget to Blogger {#manually-adding-a-disqus-gadget-to-blogger}

If the Disqus gadget installer isn't working, you have the option of manually installing the gadget on your Blogger site. This will require editing your Blogger template HTML, so it won't work with Dynamic Views templates.

#### Add a new gadget

1\. Go to your Blogger "Layout" section and click **Add a gadget** in the sidebar.
2. In the Add a Gadget popup, scroll down to find the **HTML/Javascript** gadget and click the + button.
3. Enter *Disqus* as the **title** and the following code for the **content**:

    <!-- Disqus comments gadget -->

4. Click save and the window will close.
5. Click **Save arrangement** in the Layout viewer.

#### Add the Disqus code to your template

1\. Go to your blog's **Layout** section, click the dropdown menu on the Customize button for your theme, and click "Edit HTML":

2\. Click inside the text area and search for the widget you just created in your HTML template by pressing Ctrl-F (Command-F on OSX) then typing *Disqus*. You should find the following line:

    <b:widget id='HTML1' locked='false' title='Disqus' type='HTML'>

3\. Change that line to add **mobile='yes'** to load Disqus on your mobile template. It will look like this when you're done:

    <b:widget id='HTML1' locked='false' mobile='yes' title='Disqus' type='HTML'>

4\. Below that locate and **DELETE** the following code right before the closing tag. The section you're deleting should look like this:

    <b:includable id='main'>
      <!-- only display title if it's non-empty -->
      <b:if cond='data:title != &quot;&quot;'>
        <data:title/>
      </b:if>
      <div class='widget-content'>
        <data:content/>
      </div>
      <b:include name='quickedit'/>
    </b:includable>

5\. BEFORE the closing 0 tag, add the following Disqus code (remember to replace "**EXAMPLE**" with your forum shortname and be sure to leave ''' in front of your shortname and '';' after it.):

    <b:includable id='main'>
                <script type='text/javascript'>
                    var disqus_shortname = 'EXAMPLE';
                    var disqus_blogger_current_url = "<data:blog.canonicalUrl/>";

                    if (!disqus_blogger_current_url.length) {
                        disqus_blogger_current_url = "<data:blog.url/>";
                    }

                    var disqus_blogger_homepage_url = "<data:blog.homepageUrl/>";
                    var disqus_blogger_canonical_homepage_url = "<data:blog.canonicalHomepageUrl/>";
                </script>

                <b:if cond='data:blog.pageType == &quot;item&quot;'>
                    <style type='text/css'>
                        #comments {display:none;}
                    </style>

                    <script type='text/javascript'>
                        (function() {
                            var bloggerjs = document.createElement('script');
                            bloggerjs.type = 'text/javascript';
                            bloggerjs.async = true;
                            bloggerjs.src = '//'+disqus_shortname+'.disqus.com/blogger_item.js';
                            (document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(bloggerjs);
                        })();

                    </script>
                </b:if>
                    <style type='text/css'>
                        .post-comment-link { visibility: hidden; }
                    </style>

                    <script type='text/javascript'>
                    (function() {
                        var bloggerjs = document.createElement('script');
                        bloggerjs.type = 'text/javascript';
                        bloggerjs.async = true;
                        bloggerjs.src = '//'+disqus_shortname+'.disqus.com/blogger_index.js';
                        (document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(bloggerjs);
                    })();

                </script>
    </b:includable>

6\. Click **Save template**. Assuming there are no errors, Disqus should properly show up on your site now.
​
7. (Optional) Verify that the meta tags in your Blogger Template do not force Internet Explorer to load using IE7 standards.

For more information see Troubleshooting Disqus in Internet Explorer 8/9/10

-   Blogger Troubleshooting

### Manually install Disqus on WordPress {#manually-install-disqus-on-wordpress}

If the Disqus WordPress (WP) plugin isn't functioning properly within your theme, you have the alternative option of installing Disqus manually using our Universal Embed code.

The following functionality **is not** available under the manual installation:

-   Syncing comments locally

-   Exporting comments to Disqus automatically. You'll need to import comments manually.

-   Accessing the mod-panel via the WP admin. You'll need to go through Disqus.com.

Navigate to the theme editor within your WP installation on the Appearance > Editor page.

Locate the 'comments.php' file in the theme files listed on the right side of your screen. Backup this existing code by copy and pasting it into a text-file.

Afterwards, replace the code in 'comments.php' with the code snippet below that includes the Universal Embed code with an 0 statement, which verifies that comments are enabled for the page in question.

**Note**: Don't forget to change 0 to your forum's shortname.

    <?php if (comments_open()) :?>
    <div id="disqus_thread"></div>
    <script>
        /**
         *  RECOMMENDED CONFIGURATION VARIABLES: EDIT AND UNCOMMENT THE SECTION BELOW TO INSERT DYNAMIC VALUES FROM YOUR PLATFORM OR CMS.
         *  LEARN WHY DEFINING THESE VARIABLES IS IMPORTANT: 0
         */
        /*
        var disqus_config = function () {
            this.page.url = PAGE_URL;  // Replace PAGE_URL with your page's canonical URL variable
            this.page.identifier = PAGE_IDENTIFIER; // Replace PAGE_IDENTIFIER with your page's unique identifier variable
        };
        */
        (function() {  // DON'T EDIT BELOW THIS LINE
            var d = document, s = d.createElement('script');

            s.src = '//EXAMPLE.disqus.com/embed.js';

            s.setAttribute('data-timestamp', +new Date());
            (d.head || d.body).appendChild(s);
        })();
    </script>
    <noscript>Please enable JavaScript to view the comments powered by Disqus.</noscript>
    <?php endif; // comments_open ?>

#### Using the WordPress 2013 Theme:

In the 2013 default theme for WordPress, you'll need to enable the plugin setting to output the JavaScript in the footer and also add the 0 class to the disqus_thread div.

### Moderation Profiles {#moderation-profiles}

In completing the Installation and Setup flow with Disqus, you'll be prompted to select a Moderation Profile corresponding to how strictly you'd like to filter content. While this allows you to set a baseline with a single click, all of the components of each profile may be individually customized from your Moderation Settings page.

There are two profiles to choose from, Strict and Balanced. Strict is designed to catch and automatically remove risky content, minimizing the amount Moderator oversight needed. Balanced is designed to be more open, with a few automated rules in place to help ease moderation needs.

#### Strict

-   Images, Videos, and Links will not be allowed in comments

-   Guest comments will not be allowed

-   Comments that are flagged 3 times will be sent to pending

-   Threads will be automatically closed after 30 days

-   Comments containing restricted words will be automatically deleted

-   Toxic comments will be automatically deleted

#### Balanced

-   Images, Videos, and Links will be allowed in comments

-   Guest comments will be allowed

-   Comments that are flagged 5 times will be sent to pending

-   Comments containing restricted words will be automatically deleted

-   Toxic comments will require moderator approval to be displayed

If you'd like to later change which Moderation Profile your site is using, this may be done from the Moderation section of the Installation page.

### Multi-lingual websites {#multi-lingual-websites}

To set a default language for every discussion on your site, choose a language in your Disqus Admin > Setup > Appearance.

#### Language overrides

Using the 0 JavaScript variable, you can dynamically load the Disqus embed in different languages on a per-page basis. For example, to load the embed in French:

    var disqus_config = function () {
      this.language = "fr";
    };

For a full list of language codes, see Languages on Transifex. Note however that Disqus does not support all languages offered in Transifex. For a full list of supported languages, see the Disqus project on Transifex.

#### Comment counts

Comment count strings must be translated manually in your Disqus Admin > Settings > Community under **Comment Count Link**. This can only be applied globally across your site.

#### Best Practices

-   Avoid letting the user pick the language for the discussion. In most cases, letting users who speak different languages comment in the same discussion degrades the relevance of the conversation for everyone.

-   Unless your moderators are multi-lingual, it's usually better to create separate forums for each language. Otherwise it's more difficult to split moderation duties by language if multiple languages are used in the same forum.

-   If using multiple languages on a single forum, consider removing the text from comment counts entirely and overlaying the number over a comment symbol.

-   Translating Disqus

### Publisher Quick Start Guide {#publisher-quick-start-guide}

Note: if you are not a moderator of any forum or you are currently logged out of Disqus, you will get redirected to this page when trying to access the admin panel. Below is how you can create a Disqus forum, if you'd like.

Disqus works on virtually any type of website or blogging platform, and is very simple to install through the use of our embed code. This guide will outline the steps to setting up Disqus on your website.

#### - Register with Disqus

Before going any further, you will need to register your website with Disqus. You will also need to create a user profile in order to login and administer this website. During registration, you will pick a shortname for your site, which is how Disqus identifies your website community in the system.

An organization will be automatically created which will contain your new site and any additional sites you choose to create in the future.

#### - Install Disqus

Disqus is compatible with many popular blogging platforms, content management systems, and virtually any custom website. Some examples of blogging platforms are WordPress, Tumblr, and Blogger. Visit the installation instructions section of the website to find the instructions for your platform.
​

Not using any of the suggested platforms? Use our Universal Embed Code. Follow our instructions closely and make sure you set the correct shortname variable so that comments go to your forum.

#### - Configure Disqus

Disqus provides publishers with the tools they need to cultivate thriving communities. Your forum settings allow you to add moderators and admins, set community rules, and more. Review our Getting Started guide for how to start growing your community.

#### After everything is setup

#### Moderate your comments

Disqus offers many tools to help you easily manage your community; learn more on our Moderation page.
​

#### Home Feed

Your Home is the main page you see when you visit 0. For more information about how your site's community will show up on Disqus, how it can be followed, and more, check out Your Homepage on Disqus.

### Translating Disqus {#translating-disqus}

Disqus currently supports dozens of languages, ranging from Arabic to Ukrainian. For a full list of supported languages, see the Disqus project on Transifex.

#### How do I change my site's language?

Site language can be changed at the Disqus admin > Settings > General page.

#### What if my language isn't listed?

You can request a new language be proposed for translation. To do so:

1.  Visit the Disqus project on Transifex.

2.  Select "Request language" (requires being logged-in to Transifex).

3.  Choose your desired new language. If your language is not listed, that means it is available for translation but requires more translations before it can appear in Disqus. In this case, see How can I help translate Disqus? below.

4.  Select "Request team".

Please note this does not guarantee this language will be accepted for translation. New translation proposals are reviewed on a regular basis.

#### How can I help translate Disqus?

Disqus uses Transifex to crowdsource translations. To help translate Disqus:

1.  Visit the Disqus project on Transifex.

2.  Create an account if you have not already by clicking “Help translate Disqus" in the upper right-hand corner. After creating your account navigate back to Disqus project.

3.  Select “Join team” and select a language to request to be added.

4.  Once accepted, click “Translate” in upper right corner of the Disqus project dashboard, and follow the prompts to get started.

The community tries to approve new translators regularly. If you aren't approved right away, please be patient.

#### Become A Translations Reviewer

Disqus translations are a community-powered effort. The community is always looking for passionate individuals who can help review translations, making the process even faster and more accurate. Translations reviewer responsibilities include:

-   Verifying correctness of proposed translations.

-   Approving all new translations at least 1x/week.

-   Not letting bad words through.

In general, translations reviewers should be active members in the Transifex community who are already involved in translating other Transifex projects.

To apply as a translations reviewer:

1.  Visit the Disqus project on Transifex.

2.  Follow the above instructions to create an account ***and*** join the Disqus team for a language, if you haven’t already.

3.  Click “Translate” on the language you wish to help review.

4.  Choose any of the coordinators to view their profile and send them a message explaining why you'd be a great translations reviewer.

#### Get updates when new translation strings are available

Transifex allows users to follow projects they help translate and watch them for any updates. Translators can also choose to watch only a specific language or a project.

-   Notice settings for managing notifications.

-   A list of languages you're currently watching.

-   Additional information on being notified.

-   Can Disqus be loaded in different languages per-page?

### Universal Embed Code {#universal-embed-code}

Disqus can be installed on virtually any website using the universal JavaScript embed code. The following documentation is meant for developers only. Disqus also provides installation instructions for dozens of popular blogging and website platforms such as WordPress, Blogger, and more.

-   Make sure you've registered your website with Disqus. Read the Quickstart Guide for more information.

-   You will need to be able to edit the HTML of the website you are installing Disqus on.

-   To install Disqus, you will need to know your forum shortname as registered on Disqus.

#### Embed code

This is the JavaScript embed code which loads and displays Disqus on your site, typically on the individual article or post pages. The **disqus_thread** ID is where the postbox is loaded, so make sure to include it above the embed script as shown below.

**Note**: Don't forget to change EXAMPLE to your forum's shortname.

    <div id="disqus_thread"></div>
    <script>
        /**
         *  RECOMMENDED CONFIGURATION VARIABLES: EDIT AND UNCOMMENT
         *  THE SECTION BELOW TO INSERT DYNAMIC VALUES FROM YOUR
         *  PLATFORM OR CMS.
         *
         *  LEARN WHY DEFINING THESE VARIABLES IS IMPORTANT:
         *  0
         */
        /*
        var disqus_config = function () {
            // Replace PAGE_URL with your page's canonical URL variable
            this.page.url = PAGE_URL;

            // Replace PAGE_IDENTIFIER with your page's unique identifier variable
            this.page.identifier = PAGE_IDENTIFIER;
        };
        */

        (function() {  // REQUIRED CONFIGURATION VARIABLE: EDIT THE SHORTNAME BELOW
            var d = document, s = d.createElement('script');

            // IMPORTANT: Replace EXAMPLE with your forum shortname!
            s.src = '0';

            s.setAttribute('data-timestamp', +new Date());
            (d.head || d.body).appendChild(s);
        })();
    </script>
    <noscript>
        Please enable JavaScript to view the
        <a href="`https://disqus.com/?ref_noscript`" rel="nofollow">
            comments powered by Disqus.
        </a>
    </noscript>

#### Configuration variables

Within the above embed code, there are configuration variables which tell Disqus how the system should work and behave. EXAMPLE is your "shortname" and should be replaced to tell Disqus which website account (called a forum on Disqus) this system belongs to.

**this.page.url** tells Disqus the location of the page for permalinking purposes, this value will also uniquely identify the page and discussion thread if 0 (below) is not set. The url must contain a protocol (http or https).

**this.page.identifier** tells Disqus a unique value, used to identify the page and discussion thread

There are many more configuration variables available, but these are the most important. To learn more about these and the other configuration variables, read JavaScript configuration variables.

#### Comment Counts

-   Guide to adding comment count links to your home page

#### Easy Installation

Some platforms also provide simple steps for installation that do not require any javascript coding. These integrations can be found on our installation page.
​

### Updating a Blogger template to support all versions of Internet Explorer {#updating-a-blogger-template-to-support-all-versions-of-inter}

Many older Blogger templates include a meta tag that forces all versions of Internet Explorer to behave like Internet Explorer 7. Because Disqus supports Internet Explorer 8 and higher, this will break functionality and users won't be able to comment even with the latest version of Internet Explorer. Below are instructions on how to fix this.

1\. Go to your blog's **Template** section and then click the "Edit template" button

2\. Locate the following line within the 0...</head> tags:

    <meta content='IE=EmulateIE7' http-equiv='X-UA-Compatible '/>

3\. Replace it with this line:

    <meta http-equiv="X-UA-Compatible" content="IE=9; IE=8; IE=7; IE=EDGE; chrome=1" />

4\. Click **Save Template** to apply your changes.
​

If you're using a new Blogger template and the above meta tag is not already in the template, no changes are necessary and IE readers will be able to view Disqus comments.

### What's a shortname? {#what-s-a-shortname}

A shortname is the unique identifier assigned to a Disqus site. All the comments posted to a site are referenced with the shortname. The shortname tells Disqus to load only your site's comments, as well as the settings specified in your Disqus admin.

Yes. To manage comments in Disqus you will need to register your site and install Disqus on your site using the shortname registered.
​

#### Choosing a shortname

There are a few things to keep in mind when choosing a shortname:

-   Your shortname cannot be changed once it is registered. Be sure it's the one you want.

-   Your shortname will not appear publicly. However, your website name will show up in a number of places including (but not limited to): email notifications, the My Disqus tab, the Community tab, and the Discovery box.

*Adjust your website name in forum Settings > General.*

#### What are my shortnames?

You can access the list of your sites by first visting your Home Feed and selecting “Admin” from the user drop-down menu:

Then clicking the sandwich menu button on the top left of the Admin:

The shortname can be found in the address bar of your browser as **example**.disqus.com/admin or in your General Site Settings.

#### Using Disqus shortnames on different types of websites

After creating your shortname and registering your website, you now have to install Disqus and ensure that your site is using the correct shortname.
​
For some popular content management systems, Disqus can be integrated using a simple plugin provided by the site.

#### Wordpress

*Note: These instructions are for self-hosted Wordpress.org sites. Disqus cannot be installed on Wordpress.com sites. Learn more.*

-   In the left panel of your Wordpress admin, select **Plugins > Add New**

-   Search for "Disqus" and find the plugin provided by "Disqus".

-   Select **Install Now > Activate Plugin**

-   Proceed with the onscreen install instructions provided.

-   Log into your Disqus account, then choose the forum shortname you would like to install.

#### Blogger

-   Navigate to our Blogger install instructions. If you've registered a shortname, and you're logged in, there will be a button to add your shortname to your Blogger site.

-   Next, import any existing Blogger comments into Disqus at Tools > Import

-   Update your blogger template's meta tags for full Internet Explorer compatibility. See instructions here.

#### Tumblr

-   In Tumblr, visit Account icon > Edit Appearance > Edit Theme

-   In the shortname field, enter your shortname.

-   Save your theme and you're done.

#### Squarespace

-   Navigate to your Squarespace **Settings > blogging** page.

-   In the Disqus shortname field, enter your shortname.

All of these instructions, and more, can be found by clicking on the integration of your choice at our Install Page
​

### Will I lose comments if I uninstall Disqus? {#will-i-lose-comments-if-i-uninstall-disqus}

No. You can export your comments from Disqus at any time if you ever decide to remove Disqus from your site.
​
The Disqus for WordPress plugin supports the ability to automatically sync comments from Disqus back to WordPress. These comments will remain in WordPress should Disqus be deactivated or removed.

#### Before uninstalling

We recommend verifying all comments have been synced back to the WordPress comment system.

-   Comments already synced to WordPress can be viewed at the WordPress admin > Comments > All Comments page.

#### For non-WordPress sites

We also offer a custom XML export format for backup purposes.

#### Syncing to Blogger

We're currently unable to support syncing functionality for Blogger. We recognize that this functionality is an important feature for many Blogger blogs and we apologize for any pains this disruption in service may cause. We want publishers to have a great experience and hope to restore syncing services to Blogger in the future. While Blogger syncing is unavailable, your comments will be kept safely within Disqus as long as your Disqus shortname remains undeleted. You also have the option to export your comments for backup purposes.

## Known Issues {#cat-known-issues}

### Blogger Syncing and Importing {#blogger-syncing-and-importing}

Syncing is not currently working for some Blogger sites due to a re-factoring Blogger performed which causes some authorization processes to fail during the sync process. This may also results in emails stating the following: "You are receiving this email because you've chosen to sync your comments on Disqus with your Blogger blog. Unfortunately, we were not able to access this blog." We’re currently looking into our overall Blogger integration based on this and other changes on Blogger's end.

Your forum's comments are not being lost in the meantime and are all still stored in Disqus. They can also be exported into an XML file for backup purposes at any time from your forum's Tools > Import/Export page. Sorry for the current hassle and don't hesitate to contact us with any further questions.

#### Comments Synced to Wrong Blogger Post

In some cases, comments can be synced to the wrong post in Blogger. We're looking into this issue on our end. The status in which a comment is synced from Disqus is not updated if that status later changes, so it won't be possible to correct this issue within Disqus. We are working on improving this functionality, and appreciate your patience in the meantime.

#### Synced Comments Showing Site Owner Name Instead of Commenter Name

If your blog's settings allow only certain people to comment within Blogger, all comments synced from Disqus will be shown as authored by the blog owner instead of the commenter.
To prevent this, go to **Blogger → Settings → Comments** and set the **Who Can Comment?** option to Anyone.

#### Syncing status can sometimes show the incorrect count

#### Importing reply comments from Blogger to Disqus

Due to some limitations our Blogger import tool, reply comments within Blogger will be imported into Disqus as top level comments. Once Disqus is active on your Blogger site, reply comments will be available as normal within Disqus, but old Blogger comments can not be changed. We apologize for this inconvenience.

#### Feedback

If you would like to provide us additional feedback on any of these issues, please click here.

### Follower Notification Emails - Known Issue {#follower-notification-emails-known-issue}

When a new user follows you on Disqus, a notification email will be sent to your address. However, these notifications are handled a little differently in our system than other notifications. It is not currently possible to disable notifications for new followers, unless you turn off all notifications for your user account. For more information on how to disable all notifications, see Disqus Web Notifications.

#### Temporary Solution

Unsubscribe from Disqus notifications.

or;

If you are using Gmail, you can create a filter for Disqus follower notifications by searching for cases that have the subject **"is now following you on disqus"**. If don't want to see these in your inbox, you can stash them away in a folder, or send them to the trash. Other email providers offer similar filtering features.

#### Feedback

If you would like to provide us additional feedback on this issue, please click here.

### Known Issue: Disqus not loading via Spanish ISPs: Movistar/Telefonica de Espana {#known-issue-disqus-not-loading-via-spanish-isps-movistar-tel}

Some commenters and publishers with Disqus installed on their site from Spain and Latin America have recently been reporting issues in which the Disqus embed does not load. We’re currently investigating and your reports are helpful to us in pinpointing the issue.

**Please try opening the affected Disqus enabled webpage using the following troubleshooting steps and let us know what you find by emailing us at help+ISP@disqus.com.**

-   open in an incognito (or private mode) browser

-   try a different browser

-   ping disqus.com and a.disquscdn.com and copy/paste the results. Here's how: 0

-   copy your IP address from 0

-   try this workaround: change your routers DNS to this one provided by Google (8.8.8.8). Not sure how? Go to 0

Thank you for sharing this information with our team, we greatly appreciate it and we apologize for the inconvenience as we work to resolve this issue.

For all other general inquires, please visit our Support Hub at 0.

### Security Best Practices {#security-best-practices}

With recent industry-wide increases in

-   personal
-   data
-   breaches

, software security vulnerabilities and other malicious activity, it is a good time to remind you about security and privacy measures available to you as a Disqus user. While the Disqus platform has not been directly impacted by these events, we continue to monitor for and detect suspicious activity for purposes such as spamming, spoofing, pattern matching, privacy violations and the like.
​
Disqus exercises industry best practices and we are continually enhancing our own security measures, as well as cooperating with partners to improve the integrity of the ecosystem as a whole. However, there are measures you can take to enhance the security of web services you may use, including Disqus. Key measures include:

-   Regular changing and strengthening of passwords

-   Never use the same password across multiple websites.

-   Using unique passwords that include a combination of words, numbers, special characters and both upper and lower case letters

-   Use a password manager such as 1Password or Keepass

-   Enable two factor authentication on websites/services that support it including your email account

-   Keeping application and browser software up to date for security patches

-   Exercising healthy skepticism when confronted with any unusual or obfuscated link or email

-   Using new and/or unique email address to register sensitive social accounts

-   Turning locking or private profile settings on for accounts that are at risk of unwanted following

-   Avoiding and reporting websites that appear to violate privacy law and/or terms of service

We take the safety and privacy of users very seriously, so if you ever observe suspicious activity across Disqus’ network of communities, we would encourage you to contact us through the appropriate channel — security@disqus.com or privacy@disqus.com — with as much detailed information as possible.

### Twenty Thirteen/Fourteen/Fifteen (theme) Conflict in WordPress - Known Issue {#twenty-thirteen-fourteen-fifteen-theme-conflict-in-wordpress}

The Disqus embed currently loads wider than normal in Twenty Thirteen, Twenty Fourteen, and Twenty Fifteen themes. To fix this issue, you'll need to add a short line of code to your theme's CSS so it displays at the correct width. Twenty Sixteen does not require this fix. Follow the steps below:

1\. In your WordPress Admin, go to your Appearance Editor screen and select the Stylesheet (style.css) document from the right sidebar. It should be selected by default.

2\. Locate the following line within the style sheet:

    .site-content .entry-header,
    .site-content .entry-content,
    .site-content .entry-summary,
    .site-content .entry-meta,
    .page-content {
        margin: 0 auto;
        max-width: 474px;
    }

3\. Edit the code to add , 0 after 1 . It should look like this when you're done:

    .site-content .entry-header,
    .site-content .entry-content,
    .site-content .entry-summary,
    .site-content .entry-meta,
    .page-content,
    #disqus_thread {
      margin: 0 auto;
        max-width: 474px;
    }

4\. Save your changes. You should see Disqus using the same width as the article content now.

Feedback

If you would like to provide us additional feedback on this issue, please click here.

## Moderation {#cat-moderation}

### Advanced Moderation {#advanced-moderation}

Available with our Pro or Business plans, our AI-informed Advanced Moderation tooling will provide more specific categorization and controls to your site’s comments, allowing you to remove objectionable content with heightened precision and automation. For additional moderation tools available to all sites, please see our [Moderation Settings](#moderation-settings) and [Toxicity Filter](#toxic-mod-filter) documentation.

Within Disqus’ Advanced Moderation, there are a number of different categories of objectionable content. Each category can be restricted or allowed independently. Multiple categories may be applied to a single comment.
​
The categories are as follows:

-   Hate Speech

-   Violent Content

-   Sexual Content

-   Bullying

-   Promotion

#### **Severity**

Within each category, comments are also graded based on severity. Comments with a grading of “3” will be the most explicit or extreme content for that category. Comments with a grading of “1” will be the least extreme content that still fits the content category.

Below is a breakdown of the severity ratings within each category:
​

**Hate Speech**
*3 - Hate Speech*: Slurs, hate speech, promotion of hateful ideology
*2 - Slurs*: Negative stereotypes or jokes, degrading comments, denouncing slurs, challenging a protected group's morality or identity, violence against religion
*1 - Informational*: Positive stereotypes, informational statements, reclaimed slurs, references to hateful ideology, immorality of protected group's rights

**Violent Content**

*3 - Intimidation*: Serious and realistic threats, mentions of past violence
*2 - Instigation*: Calls for violence, destruction of property, calls for military action, calls for the death penalty outside a legal setting, mentions of self-harm/suicide
*1 - Description*: Denouncing acts of violence, soft threats (kicking, punching, etc.), violence against non-human subjects, descriptions of violence, gun usage, abortion, self-defense, calls for capital punishment in a legal setting, destruction of small personal belongings, violent jokes

**Sexual Content**

*3 - Explicit*: Intercourse, masturbation, porn, sex toys and genitalia
*2 - Intent & nudity*: Sexual intent, nudity and lingerie
*1 - Statements*: Informational statements that are sexual in nature, affectionate activities (kissing, hugging, etc.), flirting, pet names, relationship status, sexual insults and rejecting sexual advances

**Bullying**

*3 - Brutalizing*: Slurs or profane descriptors toward specific individuals, encouraging suicide or severe self-harm, severe violent threats toward specific individuals
*2 - Profane*: Non-profane insults toward specific individuals, encouraging non-severe self-harm, non-severe violent threats toward specific individuals, silencing or exclusion
*1 - Insults*: Profanity in a non-bullying context, playful teasing, self-deprecation, reclaimed slurs, degrading a person's belongings, bullying toward organizations, denouncing bullying

**Promotion**
(there is only one severity rating for Promotion)
*Promotion*: Asking for likes/follows/shares, advertising monthly newsletters/special promotions, asking for donations/payments, advertising products, selling pornography, giveaways

The severity descriptions above are also visible from your Moderation Rules section. Simply click into the white space of the rule to view the breakdown for that content category.

#### **Setting up Rules**

When setting up moderation rules on the content categories, please note that a rule for a certain severity level will also be applied to all severity levels above it. For example, if you set a rule to delete all comments that match the lowest tier of Bullying (1 - Insults), this rule will also delete comments labelled as Bullying 2 (Profane) and Bullying 3 (Brutalizing). If you instead set a rule to delete Bullying 3 comments, this will only delete comments matching Bullying 3, and comments matching Bullying 2 and Bullying 1 will not be removed.

For each category, you can set a severity level and determine what happens to comments matching that severity and above. Comments matching the severity level for that category can be automatically Deleted, automatically set to Pending for Moderator review, or automatically marked as Spam. Additional instructions on setting up moderation rules may be found [here](#moderation-rules).

#### **Monitoring and reviewing comments**

Comments will show the category and severity ratings regardless of whether they are removed by an automated rule. These gradings will appear in tags on the comments in the comments stream of your Moderation Panel. Viewing the tags on your existing comments can help give a sense of what automated rules to put in place.

Additionally, moderation filters can be applied to view only comments matching one or more of the content categories. If you’d like to view or moderate only comments that contain Bullying comments, you can select the Bullying filter here:

Once the Rules have been active and running, clicking into each rule will give a report of the amount of comments that have had action taken on them by the given rule, as well as a link to view your past comments that fit the category. This can help you determine which rules are functioning as desired, and which should be adjusted.

​

### Dealing with spam {#dealing-with-spam}

Disqus uses its own anti-spam software to smartly combat comment spam. It was designed to learn over time and becomes increasingly accurate with your moderation activity. Spammers are a prolific bunch, and thus there is always some chance that newer techniques may initially get past the anti-spam. These are tips on how to reduce or altogether eliminate spam.

*Learn about the* *latest anti-spam improvements* *Disqus has made for millions of publishers.*

Be sure to mark comments as spam if they are indeed spam. However, be careful not to mark non-spam comments, even if the comments are abusive, offensive, or just plain disagreeable. Marking non-spam comments as spam pollutes the data that Disqus collects and results in less accurate anti-spam detection.

You may mark comments as spam from the Moderation panel, the comment thread itself, and from email notifications.

#### Pre-moderation

You may choose to pre-moderate all comments posted on your site. All comments, spam or not, will require moderator approval before being published for others to see. Another option you may want to try is pre-moderating all links in comments. To view and change moderation settings, head to Disqus Admin > Community > Community Rules.
​
If you are on a Pro or Business subscription, you also have the option of using our "New Commenter" Pre-moderation. This will send comments to Pending for accounts that are new to your site, for a duration of your choosing. This includes brand new accounts as well as older accounts that have not posted to your site before, but your returning commenters will not be affected. This is an effective tool for catching spam and ensuring that new commenters are vetted.
​

#### Report Spam to Disqus

We appreciate reports of spam accounts that aren't getting caught by the automatic spam filter. To report a spammer, see User Flagging.

### How do I change the ownership of a Site or Organization? {#how-do-i-change-the-ownership-of-a-site-or-organization}

***IMPORTANT:*** *If your current organization has a paid package associated with it, transferring the organization ownership may remove its package features like SSO and unlimited API. Please consider handing over the user account itself instead, explained in step #2 above.*

In many cases, you will not need to fully transfer the ownership. Instead, you can decide to choose the following options:

1.  If you do not need to relinquish full control over the organization or site, simply grant the user permissions by adding the user as an organization admin or a site moderator. You will remain the organization owner.

2.  If you do not need your Disqus user account any more, simply update the username, display name, and email address to match that of the new desired user. All of this can be done in the Edit Profile view while logged-in as the organization owner. The new user can then reset the password at 0.

#### Full Ownership Transfer

Full Site and Organization ownership can only be transferred to a different user account by contacting Disqus. If you still want to transfer ownership of a site or organization, please **contact us using the blue chat button below**, while logged in as the site or organization owner and provide all of the following information:

-   registered Disqus username of the current owner

-   shortname(s) of the site(s) which you would like transferred

-   registered Disqus username of the desired new owner

-   registered email address of new desired owner

**This request must be made by the current site or organization owner, so be sure to log in first.**

#### Lost access to owner account

If you're not the site or organization owner, kindly have him/her **contact us using the blue chat button below**. Please include information indicating that you currently have permissions for the website, i.e. admin email access, consent from the previous primary moderator, etc.

### How do I Close and Open comment threads? {#how-do-i-close-and-open-comment-threads}

Closing a comment thread is used to halt the discussion while allowing the current comments to remain visible. Closing comments can be done several ways:

For your entire site you can specify how long you want comments to remain open after the thread has been created. You can find this option in your **Disqus Admin > Settings > Moderation Settings** page under **Automatic Closing**.

*Note: You can override this length for individual posts by going to the page directly and clicking the **Settings gear** to change the option.*

#### Manually closing from the thread

To start, make sure you are logged in to your Disqus account. While viewing the Disqus comment section of an article, select your name on the top-right to trigger the **Settings dropdown** and then select **Close Thread**.

This method can also be used to re-open closed threads.

#### Manually closing from the Discussions tab

To close a thread from your Discussions Editor, simply close the Lock icon into the locked position.

####   Manually closing from the admin

To close a thread in the Disqus admin:

1.  Locate a comment that belongs to the thread.

2.  Expand the comment.

3.  Click the Discussion menu.

4.  Choose Close Discussion.

#### Using the WordPress plugin

Disqus will respect the Wordpress settings with regards to disallowing commenting. If you're using Wordpress, you can remove Disqus on a per-post basis by disabling comments for each post in WordPress itself, or setting an Automatic closing timer in your Wordpress discussion settings. See WordPress' Enable and Disable Comments guide for steps.

### How to add Admins and Moderators to your Organization {#how-to-add-admins-and-moderators-to-your-organization}

Admins can be added to an organization at the Disqus Admin > Settings > Admins page after selecting your desired organization. Click the "Add admin" button and entering the person's Disqus username. **Please note that while the username is designated on the profile with an @, this symbol is not part of the username, and shouldn't be entered into the field**. Be sure to click save at the bottom of this page.

#### How to Add Moderators to Your Site

Moderators can be added to an individual site at the Disqus Admin > Moderation > Moderators page by clicking the "Add a moderator" button and entering the person's Disqus username. Be sure to click save at the bottom of this page.

Keep in mind that additional moderators *must* already be registered as Disqus users before being added to your community as moderators. If the moderator you'd like to add doesn't already have a Disqus user account, they can register with their email and password.
​
Single Sign-on (SSO) accounts can be added as moderators too by entering their SSO username (e.g., 0).
​

####  There are 3 different roles when it comes to administration and moderation on Disqus:

-   Org Owner - a founding admin of an organization of sites. This account has "can edit settings" and "can edit comments" permissions on all sites within the organization, along with the ability to delete the organizations and individual sites. To transfer ownership of individual sites, see How do I change the primary moderator?

-   Org Admin - an admin of an organization of sites. This account has "can edit settings" and "can edit comments" permissions on all sites within the organization.

-   Site Moderator - a moderator of the forum who does **not** have admin permissions for the organization or "can edit settings" permissions for the site.

The Disqus Admin provides access to different features, settings, and tools based on these user roles. More information on See the chart below to understand which Admin navigation pages each role has access to:

Learn more about about the difference between Usernames and Display Names here.

### Moderating 101 {#moderating-101}

This is a guide for navigating and using the Disqus moderation panel. For information on word filters and other Moderation Settings, please visit [this page](#moderation-settings).

#### In this Guide:

Using the Moderation Panel

-   Bulk Actions

-   Trusting and Banning Users

-   Moderate All Comments by a Single User

-   Editing Comments

-   Searching Comments

-   Moderation History

-   Keyboard Shortcuts

-   Moderating through the Embed

-   Moderate by Email

-   Closing comment threads

#### Additional Resources:

-   Guide to building community guidelines

Your moderation panel gives you a quick view of your forum's comments, and their status (Approved, Pending, etc.). There are many detailed functions that are easily accessible with a couple of clicks as well.

#### Approving, deleting, and marking comments as spam

In the moderation panel, you can perform bulk actions by checking multiple comments, or you can moderate them individually in the expanded comment view.

-   **Approve**: Approves a comment to be shown publicly on your site.

-   **Delete**: Removes a comment from your site.

-   **Spam**: Removes a comment from your site and marks it as spam. Read more about Dealing with Spam.

#### Priority Sort Order

When you only have a few minutes to moderate, the new “Priority” sort option does all the heavy lifting for you. This option intelligently surfaces comments that might need your immediate attention using a special calculation of flags, downvotes, restricted words, reputation, and guest account status. This maximizes your ability to keep your community healthy and civil. Priority sort can be found next to the existing “newest” and “oldest” options within the sort menu in the top right navigation.
​

#### Pending reasons and filters

No more guess work. New filters have been added to the top navigation of the moderation panel which allow you to quickly view groups of comments based on the reason they’re pending your approval. This allows you to divide and conquer more efficiently and use signals like reputation to focus your effort. Here are the new filters:

-   No Issue

-   Guest comment

-   Contains a link

-   Low reputation

-   Restricted word

-   Flagged

-   Toxic (What's this?)

#### Bulk Actions

#### Individual Moderation

#### Expanded comment view

Click the comment to bring up more options.

#### Trusting and Banning Users

In the expanded comment view, click the "Trust User" or "Ban User" options.

Read more about Using the "Ban User" and "Trust User" Controls.

#### Timeouts and Shadow Banning

After clicking "Ban User", you will be presented with additional options to either remove the user temporarily or allow them to continue to post comments that nobody will be able to see.

Read more about timeouts and shadow banning!

#### Moderate All Comments by a Single User

The "More Info" button allows you to view other comments made by that commenter — you can drill down by their username, email address, or IP address, as well as view only that thread's comments.

This will only search for that single account, and if the person has used multiple accounts or email addresses this won't find the others.

#### Editing Comments

As a moderator you can edit any comment if it contains undesirable information.

#### Searching Comments

You can use a variety of search commands to find certain comments:

-   **email**:user@example.com

-   **user**:username

-   **thread**:0000000 (Enter the thread ID number)

-   **id**:0000000 (Enter the comment ID number)

After clicking on a comment, can view the user's recent comments for context.

#### Can't find a missing comment?

-   Try viewing the Spam or Deleted filters.

-   Make sure there's no space between the search command and what you're searching for.

-   When searching for a comment, pick the most unique word from it to search.

-   Search a variety of words in the comment, or simplify what you're entering in the search field.

-   Take note if the person is using a third-party login, or if they're using multiple accounts and/or email addresses.

-   If you're searching by username or email for a third-party login account — users who log in to Disqus via Twitter, Facebook, or Google — the account won't have an email address or a standard Disqus username. The username will be similar to facebook-12345678. You can find the username by going to any comment, expanding it, and clicking the More Info button (the silhouette of a head and shoulders).

#### Moderation History

The Moderation panel will also show past moderation actions for comments, in addition to moderation overview for other comments made by that user

To view this information, select any comment to toggle the sidebar panel. Then select the **History** tab to view a log of all the moderation actions for the commenter.

Currently, we show the following actions: comment approval, deletion, marked as spam, and featured.

#### Keyboard Shortcuts

Moderate faster in the Moderation Panel using keyboard shortcuts. To display the full list of keyboard shortcuts in the Moderation Panel, type '**?**' (without the quotes).

#### Moderate through the Embed

You can moderate comments directly on the thread's page while logged in as a moderator. After mousing over a comment, click the comment menu dropdown and then 'Moderate' link.
​

#### Moderate by Email

You can moderate your comments via their email notifications with three simple commands: Approve, Delete, Spam. Remember to reply above the notification text, respond only with the appropriate single command word, and to not include any email signature.

You can also post a comment reply to a new comment notification by responding directly to the email.

#### Closing comment threads

Closing a comment thread is used to halt the discussion while allowing the current comments to remain visible. Closing comments can be done several ways.

### Moderation Rules {#moderation-rules}

*Full Access to Moderation Rules is currently available with a Pro or Business subscription. If you would like to subscribe to a Pro or Business plan, you can do so in your Subscription & Billing settings.*

Moderation rules allow you to define your automated moderation practices by assigning certain actions to filters. This feature allows your team to work more efficiently by freeing up your time to focus on non-repetitive moderation tasks.

Start by clicking the button labeled **+ Add Rule** to add your first rule. Below is a sample rule that you can configure:

You can choose any combination of the following:

**If comment matches**:

-   Contains link

-   Flagged at least 5 times

**If user matches:**

-   Profile flagged at least 5 times

**Then:**

-   Send to Pending

-   Delete

-   Mark as spam

To enable the rule, click the toggle button from OFF to On. Note that individual rules can be assigned different priority by using the up and down arrows, found in the left section of each rule. If a comment matches multiple rules, the top most rule will take the highest priority.
​
Click **Save** to save your current set of moderation rules.

Comments affected by a moderation rule will be marked with a reason like: **In Pending because Toxic**

Moderation Rules may be set up from your Moderation Settings page.

### Moderation Settings {#moderation-settings}

This article explains how to configure your moderation tools, to ensure that unwanted content is removed or sent to pending for moderator review. You may access all of the tools in the Settings -> Moderation page of your moderation panel. For instructions on moderating comments and using the moderation panel itself, please go [here](#moderating-101).
​
*In addition to what's listed here, Pro and Business publishers will have access to additional AI filtering and tooling covered in [Advanced Moderation](#advanced-moderation).*

#### Setting up Moderation Profiles

The first step in setting up your moderation tools will be to choose a Moderation Profile. This will configure your rules and apply presets for general content safety, which you can then customize as needed.
​
There are two Moderation Profiles:

-   **Balanced** *(applied by default)*

    -   Comments containing Restricted Words are deleted automatically

    -   Comments labelled Toxic will be sent to Pending for moderator review

    -   Guest Commenting is allowed

    -   Comments Flagged 5 times will be sent to Pending for moderator review

-   **Strict** *(recommended for publishers who want to be considered Brand Safe by advertisers)*

    -   Comments containing Restricted Words are deleted automatically

    -   Comments labelled Toxic are deleted automatically

    -   Comments containing links are sent to Pending for moderator review

    -   Comments Flagged 3 times will be sent to Pending for moderator review

    -   Images and Videos not allowed in comments

    -   Threads are automatically closed to new comments after 30 days
        ​

#### Adding Restricted Words

Any time a comment or name contains a word you've specified in this filter, it will be sent to the Pending queue. We pre-populate this list with a recommended set of words, which can be edited at any time. You may access the list from the left-side menu in your main moderation panel page, or may visit it directly here. Up to 2,000 words may be added to the list.
​
You may use wildcards by entering a 0, but be careful where you use it; for example,

s.\*ck will match suck, but also sock and stack.

In addition to being set as Pending, Comments containing a Restricted Word will highlight the trigger word with yellow background and will contain a red "Restricted Word" tag.

*Note: If a restricted word appears in the display name of an account, that account’s comments will remain pending until explicitly approved. Additionally, words are stripped of punctuation when measured against the Restricted Word filter, so adding “hell” will also set all comments containing “he’ll” to Pending until approved by a site moderator.*

The restricted words list will also work for blocking certain HTML code.

#### Setting Moderation Rules

All sites are able to add automated moderation rules based on comment [toxicity](#toxic-mod-filter) and their Restricted Words list. Sites on the Pro and Business plans will have access to additional moderation rules to target objectionable content more specifically by category and severity. More information on our Advanced Moderation Rules may be found [here](#advanced-moderation).
​
As an example, a rule could be set up so that comments containing Restricted Words will be deleted automatically, without any action from moderators. Clicking on each rule will expose analytics, showing how many comments from the last 30 days fit the conditions in your rule, and how you moderated them. This can be used to forecast how effective each rule will be.

More information on setting up Moderation Rules may be found [here](#moderation-rules).

#### Pre-moderation

Pre-moderation options can be found in your forum's Disqus Admin > Moderation Settings page. Setting Pre-moderation to "All" will set all incoming comments to Pending and will prevent them from appearing publicly until they've been explicitly approved by a moderator.

Pro and Business publishers will also have the option of Pre-moderating **"New Commenters"**, or commenters who have not posted to their site before. This can help catch Spam and unwanted trolling, without interrupting the posting ability of longstanding site commenters. When using this feature, you'll also set the number of days that you want New Commenters to be sandboxed, before they are subject to the same moderation rules as the rest of your users.
​
You may also choose to pre-moderate comments containing links by enabling the setting for Links in Comments.
​
With **Thread-level Pre-moderation**, sites can choose specific discussion threads (article pages) to apply Pre-moderation. This added flexibility allows you to apply Pre-moderation to specific threads without pre-moderating comments across the rest of their sites.

You can enable pre-moderation for an individual thread either via the Moderation Panel or in the dropdown menu in the comment embed.

**Guest Commenting**

Guest commenting allows users to leave comments without creating a Disqus account. Guest commenters are required to provide an email address for notifications and moderation purposes. However, the email address input will not be verified by Disqus.

As Guests are posted without a verified account, all Guest Comments will require explicit Moderator approval, regardless of whether Pre-Moderation has been enabled.

*Note: Registered users must verify their email address prior to posting a comment. More information on Guest commenting may be found [here](#guest-commenting).*

#### **Links in Comments**

This setting will allow you to decide if comments containing links (including any posted images and videos) will require moderator approval. When this feature is enabled, all comments containing links or images will remain in the Pending folder until they are approved by a site Moderator.

#### Flagged Comment Moderation

You can choose to email all moderators when a comment is flagged. You can also specify how many unique flags are needed to hide a comment (if any). More information on Flagging may be found here.

*Note: We don't count duplicate flaggings based on the user and IP address.*

#### Images and Videos

You can choose whether or not Images and Videos are enabled for your forum. If enabled, thumbnails for images and YouTube links will display within comments.

Read more at Adding images or videos to comments.

#### **Gif Picker**

The Gif Picker can also be enabled or disabled from the Moderation settings page. More information on the Gif Picker may be found in the Gif section [here](#adding-images-and-videos).

#### Automatic closing of comment threads

You can specify globally how long you want a comment thread to remain open after it's created. Entering 0 will disable the automatic closing process. This can be overridden for particular threads using the Settings panel on the embed.

Read more at How do I close comment threads?

### Reactions {#reactions}

Reactions allow your readers to give an initial response to an article or question, and see how the rest of the community feels about it before getting into the details in the comments.

If you’re the moderator of your Disqus forum, you can enable reactions from your Disqus admin in your Admin > Settings > Reactions.

#### Configure images and text

Reactions can be configured with emoji images and text, or text only. Select your desired emoji from the “Image” dropdown menus for each Reaction and edit the text in the “Description” field.

To add or remove a Reaction, use the “Remove” or “+ Add reaction” buttons. Keep in mind that you can save a minimum of 2 Reactions and a maximum of 6.

#### Preview before saving

Use the “Show Preview” button to expand a mock discussion that includes the customization you have made. This is how the Reactions feature will appear on your website. Immediately after clicking “Save”. You can make further customizations to the feature if you are not happy with the current preview.

After clicking “Save”, you will be presented with a popup that explains how the changes you have made to the Reactions feature will only apply to new articles. If you would like to apply your changes to older **articles that previously had Reactions enabled**, toggle the checkbox shown below.

#### Custom Reactions

\*\***Access to Custom Reactions is available as part of our Pro and Business subscription plans. You may subscribe to Pro from your Subscription page, or can request information on the Business plan from one of our account managershere**\*\*

#### Setting Up Custom Reactions

Custom Reactions may be set up from your admin panel, and will appear below the “Configure and enable Reactions” section of your Reactions Settings page.

New Custom Reactions may be created by uploading image files from your computer. You may use images that are in the JPEG, PNG, or GIF formats (including animated GIFs). When choosing images for your Custom Reactions, please ensure that the files are smaller than 5mb.

To upload a custom reaction Image, simply click the Plus button, select your images, and hit open. After these have been successfully uploaded, they may be selected in the Reactions dropdown.

When selecting images to use, please note that all uploaded images will be resized to a 1:1 ratio.

When uploading Custom Reaction images, you may select multiple gifs/images to be uploaded at the same time. Images cannot be added twice, as our system will automatically prevent duplicates.

Once uploaded, you may select a new custom image as you would any regular Reaction image in the Reactions drop-down.

#### Deleting Customized Reactions

To delete custom images, select the images you want to delete, and then click the delete button. This will remove images that you have uploaded.

Please note that you are only able to delete custom images that are not currently being used by a Reaction, and images will need to be removed from the Reactions template (saved or unsaved) prior to removal.

Deleted images will continue to show in the Reactions section for threads created during their use, even if one of these Reaction images is later deleted.

#### Closing Reactions for specific discussions

**Note**: When Reactions are enabled, **they will appear in new article discussions only**. Existing article threads will not display the Reactions widget.

To disable Reactions on individual discussions, visit the article on your website you would like to edit and toggle the “Remove Reactions” dropdown option found in the top-right dropdown the embed.

####  Mobile

The Reactions feature is optimized to look great on mobile so that readers of your site can share what they think from any device.

#### FAQ

##### Does a reader have to be logged in to Disqus in to share their Reaction?

No, logged out readers can still click or tap on the Reaction responses and will be counted alongside logged in commenters. This helps increase engagement on your site from logged out readers.

##### **Can I enable reactions to appear in all existing articles when turning on Reactions?**

No, Reactions will only appear on newly published articles after you enable Reactions.

**CUSTOM REACTIONS IMAGE SPECIFICATIONS:**

-   size restriction is up to 5mb

-   file format is jpeg, png, gif

-   the aspect ratio is 1:1 (a square)

-   End image will appear as 24px by 24px

### Site Moderators {#site-moderators}

Site Moderators are the people in charge of managing and maintaining the commenting community for a given website or blog. They are the first person to contact if you're having trouble with moderation or abuse on a given site.

Commenting communities vary greatly from site to site – certain language or behavior that is considered acceptable on one site can be considered harmful on another. Site Moderators can help keep their unique community on track by communicating to new or unruly users when they're out-of-line and moderating comments when needed.

#### As Disqus doesn't provide site moderation, nor do we know the ins and outs of each community, only the Site Moderator for a given site can provide reasons for moderation actions within that community.

#### Site Moderators are the decision makers for the following:

-   Deleting comments, approving Pending comments, approving comments marked as Spam

-   Handling disputes among commenters on the site

-   Blocking or unblocking an account from posting on the site

#### Contacting a Moderator

You can contact the site moderator through a Contact page, commenting policy, or community guidelines. Site moderators are not Disqus staff and should be contacted at the site in question, as opposed to through Disqus. Please note that sites are not required to provide contact information and, in cases where they don't, Disqus will not be able to put you in contact with that site's moderator.

#### Related:

-   How to Report Abuse

-   [Adding Moderators and Admins](#how-to-add-admins-and-moderators-to-your-organization)

### Toxic Mod Filter {#toxic-mod-filter}

Toxic comments disrupt communities, drive users away, and strain moderation efforts. The Toxicity Mod Filter empowers moderators to prioritize toxic content for moderation in order to lower their negative impact on the community and decreases the reliance on users flagging comments.
​

The toxicity filter utilizes natural language processing and machine learning to analyze and identify comments likely to be toxic. We’ve integrated our moderation system with Google’s Perspective API to deliver this capability.
​
More information and insights on this technology can be found in this blog post.

#### What are toxic comments?

Toxic comments are defined as having at least two of the following properties:

-   Abuse: The main goal of the comment is to abuse or offend an individual or group of individuals.

-   Trolling: The main goal of the comment is to garner a negative response.

-   Lack of contribution: The comment does not actually contribute to the conversation.

-   Reasonable reader property: Reading the comment would likely cause a reasonable person to leave a discussion thread.

This two-property guideline should help prevent comments like “haha”, that don’t add to the conversation as well as comments that provide opposing viewpoints from being flagged as toxic.

Filter for toxic comments using your Moderation Panel and decide on the moderation action to perform.

#### Frequently Asked Questions:

**Does this auto-moderate comments on my site?**
No. Toxicity is a tag / label that DIsqus provides for publishers in the moderation panel. From their moderation panels, publishers can see which comments are labelled as toxic. Publishers can also sort by “toxic” to see all toxic comments. Comments that are labelled as “toxic” are not moderated, or pre-moderated. At this time, the toxicity label is just another piece of information to help publishers moderate more effectively.
​
**If a comment is not Toxic, can I remove the “Toxic” tag?**
No. While this ability currently isn’t available, we are looking into methods of improving its use in the future.
​
**Will commenters see this?**
No. Only site moderators can see this within the Moderation Panel.
​
**Will you show a numeric score for toxic comments?**
No. This only tags a comment with the label “Toxic” within your Moderation Panel.
​
**What happens when a user edits their comments?**
The filter re-checks their comment.
​
**What languages does the filter support?**
Currently, only English. We hope to expand support for non-English languages in the near future.
​

### User Reporting {#user-reporting}

If you've encountered a user that is breaking the Basic Rules of Disqus, use the "Report User" button to flag the account for Disqus attention.
​
​

1.  Go to the Disqus profile you want to report by clicking their username or avatar

2.  Click the dropdown menu adjacent to the user's name and select 'Report User'

3.  Complete the report following the instructions

#### The person I want to report isn't breaking the Basic Rules of Disqus but is REALLY annoying, can I flag their account?

No. Flag a user only if they're breaking the Basic Rules of Disqus. If they're really bugging you, you can block the user. Or, if they are violating the community guidelines of the site where you both comment, contact the site moderator.

#### The person I want to report isn't breaking the Basic Rules of Disqus but is breaking the community guidelines of the site where we both comment, can I flag their account?

No. Flag a user only if they're breaking the Basic Rules of Disqus. Disqus doesn't moderate the comments on sites that use Disqus, so you'll need to report the user directly to the site moderator. We also recommend flagging bad comments from the user; flagging comments raises those comments for site moderator attention.

#### What's the difference between flagging a comment and flagging a user?

Flagging a comment raises that one comment to the attention of the site moderator. Flagging a user raises the entire account to the attention of Disqus.

-   When flagging a comment, consider the community guidelines of the site where the comment was posted.

-   When flagging a user, consider the Basic Rules of Disqus.

### User Reputation {#user-reputation}

User reputation enables you to make smarter decisions while moderating your community. Use it when evaluating your users to determine who should be added to Trusted Users, or who may be a bad actor.

#### How do I activate user reputation?

No activation is necessary. Reputation will appear in the comments view of the Moderation panel for High and Low reputation users.

#### Is reputation unique to my site?

Reputation is platform-wide across the entire Disqus network. A user's reputation is the same across all Disqus-powered sites.

#### Is reputation public?

Reputation cannot be customized and is only visible in the moderation interface of a Disqus admin; it is not publicly visible.

#### What does it look like?

There are three reputation tiers:

-   High: These are active and up-voted users.

-   Average: This is where everyone starts.

-   Low: These users likely have many flagged and/or deleted comments.

You can also see how long your users have been around and how active they are.

#### Why are accounts marked as low reputation?

Accounts may be marked as having low reputation for multiple reasons, for example:

-   Posting multiple comments subsequently deleted by a moderator

-   Posting multiple comments subsequently marked as spam

-   Posting multiple comments subsequently flagged by other users

If you want to filter comments by Low Rep users, you can do that.
​
​

Note that being marked as having low reputation does **not** affect how that account's comments are treated by our system. For example, comments posted by an account with low reputation are by default no more likely to be marked as spam or pre-moderated than comments posted by an account with average or high reputation. Reputation is only used as a visual indicator within the Disqus moderation interface.

#### API usage

Reputation is available via a number of posts and users API endpoints.

### Using the "Ban User" and "Trust User" controls {#using-the-ban-user-and-trust-user-controls}

Use the "Ban User" control to block spammers, offensive commenters and/or those who violate your commenting policy. This will block the user from posting future comments and give you the option to retroactively delete their comments from the past 30 days.

Adding users to your Trusted Users list will ensure that their comments are not caught as Spam, though they'll still be held for review if they violate any of your moderation rules or contain a word from your Restricted Words list. This is used for trusted commenters, such as community regulars or website staff who aren't listed as moderators.
​
While banning users is immediate, retroactive deletions may take up to 24 hours to complete.

#### On the Embed (Ban User only)

Click the three dots in the upper right corner of the comment, and then **Ban User**.
​

#### Moderation Panel

Expand the comment, then click *Trust User* or *Ban User*.
​

If you select *Ban User*, you’ll see options to either remove the user permanently, Shadow Ban, or give them a Timeout.
​

\*\Shadow Ban*** *and* ***Timeouts*** *are Pro level features. Check the links above for more information!*

#### Community > Banned Users

If you're moderator with 'Settings' permissions, you can add a commenter's **username**, **email address**, and/or **IP address** directly to your forum's master Banned User or Trusted User list.

#### Banning/trusting email domains

Email domains like 0 can also be banned/trusted at Settings > Access Control by selecting Domain in the Add item > Type dropdown.

*Keep in mind*: Banning an email domain will block anyone using that domain from posting comments to your site, whether intentional or not. For example, banning the 0domain will block all users with an email address ending in 1 from posting comments to your site. When possible, we recommend blocking individual users before blocking domains.

#### Finding Banned/Trusted Items

You can search your Trusted/Banned Users lists at Disqus Admin > Community > Banned Users.

#### Can't find a blocked commenter?

-   If possible, find the blocked commenter's IP address from a previous comment and search your Banned Users list for it.*Keep in mind*: Banning an IP will block anyone using that IP from posting comments to your site, whether intentional or not.

-   Guest commenters can only be banned using an email address or IP address since they don't have a username.

-   Non-Disqus commenters have a username similar to *twitter-123456789*. You can find this username by expanding a user's comment in the **Moderation Panel** and then clicking the user drop-down.

#### Check your banned IP addresses

#### Search by email or IP for guest commenters

#### Looking for a third-party service username?

-   Moderating your Community

## Other Integrations {#cat-other-integrations}

### Sitefinity Installation Instructions {#sitefinity-installation-instructions}

You may browse the instructions, but you will need to register your site before installing.

#### To install Disqus on Sitefinity:

1\. Visit this link to download Random Site Controls: 0

2\. Copy the RandomSiteControls.dll to the bin folder of your project.

3\. Add in the Toolbox entries from the included ToolboxesConfig.config byeither copying the entires into your own, or add them with theSitefinity backend Settings UI at Settings > Advanced > Toolboxes > Toolboxes

4\. Add in the virtual path entry as seen in"VirtualPathSettingsConfig.config"...again by copying the entry into yourown, or use the backend UI located atSettings > Advanced > VirtualPathSettings > Virtual paths > Create New

5\. Add this to the global.asax to register the section.

    protected void Application_Start(object sender, EventArgs e)
    {
    Telerik.Sitefinity.Abstractions.Bootstrapper.Initialized += new EventHandler(Bootstrapper_Initialized);
    }

    void Bootstrapper_Initialized(object sender,Telerik.Sitefinity.Data.ExecutedEventArgs e)
    {
    Telerik.Sitefinity.Configuration.Config.RegisterSection(
    );
    }

6\. Recompile your project.

7\. Controls will now be available to drag\\drop from the page designer under the Disqus section, and global options can be set in the backend under Advanced Settings > SitefinitySteve > Disqus

### Tumblr Manual Installation Instructions {#tumblr-manual-installation-instructions}

For displaying Disqus on your site, typically on individual article or post pages.

1\. In Tumblr, visit Settings > Edit theme
2. Select "Edit HTML".
3. Copy and paste the following code anywhere between the 0 and 1 tags: 2
4. Copy and paste the following code immediately after 0:

    {block:IfDisqusShortname}

        /* * * CONFIGURATION VARIABLES: EDIT BEFORE PASTING INTO YOUR WEBPAGE * * */
        var disqus_shortname = '{text:Disqus Shortname}'; // Required - Enter shortname in Tumblr Theme Options
        var disqus_url = '{Permalink}';     /* * * DON'T EDIT BELOW THIS LINE * * */
        (function() {
            var dsq = document.createElement('script'); dsq.type = 'text/javascript'; dsq.async = true;
            dsq.src = '//' + disqus_shortname + '.disqus.com/embed.js';
            (document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(dsq);
        })();

    Please enable JavaScript to view the comments powered by Disqus.blog comments powered by Disqus
    {/block:IfDisqusShortname}

5\. Select "Update Preview" and then "Save"
6. Go back to "Edit Theme"
7. Enter your Disqus site shortname in the "Disqus Shortname" field. **If you don't, Disqus won't work.**

For displaying comment counts on your site's home page.

1.  In Tumblr, visit Settings > Your Site > Customize

2.  Select "Edit HTML".

3.  Copy and paste the following code immediately before 0:

    {block:IndexPage}Comments

        var disqus_shortname = '{text:Disqus Shortname}';    (function () {
            var s = document.createElement('script'); s.async = true;
            s.type = 'text/javascript';
            s.src = '//' + disqus_shortname + '.disqus.com/count.js';
            (document.getElementsByTagName('HEAD')[0] || document.getElementsByTagName('BODY')[0]).appendChild(s);
        }());

    {/block:IndexPage}

### Use Zapier to connect other apps {#use-zapier-to-connect-other-apps}

Use Zapier to connect Disqus to over 1,000 other apps. Zapier is a workflow automation tool that helps you managing new comments and taking action on autopilot. The Disqus-Zapier integration allows you to send Disqus data to hundreds of other apps. To see a few examples, check out our integrations page with Zapier.

For example, you can send New Comments in Disqus to a Slack channel to alert your team, automatically send an email, or even trigger an SMS message. Visit How To Use Slack To Supercharge Your Disqus Moderation to learn more.

You'll connect Zapier to Disqus while setting up your first Zap.

"Zaps" (which are Zapier’s name for app workflows) can be created quickly and easily. Use them to send information from Disqus to other other apps like Slack , Zendesk, Mailchimp, Gmail , and Facebook Pages automatically.

First, you'll need a Zapier account. Sign up for a free Zapier account if you don't already have one.

Choose from popular Zaps here:

See more Disqus integrations powered by Zapier

You can also create custom Zaps by visiting the Zapier-Disqus Integration Page and choosing Make A Zap.

#### Things to keep in mind

-   When selecting a Forum in the Trigger step, you can enter a Custom value if your site’s shortname does not appear when selecting from the dropdown menu. The Forum should be the site where Disqus is installed.

-   In order to use the **New Email Subscriber** trigger, you will need to have a Disqus Pro subscription. If you would like to subscribe to **Pro**, you can do so in your subscription settings at **AdminSettings** > **Subscription and Billing**.

-   Zapier offers a free plan that you can use to set up 5 Zaps and run 100 tasks per month. Zaps on the free plan run automatically every 15 minutes and is only counted as a task if it was triggered successfully (e.g. a new email subscriber was found). More info on paid plans can be found on their Pricing page.

## Terms and Policies {#cat-terms-and-policies}

### Abusive Behavior Policy {#abusive-behavior-policy}

Engaging in targeted abuse or harassment on Disqus is a violation of the Basic Rules and Terms of Service. We want to foster a positive and diverse community for rich discussions. It’s our responsibility to cultivate an environment for communities and discussions to thrive and for users to feel safe when participating. To that end, we’ve established this policy to help us evaluate abusive behavior in its various forms.
​
We do not condone and can take action on the following types of abusive behavior including but not limited to:

-   targeted harassment or encouraging others to do so

-   threat of violence or inciting it

-   self-harm or suicide

-   impersonating someone in a misleading or deceptive manner

-   posting personally identifiable information

-   improper use or breaking Disqus in such a way that it negatively hurts the experience for others

-   posting illegal content such as copyrighted material or child pornography

All types of content is tolerated as long as it does not violate the Disqus Basic Rules and Terms of Service. Additionally, we do not mediate content or intervene in disputes between users.
​
As Disqus doesn't provide moderation, nor do we know the ins and outs of each community, the site moderators are the people you should contact to report abusive behavior that isn’t covered by the Basic Rules. Learn more about how moderation works on Disqus.

We review flagged user reports and will act to enforce the Basic Rules and Terms of Service when we find they have been broken. Actions vary depending on the severity of the violation(s) and may include but not limited to:

-   removal of content (e.g. comments, discussions)

-   warning the user of the violation

-   resetting the user's profile to default

-   global banning accounts or communities

#### Reporting abusive behavior to Disqus

If you've become the target of abusive activity or are seeing it take place in a discussion or profile powered by Disqus, there are a few things you can do to raise awareness of the issue to the right people, and hopefully, find a resolution. For instructions about how to report abusive behavior, read this article.

### Ads-Free Subscription & Payments FAQ {#ads-free-subscription-payments-faq}

#### How do I start a Subscription?

To begin a subscription, you may navigate to your Subscription and Billing page, add your billing information, and then select the plan you'd like to subscribe to. Billing information must be added before you can initiate a trial of your subscription.
​
If you are looking to remove ads, this can be completed from each site's Ads Settings page after you've subscribed.

#### How long does the trial period last?

The trial period lasts for the first 30 days after you subscribe to a plan. You can see how much time you have left in your trial from the Subscription and Billing tab on your Admin Settings page.
​
You are granted 1 trial per plan, but if you click between plans before your 30 day trial is complete, you will forfeit the rest of your trial period for that plan. If you cancel at any point during your trial period, you will not be billed, and will be reverted to the Basic plan accordingly.

#### What happens at the end of the trial?

At the end of the 30-day trial period, billing will initiate and you will move from your trial period into your subscription. Subsequent billing will occur monthly or yearly, depending on what you've selected.

#### When will my subscription start?

Your subscription will start with the first bill you receive, 30 days after you select your plan and initiate your trial. All functionality provided in your selected plan will be available to you within your initial trial period.

#### How does Disqus measure traffic for my site?

Disqus measures traffic by the number of times the Disqus comment embed loads on your site, which happens on every page load where Disqus is enabled. Page loads that do not have the Disqus embed, such as homepages, do not count towards the traffic measured by Disqus.

#### Do you offer an annual pricing option?

Yes, annual pricing can save you 10%. Just visit Admin > Subscription & Billing to select "Switch to Yearly billing" for an active plan or toggle the "Billed Yearly" button when selecting a new plan subscription.

#### **How can I cancel my subscription?**

To cancel your subscription, please email our team at: cancellation@disqus.com. They will be able to assist with any issues and walk you through the cancellation procedure.

#### Ad configuration and settings

#### What configuration options are available with ads?

For publishers on our Basic Plan, an ad unit in the above-the-comments position, is included. Publishers can reach out to our team for questions specific to ads on their site. Additional positions and configurations are available for larger publishers that are part of our revenue share program.

#### Why were ads activated on my site?

The free version of Disqus is supported by advertising. Commercial sites on this version of Disqus will likely have some advertising. We email publishers before activating advertising. If no one at your organization received these emails, please reach out to us so that we can verify your contact information.

#### Do I earn money from ads while using the Basic package?

The Basic package does allow a site to use the Disqus service for free, but does not necessarily mean that they earn revenue from the ads being shown. You can visit this page for more information on earning revenue from your Disqus Ads.

### Amendment to Disqus Terms of Service Applicable to U.S. Federal Government Users {#amendment-to-disqus-terms-of-service-applicable-to-u-s-feder}

This Amendment to the standard Terms of Service (“Terms”) of Disqus, Inc. (“Disqus” or "Company") posted at 0 is an agreement between the Company and the U.S. Government and it applies to the use of the Disqus Service by U.S. Government entities (“You” or the “Agency”).
​
The reason for this Amendment is that the Agency must follow federal laws, regulations, rules, and practices when entering into a binding agreement with a provider such as Disqus. This Amendment allows Agencies to use the Disqus Service under federal-compatible terms that respect the Agency’s legal status, its public mission, and other circumstances unique to the U.S. Government.
​
A. ***Government entity***: “You,” “Your,” and “User” within the Terms shall mean the Agency itself and shall not apply to, nor bind (i) the individual(s) who use the Services on Agency's behalf, or (ii) any individual users who happen to be employed by or otherwise associated with the Agency.
​
B. ***Public purpose***: Agency will use the Service solely in furtherance of Agency's public purpose. Any provision in the Terms requiring that use of the Service be for private, personal and/or non-commercial purposes is waived.
​
C. ***Agency content serving the public***: Company will allow Agency's distribution or other publication via the Service of material that may contain or constitute promotions, advertisements or solicitations for goods or services, so long as the material relates to the Agency's mission.
​
D. ***Advertisements***: Neither Disqus nor the Agency wants the Government’s Content to be confused with other parties’ content. To minimize that risk, Disqus agrees not to serve or display any commercial advertisements or solicitations on any part of the website displaying content uploaded by or under the control of the Agency. This exclusion will not extend to house ads that Company may serve on such pages in a non-intrusive manner. This exclusion shall also not extend to windows or interfaces that are part of Company’s Service but not specific to Agency’s site, such as global commenter profiles.
​
E. ***No endorsement***: Disqus agrees that Your name, seals, trademarks, logos, service marks, trade names, and the fact that You use the Service, shall not be used by Company in such a manner as to state or imply that Disqus’ products or services are endorsed, sponsored or recommended by You or by any other element of the Federal Government, or are considered by You or the Federal Government to be superior to any other products or services. Except for pages whose design and content is under the control of the Agency, or for links to or promotion of such pages, Company agrees not to display any Agency or government seals, trademarks, logos, service marks, and trade names on the Company's homepage, elsewhere on the Service, or in Company advertisements and promotions, unless permission to do has been granted by the Agency or by other relevant Federal Government authority. Company may list the Agency's name in a publicly available customer list so long as the name is not displayed in a more prominent fashion than that of any other third party name.
​
F. ***Indemnification, Liability, Damages, Arbitration***: Any provisions in the Terms related to indemnification are waived and shall not apply except to the extent expressly authorized by federal law. Liability and damages for any breach of the Terms as modified by this Amendment, or any claim arising from the Terms as modified by this Amendment, shall be determined under the Federal Tort Claims Act or other governing federal authority. Any mandatory arbitration, mediation or similar dispute resolution provisions in the Terms are waived and shall not apply unless expressly agreed to by the Agency. Federal statute of limitations provisions shall apply to any breach or claim.
​
G. ***Governing law***: The Terms as modified by this Amendment shall be governed by and interpreted and enforced in accordance with the laws of the United States of America without reference to conflict of laws. To the extent permitted by federal law, the laws of the State of California (excluding California’s choice of law rules) will apply in the absence of applicable federal law.
​
H. ***Changes to standard Terms***: Language in the Terms reserving to Company the right to change the Terms without notice at any time is amended to grant You at least three days advance notice of any material change to the Terms. Company will send this notice to the email address You designate at the time You sign up for an account. You agree to notify Company of any change in Your notification email address during the life of the account.
​
I. ***Access and use***: Company acknowledges that the Agency's use of the Service may energize significant citizen engagement and otherwise become important to the Agency's mission. Language in the Terms allowing Company to terminate service or close the Agency's account at any time, for any reason or no reason, is modified to reflect the parties' agreement that Company may unilaterally terminate service and/or terminate Agency's account only for breach of Agency’s obligations under the Terms or Agency's material failure to comply with the instructions and guidelines posted by Disqus, or if Company ceases to operate its Service generally. Company will provide Agency with a reasonable opportunity to cure any breach or failure on Agency's part.
​
J. ***Provision on crawlers***: Any provision in the Terms prohibiting "crawl," "spider" or similar processes is amended to allow the Agency to apply such tools solely to its pages and Content, and solely to fulfill Agency's obligations under the Federal Records Act or other applicable federal law or regulation.
​
K. ***Ownership of names***: Any provision in the Terms related to Company's ownership of and right to change Your selected User name(s), User ID(s), domain name(s), channel name(s), and group name(s), are modified to reasonably accommodate Agency's proprietary, practical, and/or operational interest in its own publicly-recognized name and the names of Agency programs.
​
L. ***Modifications of Agency content***: Any right the Company reserves in the Terms to modify or adapt Agency Content is limited to technical actions necessary to index, format and display that Content. The right to modify or adapt does not include the right to substantively edit or otherwise alter the meaning of the Content. In the event Agency finds that Agency Content has been modified in a manner that alters its meaning, Agency may contact Company and together the parties will work in good faith to resolve the matter. Notwithstanding the foregoing, nothing in this Amendment shall result in an expansion of Agency's rights as a United States Government entity under the Copyright Act of 1976 (17 U.S.C. §§101 et sec.), specifically including Section 105 of the Act.
​
M. ***Limitation of liability***: Nothing in the Limitation of Liability clauses or elsewhere in the Terms in any way grants Company a waiver from, release of, or limitation of liability pertaining to, any past, current or future violation of federal law.
​
N. ***Uploading, deleting***: You are not obligated to place any User Content on the Service, and You reserve the right to remove all or any portion of Your Content at Your sole discretion.
​
O. ***No business relationship created***: Disqus and the Agency are independent entities and nothing in the Terms as modified by this Amendment creates a partnership, joint venture, agency, or employer/employee relationship.
​
P. ***No-cost agreement***: Nothing in the Terms as modified by this Amendment obligates the Agency to expend appropriations or incur financial obligations. None of the obligations arising from the Terms as modified by this Amendment are contingent upon the payment of fees by one party to the other.
​
Q. ***Paid Services and Agency Obligation***: This Amendment applies to Agency’s usage of both free and paid (fee-based) Services that Disqus may provide. The parties understand that fee-based products and services are categorically different than free products and services, as the former are subject to federal procurement regulations. Before an Agency decides to enter into a fee-based service that this Company or alternative providers may offer now or in the future, Agency und***erstands it must***: determine whether it has a need for those additional services for a fee; consider the service’s value in comparison with comparable offerings available elsewhere; to determine that Agency funds are available for payment; to properly use the Government Purchase Card if that Card is used as the payment method; to review the provider’s Terms for conformance to federal procurement law; and in all other respects to follow applicable federal acquisition laws, regulations, and Agency guidelines when contracting for a paid service.
​
R. ***Assignment***: Neither party may assign its obligations under the Terms as modified by this Amendment to any third party without prior written consent of the other; provided, however, Company or its subsidiaries may assign the Terms as modified by this Amendment to a subsidiary or parent without written consent from the Agency if the successor agrees to assume Company's obligations under the Terms as modified by this Amendment.
​
S. ***Provision of data***: In the event of termination of service, within 30 days of such termination Company will provide You with all User-generated Content that is publicly visible on the Service. Data will be provided in a commonly used file or database format as Company deems appropriate. Company will not provide data if doing so would violate its privacy policy (0).
​
T. ***Security***: Company will, in good faith, exercise due diligence using generally accepted commercial business practices for IT security, to ensure that systems are operated and maintained in a secure manner, and that management, operational and technical controls are employed to ensure security of systems and data. Recognizing the changing nature of the Web, Company will continuously work with Users to ensure that its Site and Services meet Users' requirements for the security of systems and data. Company agrees to discuss implementing additional security controls as deemed necessary by Agency to conform to the Federal Information Security Management Act (FISMA), 44 U.S.C. 3541 et seq.
​
U. ***Federal Records***: Agency acknowledges that its use of the Service as a public comment platform may generate federal records. If so, it is the Agency’s responsibility to manage those federal records in accordance with all applicable records management laws and regulations, including but not limited to the Federal Records Act (44 U.S.C. chs. 21, 29, 31, 33), and regulations of the National Archives and Records Administration (NARA) at 36 CFR Chapter XII Subchapter B). Managing the records includes, but is not limited to, secure storage, retrievability, and proper disposition of all federal records including transfer of permanently valuable records to NARA in a format and manner acceptable to NARA at the time of transfer.
​
V. ***Precedence; Further Amendments***: If there is a conflict between this Amendment and the Terms, or between this Amendment and other Disqus terms, rules and policies related to its Service, then this Amendment shall prevail. Any language in the Terms indicating it alone represents the entire agreement between Disqus and Agency is waived. Any further amendment must be agreed to by both parties.
​
W. ***Additional Items for discussion and possible inclusion in this Amendment***: Company understands that changes in federal law, regulation and policy may affect Agency's use of the Company's Service in ways not addressed in the preceding clauses. Among the topics Agency may seek to discuss with Company, and which may lead to a mutual agreement to insert additional clauses in this Amendment, are the subjects of privacy and accessibility.
​
\[July 2016\]

### Basic Rules for Disqus {#basic-rules-for-disqus}

Disqus doesn't moderate or manage the communities that use Disqus, but using Disqus to do any of the following things breaks our Terms of Service and appropriate action (which can include removing a comment or discussion, resetting a profile, or banning an account) will be taken to enforce them.

The following are not allowed anywhere on Disqus:

-   **Targeted harassment or encouraging others to do so**
    Hate speech and other forms of targeted and systematic harassment of people have no place on Disqus, nor do we tolerate communities dedicated to fostering harassing behavior.

-   **Spam**
    Examples include 1) comments posted in large quantities to promote a product or service, 2) the exact same comment posted repeatedly to disrupt a thread. 3) following users multiple times

-   **Impersonation**
    You may not impersonate others in a manner that does or is intended to mislead, confuse, or deceive others.

-   **Direct threat of harm**
    This covers active threats of harm directed towards a specific person or defined group of individuals. Contact local authorities if you feel a crime has been committed or is imminent.

-   **Posting personally identifiable information**
    Examples of protected information: credit card number, home/work address, phone number, email address, social security number. Real name isn't currently covered.

-   **Inappropriate profile content**
    Graphic media containing violence and pornographic content are not allowed. Profile content allowed by Disqus may not be allowed on all communities, so report such profiles to the site moderator.

To report a user for a Basic Rules violation, click the flag icon in their profile and complete a short report.

Learn more about how to report abuse to site moderators here.
​
For more information on how we enforce against abusive accounts that violate the Basic Rules, read our Abusive Behavior Policy.

### Basic Rules for Disqus-powered Sites {#basic-rules-for-disqus-powered-sites}

Disqus enables online discussion communities, and in doing so, freedom of expression and identity are core values of the Service. There are a number of categories of content and behavior, however, that jeopardize Disqus by posing risk to commenters, publishers, and/or third party services utilizing the Disqus platform.

Websites or website representatives, including site moderators, publishing inappropriate content or exhibiting inappropriate behaviors in connection with their use of the Service may have their Disqus account and/or Disqus forum suspended or terminated.

The following are not allowed on sites that use Disqus:

-   **Copyright or trademark Infringement**
    Sites hosting or linking to numerous copyrighted works may be removed from the Disqus network without warning. If we receive a valid DMCA claim against your site, your forum may be suspended until the claim is resolved.

-   **Deceitful data collection or distribution**
    User information is for moderation purposes only and collecting any information in a misleading way is prohibited. Distribution of personal identifiable information is prohibited.

-   **Intimidation of users of the Disqus Service**
    Blackmail, extortion, extreme discrimination, and other forms of threatening behavior are prohibited.

-   **Malware**
    If a site is found to be distributing malware, Disqus will be removed from that site.

-   **Misinformation**

    Sites containing harmful misinformation may be removed from the Disqus network.

-   **Unlawful activities**
    Disqus is controlled and operated from its facilities in the United States. Disqus makes no representations that it is appropriate or available for use in other locations. Those who access or use Disqus from other jurisdictions do so at their own volition and are entirely responsible for compliance with all applicable United States and local laws and regulations.

-   **Misuse of the Disqus Service**
    Sites that take any action that imposes, or may impose (at our sole discretion) an unreasonable or disproportionately large load on our infrastructure. This includes, but is not limited to: excessive creation of threads/forums/posts, misuse of the API, or a lack of moderation activity resulting in large volumes spam comments.

-   **Multiple violations of the Disqus Basic Rules**
    The Disqus Basic Rules apply to all Disqus accounts and violating them as a publisher or moderator may carry additional consequences for your forum(s).

The above list may be modified or expanded at any time and individual account/forum deactivation decisions remain at the sole discretion of Disqus. For complete details, please visit our Terms of Service.

Click here to report a violation of the Basic Rules for Disqus-powered sites.

### Comments Pricing and Plans {#comments-pricing-and-plans}

Disqus Comments is supported by a tiered pricing model based on feature set and site traffic. From personal blogs to media-giant, we’ve got you covered!
​
Please note that each plan has limits of eligibility based on pageviews and the number of sites in the Organization. If your subscription tier is not updated after crossing the pageview or site count threshold, you'll receive a warning and ads may run on your site.
​

*Disqus Comments does not include access to Disqus Polls, which require a separate subscription. Our Polls Pricing and Plans may be found [here](#disqus-polls-pricing-and-plans).*

Get all of the core Disqus features including: Comments plug-in, advanced spam filtering, moderation tools, basic analytics, configurable ads, and more.

Free, ads-supported with a Top ads placement required.
​
*Note: Sites that cannot run ads (due to adult content or other restrictions) may not be eligible for the Basic plan, and will be required to run an ads-free subscription to remain on the Disqus network.*

#### Mobile friendly

Optimized and designed for desktop, tablet, or mobile.

#### SEO optimized

Improve the visibility of your site in organic search results.

#### Export/Import comments

Bring in your old comments or have the ability to take your Disqus comments with you with our import and export tools.

#### 1,000 API requests per hour

Gather supplemental data and see many different points of data relating to comments on any site.

#### Multi-language support

Disqus currently supports over 65 languages, including Spanish, French, and Russian.

#### Social account login

Allow your audience to log in using Facebook, Twitter, and Google.

#### Cross-site web and email notifications

Instant activity notifications via web and email to pull readers back in.

#### Robust moderation tools

Automated Moderation Rules, Pre-moderation controls, Banning, email moderation notifications, and more.

#### **[Badges](#badges)**

Create custom badges to reward and recognize the commenters in your community. May be applied manually, or based on automated presets.

#### Embeddable comments

Package up notable comments and display them directly within your site content.

#### Community support

Search using our knowledge base and get help from community experts in real-time discussions on Discuss Disqus.

#### **Plus**

Get everything in Disqus Basic and the option to turn ads on and off.

Available to organizations containing up to 3 sites and a combined monthly pageview total listed below:
**up to 100,000 monthly pageviews**

\$12 per month, or \$132 per year (\$11/month with annual billing enabled)
**up to 350,000 monthly pageviews**

\$20 per month, or \$216 per year (\$18/month with annual billing enabled)
**up to 900,000 monthly pageviews**

\$35 per month, or \$372 per year (\$31/month with annual billing enabled)

#### Ads optional

Choose whether you would like to display ads with Disqus.

#### Direct Support

Email support with our friendly Customer Success team to ask questions and provide feedback.

#### **Pro**

Get everything in Disqus Plus and all of the additional features listed below.

Available to organizations containing up to 20 sites and a combined monthly pageview total listed below:

**up to 1,000,000 monthly pageviews**

\$115 per month, or \$1260 per year (\$105/month with annual billing enabled)

**up to 2,500,000 monthly pageviews**

\$140 per month, or \$1,500 per year (\$125/month with annual billing enabled)

**up to 5,000,000 monthly pageviews**

\$180 per month, or \$1,920 per year (\$160/month with annual billing enabled)

#### [Advanced Moderation](#advanced-moderation)

Get access to our automated moderation rules to make moderation a breeze.

#### Shadow Banning

Discretely remove troublesome users from your community without their knowledge.

#### Timeouts

Give users a chance to correct their behavior with temporary bans with the option to provide feedback.

#### **[New Commenter Pre-Moderation](#moderation-settings)**

Accounts new to your site can be set to Pending for a number of days specified, while not affecting your returning commenters. This can act as an additional guard against spam and unwanted behavior.

#### Advanced analytics

Uncover insights about your audience including top performing stories by engagement, the growth of your community over time, and your most engaged readers.

#### Additional API access

Utilize the Disqus API to build powerful integrations.

#### Priority support

Receive prioritized email support from our friendly Customer Success team to ask questions and provide feedback.

#### Branding and Style Options

Customize the look-and-feel of Disqus on your site including custom styling of the comment widget and the ability to remove Disqus branding.

#### **Star Ratings**

Allow users to add a 1-5 star rating along with their comment. The ratings spread will be shown at the top of the thread.

#### **[Custom Reactions](#reactions)**

Upload custom images for a unique set of reactions.

#### **[Email Subscriptions](#email-subscriptions)**

Allow users to opt into your newsletter through Disqus.

#### **Business**

Get everything in Pro, and more. Custom pricing. For enterprise companies and large publishers who want powerful tools and additional support.

#### Single Sign-On (SSO)

Allows users in your own database to comment without forcing them to register with Disqus.

####  Enterprise Login

Manage Disqus using your company login.

#### Direct Account Manager

1-on-1 account management that facilitates a closer business relationship with Disqus.

#### [Full Branding Removal](#disqus-appearance-customizations)

Remove all Disqus branding for a more native experience.

#### [Brandless Web Notifications](#disqus-web-notifications)

Provide a brandless on-site sidebar for commenting web notifications.

#### Unlimited API Access

Unlimited use of the Disqus API.

####  Custom Reporting & Optimizations

Contact us here for more info.

### Contacting Disqus about a deceased user {#contacting-disqus-about-a-deceased-user}

In the event of the death of a Disqus user, we can work with a person authorized to act on the behalf of the estate or with a verified immediate family member of the deceased to have an account deactivated. In order for us to process an account deactivation, please provide us with all of the following information:
- A URL that is a direct link to a comment made by the deceased person
- A copy of the deceased user’s death certificate
- A copy of your government-issued ID (e.g., driver’s license)
- A signed statement including:

-   Your first and last name

-   Your email address

-   Your current contact information

-   Your relationship to the deceased user or their estate

-   Action requested (e.g., ‘please deactivate the Disqus account’)

-   A brief description of the details that evidence this account belongs to the deceased, if the name on the account does not match the name on death certificate.

-   A link to an online obituary or a picture of the obituary from a local newspaper (optional)

Should we require any other information, we will contact you at the email address you have provided in your request. If you have any questions, you can contact us with our Support Form or through Discuss Disqus.

Please note: We are unable to provide account access to anyone regardless of his or her relationship to the deceased.

### Cookies and Data Recipients {#cookies-and-data-recipients}

**Data Recipients**
Disqus works with LiveRamp to help marketers connect browsers and devices with data from other sources that has been obfuscated to remove any directly identifying information. LiveRamp provides a privacy policy and opt-out options. We share information that we collect from you, such as your email (in an encoded form), IP address or information about your browser or operating system, with our identity partners/service providers, including LiveRamp Inc. LiveRamp matches your email with an online identification code that we may store in our first-party cookie for our use in online, in-app, and cross-channel advertising and it may be shared with advertising companies to enable interest-based and targeted advertising. To opt-out of this use, please click here.
​

We share your web browsing activity with Viglink to allow advertisers to personalize ads based on the types of products and services in which you seem to be interested. You may read Viglink’s privacy policy and opt-out.

We share your web browsing activity with Disqus' parent company, Zeta Global, to enable personalized marketing based on your interests, also known as cross-context behavioral advertising. Please see Zeta's privacy policy.
​
**Disqus Data Sub-processors**
We may share Disqus data with the following sub-processors to run and administer the Disqus service.
Amazon Web Services
Fastly
Hive AI
Hubspot
Intercom
Osano
Stripe
​

**Other Third Party Advertising Partners**

We work with several third party partners to serve relevant advertising within the Disqus comment embed on partner websites. We do not share your information with these partners, however, in the course of serving or displaying ads through Disqus, these partners may place a cookie on your browser. We expect all third party partners to follow all relevant privacy laws and regulations when serving ads through Disqus. 
​
AdaptMX
​
Amazon
​
AniView

Criteo
​
Google

Magnite

Mediafuse

OneTag

OpenX

PubMatic

Revcontent
​
Sonobi

Sovrn
​
Taboola

Xandr

Yahoo
​
Zeta Global SSP

**Updates to this Data Recipients Policy**. Disqus may, in its sole discretion, modify or update this Policy from time to time, and so you should review this page periodically. When we change the policy, we will update the ‘last modified’ date at the top of this page. Your continued use of the Site following the posting of any changes to this policy means you accept such changes.

**Contact**. If you have any questions about this Policy, please email us at privacy@disqus.com, or contact us by mail at 3 Park Avenue, 33rd Floor, New York, NY 10016.

### Data Processing Agreement for Publishers {#data-processing-agreement-for-publishers}

This Disqus Data Processing Agreement (“**DPA**”), that includes the Standard Contractual Clauses adopted by the European Commission, as applicable, reflects the parties’ agreement with respect to the terms governing the Processing of Personal Data under the Disqus Terms of Service (the “**Agreement**”). This DPA is an amendment to the Agreement and is effective upon its incorporation into the Agreement (sign-up). Upon its incorporation into the Agreement, the DPA will form a part of the Agreement.

We understand that some publishers may prefer to have a signed DPA for their records. Publishers can download a pre-signed version of the Disqus Publisher DPA via the link below. For any questions, please contact us at privacy@disqus.com.

Download Disqus Publisher DPA

**DATA PROCESSING AddendumAddendum to the Disqus Publisher Terms of Service Agreement**

This Addendum to the Disqus Publisher Terms of Service Agreement (the “DPA Addendum” or “DPA”), effective as of the Effective Date as set forth on the Agreement, specifies the global data protection obligations of Disqus Inc. (“Disqus”) and publisher (“Publisher”) under any agreement by which Disqus and Publisher process Personal Data and forms part of the Disqus Publisher Terms of Service Agreement (“Agreement”) previously entered into by the parties hereto.

WHEREAS, Disqus provides Publisher with the Disqus commenting application service (the “Disqus Comments” or “Services”) through which Disqus collects certain Personal Data from website users visiting the Publisher’s websites where the Disqus Comments are loaded, and Disqus further provides Publisher with the ability to access the comments left users on their website as well as some of the Personal Data associated with such comments;

WHEREAS, Privacy and Data Protection Laws (as defined below) impose compliance obligations upon Disqus and Publisher in relation to the collection and processing of Personal Data.

NOW THEREFORE, Pursuant to the requirements of the Privacy and Data Protection Laws , Disqus and Publisher hereby enter into this DPA.
​

**Definitions**

1.1 For the purposes of this DPA:

\(a\) **“EEA"** means the member states of the European Union and Iceland, Liechtenstein, Norway and the United Kingdom.

\(b\) **"Controller" or “Co-Controller”** shall mean an entity which, alone or jointly with others, determines the purposes and means of the processing of Personal Data;

\(c\) **"Processor"** shall mean an entity which processes Personal Data on behalf of the Controller;

\(d\) **“Personal Data”** means any information relating to an identified or identifiable natural person; an identifiable natural person is one who can be identified, directly or indirectly, in particular by reference to an identifier such as a name, and identification number, location data or online identifier.

\(e\) **“Disqus Personal Data”** means comments, content, data and information that is displayed, uploaded, exchanged, transmitted or collected through the Services provided to Publisher.

\(f\) **“Publisher Personal Data”** means all Personal Data provided by or collected on behalf of Publisher like single sign on data under the Agreement.

\(g\) “**Business Purpose**” means the purpose of providing the Services or any other purpose specifically identified in Exhibit A.

\(h\) “**Process**” **or “Processing”** means any operation or set of operations performed upon Personal Data, whether or not by automatic means

\(i\) “**Standard Contractual Clauses**” or **“SCC”** means the applicable module(s) of the European Commission’s standard contractual clauses for the transfer of personal data to third countries pursuant to Regulation (EU) 2016/679 of the European Parliament and of the Council, as set out in the Annex to Commission Implementing Decision (EU) 2021/914, a completed copy of which comprises Exhibit B.

\(j\) **“Restricted Country”** means a member state of the European Economic Area, Argentina, Brazil, China, Costa Rica, Ghana, Hong Kong, Israel, Malaysia, Mexico, Morocco, Russia, Singapore, Switzerland, Tunisia, Turkey, the United Kingdom, or Uruguay.
​

**2. Applicability of DPA.**
​

2.1. This DPA will apply to the extent that Publisher and Disqus Process Disqus Personal Data as Co-Controllers. To the extent that Publisher transfers Publisher Personal Data to Disqus, Disqus shall be a Processor and Publisher shall be a Controller.
​

2.2 This DPA is subject to the terms of the Agreement and is incorporated into the Agreement. Interpretations and defined terms set forth in the Agreement apply to the interpretation of this DPA.
​

2.3 The Appendices form part of this DPA and will have effect as if set out in full in the body of this DPA. Any reference to this DPA includes the Appendices.

**3. Roles and responsibilities.**
​

3.1 *Parties' Roles.*

Co-Controller. Disqus and Publisher each act as a Co- Controller with respect to the Disqus Personal Data processed hereunder. EXHIBIT A describes the Personal Data that Disqus makes available to Publisher and the purposes therefor. Publisher and Disqus undertake to access and use the Personal Data provided by Disqus only to the extent reasonably necessary to achieve the purposes of the processing.

Publisher as Controller and Disqus as Processor. To the extent that Publisher transfers Publisher Personal Data to Disqus, Publisher shall be the Controller of Publisher Personal Data, and Disqus shall be the Processor and Process Publisher Personal Data only in accordance with the permitted purposes.
​

3.2 *Purpose Limitation.* Both parties shall process the Personal Data solely for the purposes described in EXHIBIT A, except where required by applicable law.
​

3.3 Compliance: Each party, as Controller and Disqus as Processor, where applicable, shall be responsible for ensuring that it has complied, and will continue to comply, with all applicable laws relating to privacy and data protection, including but not limited to the EU data protection legislation (“Privacy and Data Protection Laws”).
​

3.4 *Representations and warranties.* Each party represents and warrants that it has sufficient legal rights to and in any Personal Data in order to transmit it to the other party as set forth herein or in the Agreement.
​

3.5 *CPRA.* Both parties, for the purposes of this DPA, may be deemed under the California Consumer Privacy Act of 2018 as amended by the California Privacy Rights Act of 2020 (“CPRA”) to share Personal Data to the other party. Both parties agree to comply with the requirements set forth in the CPRA Addendum described in Exhibit C.
​

3.5 *Written authorization.* Each party will only Process Personal Data pursuant to written directions as specified in the Agreement.
​

3.6 *Publisher’s rights and responsibilities.* Publisher has the technical means to turn off the collection and sharing of Personal Data by the Services at any time. Publisher is responsible for all content moderation which includes approval or removal of comments, blocking of users, setting up keyword alerts for content violation, and editing keywords.
​

3.7 *Accuracy.* Both parties shall ensure that Personal Data is accurate and, where necessary kept up to date, relevant, adequate, and in compliance with all applicable privacy and data security laws, rules and regulations.
​

**4. Security**
​

4.1 Security. Publisher and Disqus shall implement appropriate technical and organizational measures to protect the Personal Data from accidental or unlawful destruction, loss, alteration, unauthorized disclosure or access (each a "Security Incident"). Disqus will allow and cooperate with Publisher to conduct reasonable assessments or Disqus may arrange for a qualified and independent assessor to conduct an assessment of Disqus’ policies and technical and organizational measures, at least annually and at Disqus’ expense. Disqus shall provide a report of such assessment to Publisher upon request.
​

4.2 Confidentiality of Processing. Publisher and Disqus shall ensure that any person that it authorizes to process the Personal Data shall be subject to a contractual or statutory duty of confidentiality. Neither party shall sell, rent, lease, disclose, disseminate, make available, transfer, or otherwise communicate orally, in writing, or by electronic or other means, Personal Data to another business, person, or third party without the other party’s prior written consent.
​

4.3 Security Incidents. Each party will promptly and without undue delay and in any case no later than twenty-four (24) hours of becoming aware, inform the other party in the event of: (i) any breach of security leading to the accidental or unlawful destruction, loss, alteration, unauthorized disclosures of, or access to, Personal Data (altogether, a “**Security Incident**”), or (ii) any reasonable suspicion of a Security Incident, regardless of its cause. At Co-Controller’s direction, Controller will provide all relevant information and assistance required by Co-Controller to investigate, mitigate and respond to a Security Incident, including at a minimum, any information or assistance required by applicable Privacy and Data Protection Laws.
​

4.4 Requests from data protection authorities. Co-Controller shall reasonably assist Controller in response to any requests from data protection authorities relating to the Processing of Personal Data in connection with the Agreement. In the event that any such request is made directly to Co-Controller, Controller shall not respond to such communication directly without Co-Controller’s prior authorization, unless legally compelled to do so. If Controller is required to respond to such a request, Controller shall promptly notify Co-Controller and provide it with a copy of the request unless legally prohibited from doing so.

**5. Sub-processing**
​

5.1 Processors and sub-Processors. Co-Controller may engage Co-Controller affiliates and third party Data Processors or sub-Processors to process the Personal Data. Co-Controller shall inform Controller of any intended changes concerning the addition or replacement of sub-Processors, giving Controller the opportunity to object to such changes. Co-Controller shall impose on such Processors data protection terms that protect the Personal Data to the same standard provided for by this DPA and shall remain liable for any breach of the DPA caused by a Processor or sub-Processor. If Co-Controller subcontracts or assigns any of Co-Controller’s obligations to a third party, Co-Controller will in each case: (a) first ensure that each and every such subcontractor, partner or assignee (as the case may be) has undertaken in signed writing to comply with obligations no less protective than the obligations undertaken by Co-Controller in this Addendum; (b) perform appropriate due diligence to ensure that all subcontractors, partners and assignees can meet all Co-Controller obligations in the Agreement, including all requirements related to features, functionality and assistance necessary for data subject requests; (c) remain fully liable for the performance of each subcontractor, partner and/or assignee; and (d) enter into Standard Contractual Clauses.
​

**6. International transfers.**
​

6.1 *International Transfers:*

Where Co-Controller transfers Personal Data outside the EEA in a country in respect of which a valid adequacy decision has not been issued by the European Commission or adequacy has not otherwise been determined in another valid method under applicable data protection laws then an adequate level of protection shall be put in place by entering into Standard Contractual Clauses, a completed copy of which comprises Exhibit B and which are hereby incorporated by reference or through any other recognized methods. Co-Controller authorizes any transfers of Personal Data to, or access to Personal Data from, such destinations outside the EEA subject to such adequacy measures having been taken. The Controller-to-Processor Standard Contractual Clauses shall apply in all cases where Personal Data that relates to residents of the EEA is Processed by Disqus. The Controller-to-Controller Standard Contractual Clauses will also apply where, and to the extent that, Publisher acts as a Co-Controller with respect to any Personal Data that relates to a resident of the EEA. In particular, and without limiting the above obligations:

i\. Publisher and Disqus agree that their respective obligations under the Standard Contractual Clauses shall be governed by the law(s) of the Member State(s) (or Switzerland or the United Kingdom) in which Publishers are established; and

ii\. the details of the appendices applicable to the Standard Contractual Clauses are set out in **Exhibit B** to this Addendum.

6.2 Disclosure to authorities: Co-Controller acknowledges that Controller may disclose the privacy provisions in this DPA and the Agreement to the US Department of Commerce, the Federal Trade Commission, a European data protection authority, or any other US or EU judicial or regulatory body upon their lawful request.
​

**7. Cooperation**
​

7.1 Cooperation and data subjects' rights. Co-Controller shall reasonably cooperate with Controller in all matters pertaining to the Personal Data and shall provide Controller information about its uses of Personal Data upon request. Co-Controller shall respond and give effect to requests from data subjects seeking to exercise their rights under Privacy and Data Protection Laws. If Co-Controller cannot reasonably respond to a request by a data subject it may refer the data subject to Controller as appropriate. Co-Controller will provide all other reasonable assistance and execute such agreements as may be necessary to legitimize any Processing or data transfer of Personal Data to Controller or a subcontractor and to ensure an adequate level of protection for Personal Data. In the event that any competent authority holds that a data transfer mechanism relied on by the parties is invalid, or any supervisory authority requires transfers of Personal Data made pursuant to such decision to be suspended, then Co-Controller may, at its discretion, require Controller to cease Processing Personal Data, or co-operate with it to facilitate use of an alternative transfer mechanism.

7.2 Data Protection Impact Assessments: Co-Controller shall, to the extent required by Privacy and Data Protection Laws, provide Controller with commercially reasonable assistance with any future data protection impact assessments or prior consultations with data protection authorities that Controller is required to carry out under Privacy and Data Protection Laws.

**8. Security reports and audits.**
​

8.1 Co-Controller shall provide, upon Controller's request, copies of any relevant summaries of external security certifications or security audit reports necessary to verify Publisher’s compliance with this DPA.
​

8.2 While it is the parties' intention ordinarily to rely on the provision of the documentation at 8.1 above to verify Co-Controller's compliance with this DPA, Co-Controller shall permit Controller (or its appointed third party auditors) to carry out an audit of Co-Controller’s Processing of Personal Data under the Agreement following a Security Incident suffered by Controller, or upon the instruction of a data protection authority. Controller must give Co-Controller reasonable prior notice of such intention to audit, conduct its audit during normal business hours, and take all reasonable measures to prevent unnecessary disruption to Co-Controller's operations. Any such audit shall be subject to Co-Controller's security and confidentiality terms and guidelines.
​

8.3 Each Controller shall implement reasonable and appropriate technical, physical, and organizational measures designed to adequately safeguard and protect against a Security Incident (each, a “**Security Measure**”). Such Security Measures shall require each Controller to have regard to industry standards and costs of implementation as well as taking into account the nature, scope, context, and purposes of the Processing as well as the risk of harm that may result from a Security Incident to Co-Controller.

9\. Deletion or return of data: Upon termination or expiry of the Agreement, each Controller shall delete the Personal Data (including copies) then in Controller’s possession, except to the extent that Controller is required by an applicable law to retain some or all of the Personal Data.

10\. Term: The term of this Addendum commences as of the Addendum Effective Date and will end upon the termination of the Agreement. However, each Controller’s obligations hereunder continue in effect until any Personal Data subject to this DPA is returned or destroyed

11\. Indemnity: Any indemnity obligations will be covered pursuant to the Agreement.

12\. Governing Law: Unless otherwise required by the Standard Contractual Clauses or other data transfer requirements, this Addendum will be subject to the governing law identified in the Agreement without giving effect to conflict of laws principles.

13\. Counterparts: This Addendum may be entered into by the parties in any number of counterparts. Each counterpart will, when executed and delivered, be regarded as an original, and all the counterparts will together constitute one and the same instrument.

14\. Modifications: During the term of this DPA, Disqus may revise the terms and conditions of this DPA at any time. Any such revision or change will be binding and effective immediately on posting of the revised DPA on Disqus’ homepage.

IN WITNESS WHEREOF, Disqus and Publisher have executed this Addendum, effective as of the date the Agreement is signed (the “**Addendum Effective Date**”).

**IN WITNESS WHEREOF** the parties hereto have executed this Addendum as of the date first mentioned above:
​

Acknowledged and Agreed to:

Disqus, Inc. PUBLISHER:

Signed: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ Signed: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Name: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ Name: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Title: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ Title: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_ Date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

​

**EXHIBIT A DETAILS OF THE PROCESSING*Description of Disqus:*** Disqus, Inc. is the legal entity that has executed the Agreement with Publisher for the provision of Disqus' commenting application services on Publisher’s website.

***Purposes of Processing:*** Disqus provides a commenting application service (“Disqus Comments”) to Publisher for use as a comment forum on the Publisher’s website. Disqus collects Personal Data from users commenting in the Disqus Comments on Publisher’s website. Disqus provides Publisher with access to the comments so that Publisher may act as moderator on its website, and to meet its relevant obligations under applicable laws. The comments may include Personal Data such as email address, username, IP address or other online identifier, which Publisher may process solely for the purpose of moderating Publishers’ site(s).

***Type(s) of Personal Data Processed:*** Email address, username, IP address, other online identifier, information revealed in user comments.

Special categories of data (if applicable): Disqus does not intentionally collect, and Publisher does not intentionally transfer any sensitive personal data in relation to these data subjects. Publisher may collect categories of sensitive personal data contained in user comments as part of its comment moderation activities only in accordance with applicable privacy laws.

***Categories of Data Subjects:*** The Personal Data Processed concern individuals who access the Publishers website on which the Disqus Comments are loaded.

Nature of Processing Operations: Publisher will Process the Personal Data solely for the purpose of moderating the comments on their website and meeting any applicable legal requirements.

***Duration of the Processing:*** As set forth in the Agreement.

​

**EXHIBIT B STANDARD CONTRACTUAL CLAUSES**

#### **STANDARD CONTRACTUAL CLAUSES**

Controller to Controller

**SECTION I**
​

***Clause 1:* Purpose and scope**

\(a\) The purpose of these standard contractual clauses is to ensure compliance with the requirements of Regulation (EU) 2016/679 of the European Parliament and of the Council of 27 April 2016 on the protection of natural persons with regard to the processing of personal data and on the free movement of such data (General Data Protection Regulation) for the transfer of personal data to a third country.

\(b\) The Parties:

\(i\) the natural or legal person(s), public authority/ies, agency/ies or other body/ies (hereinafter ‘entity/ies’) transferring the personal data, as listed in Annex I.A (hereinafter each ‘data exporter’), and

\(ii\) the entity/ies in a third country receiving the personal data from the data exporter, directly or indirectly via another entity also Party to these Clauses, as listed in Annex I.A (hereinafter each ‘data importer’)

have agreed to these standard contractual clauses (hereinafter: ‘Clauses’).

\(c\) These Clauses apply with respect to the transfer of personal data as specified in Annex I.B.

\(d\) The Appendix to these Clauses containing the Annexes referred to therein forms an integral part of these Clauses.
​

***Clause 2:* Effect and invariability of the Clauses**

\(a\) These Clauses set out appropriate safeguards, including enforceable data subject rights and effective legal remedies, pursuant to Article 46(1) and Article 46(2)(c) of Regulation (EU) 2016/679 and, with respect to data transfers from controllers to processors and/or processors to processors, standard contractual clauses pursuant to Article 28(7) of Regulation (EU) 2016/679, provided they are not modified, except to select the appropriate Module(s) or to add or update information in the Appendix. This does not prevent the Parties from including the standard contractual clauses laid down in these Clauses in a wider contract and/or to add other clauses or additional safeguards, provided that they do not contradict, directly or indirectly, these Clauses or prejudice the fundamental rights or freedoms of data subjects.

\(b\) These Clauses are without prejudice to obligations to which the data exporter is subject by virtue of Regulation (EU) 2016/679.
​

***Clause 3:* Third-party beneficiaries**

\(a\) Data subjects may invoke and enforce these Clauses, as third-party beneficiaries, against the data exporter and/or data importer, with the following exceptions:

\(i\) Clause 1, Clause 2, Clause 3, Clause 6, Clause 7;

\(ii\) Clause 8.5 (e) and Clause 8.9(b);

\(iii\) N/A

\(iv\) Clause 12(a) and (d);

\(v\) Clause 13;

\(vi\) Clause 15.1(c), (d) and (e);

\(vii\) Clause 16(e);

\(viii\) Clause 18(a) and (b).

\(b\) Paragraph (a) is without prejudice to rights of data subjects under Regulation (EU) 2016/679.
​

***Clause 4:* Interpretation**

\(a\) Where these Clauses use terms that are defined in Regulation (EU) 2016/679, those terms shall have the same meaning as in that Regulation.

\(b\) These Clauses shall be read and interpreted in the light of the provisions of Regulation (EU) 2016/679.

\(c\) These Clauses shall not be interpreted in a way that conflicts with rights and obligations provided for in Regulation (EU) 2016/679.
​

***Clause 5:* Hierarchy**

In the event of a contradiction between these Clauses and the provisions of related agreements between the Parties, existing at the time these Clauses are agreed or entered into thereafter, these Clauses shall prevail.
​

***Clause 6:* Description of the transfer(s)**

The details of the transfer(s), and in particular the categories of personal data that are transferred and the purpose(s) for which they are transferred, are specified in Annex I.B.

***Clause 7:* Docking clause**

\(a\) An entity that is not a Party to these Clauses may, with the agreement of the Parties, accede to these Clauses at any time, either as a data exporter or as a data importer, by completing the Appendix and signing Annex I.A.

\(b\) Once it has completed the Appendix and signed Annex I.A, the acceding entity shall become a Party to these Clauses and have the rights and obligations of a data exporter or data importer in accordance with its designation in Annex I.A.

\(c\) The acceding entity shall have no rights or obligations arising under these Clauses from the period prior to becoming a Party.
​

**SECTION II – OBLIGATIONS OF THE PARTIES**
​

***Clause 8:* Data protection safeguards**

The data exporter warrants that it has used reasonable efforts to determine that the data importer is able, through the implementation of appropriate technical and organisational measures, to satisfy its obligations under these Clauses.

**8.1 Purpose limitation**

The data importer shall process the personal data only for the specific purpose(s) of the transfer, as set out in Annex I.B. It may only process the personal data for another purpose:

\(i\) where it has obtained the data subject’s prior consent;

\(ii\) where necessary for the establishment, exercise or defence of legal claims in the context of specific administrative, regulatory or judicial proceedings; or

\(iii\) where necessary in order to protect the vital interests of the data subject or of another natural person.

**8.2 Transparency**

\(a\) In order to enable data subjects to effectively exercise their rights pursuant to Clause 10, the data importer shall inform them, either directly or through the data exporter:

\(i\) of its identity and contact details;

\(ii\) of the categories of personal data processed;

\(iii\) of the right to obtain a copy of these Clauses;

\(iv\) where it intends to onward transfer the personal data to any third party/ies, of the recipient or categories of recipients (as appropriate with a view to providing meaningful information), the purpose of such onward transfer and the ground therefore pursuant to Clause 8.7.

\(b\) Paragraph (a) shall not apply where the data subject already has the information, including when such information has already been provided by the data exporter, or providing the information proves impossible or would involve a disproportionate effort for the data importer. In the latter case, the data importer shall, to the extent possible, make the information publicly available.

\(c\) On request, the Parties shall make a copy of these Clauses, including the Appendix as completed by them, available to the data subject free of charge. To the extent necessary to protect business secrets or other confidential information, including personal data, the Parties may redact part of the text of the Appendix prior to sharing a copy, but shall provide a meaningful summary where the data subject would otherwise not be able to understand its content or exercise his/her rights. On request, the Parties shall provide the data subject with the reasons for the redactions, to the extent possible without revealing the redacted information.

\(d\) Paragraphs (a) to (c) are without prejudice to the obligations of the data exporter under Articles 13 and 14 of Regulation (EU) 2016/679.

**8.3 Accuracy and data minimisation**

\(a\) Each Party shall ensure that the personal data is accurate and, where necessary, kept up to date. The data importer shall take every reasonable step to ensure that personal data that is inaccurate, having regard to the purpose(s) of processing, is erased or rectified without delay.

\(b\) If one of the Parties becomes aware that the personal data it has transferred or received is inaccurate, or has become outdated, it shall inform the other Party without undue delay.

\(c\) The data importer shall ensure that the personal data is adequate, relevant and limited to what is necessary in relation to the purpose(s) of processing.

**8.4 Storage limitation**

The data importer shall retain the personal data for no longer than necessary for the purpose(s) for which it is processed. It shall put in place appropriate technical or organisational measures to ensure compliance with this obligation, including erasure or anonymisation of the data and all back-ups at the end of the retention period.

**8.5 Security of processing**

\(a\) The data importer and, during transmission, also the data exporter shall implement appropriate technical and organisational measures to ensure the security of the personal data, including protection against a breach of security leading to accidental or unlawful destruction, loss, alteration, unauthorised disclosure or access (hereinafter ‘personal data breach’). In assessing the appropriate level of security, they shall take due account of the state of the art, the costs of implementation, the nature, scope, context and purpose(s) of processing and the risks involved in the processing for the data subject. The Parties shall in particular consider having recourse to encryption or pseudonymisation, including during transmission, where the purpose of processing can be fulfilled in that manner.

\(b\) The Parties have agreed on the technical and organisational measures set out in Annex II. The data importer shall carry out regular checks to ensure that these measures continue to provide an appropriate level of security.

\(c\) The data importer shall ensure that persons authorised to process the personal data have committed themselves to confidentiality or are under an appropriate statutory obligation of confidentiality.

\(d\) In the event of a personal data breach concerning personal data processed by the data importer under these Clauses, the data importer shall take appropriate measures to address the personal data breach, including measures to mitigate its possible adverse effects.

\(e\) In case of a personal data breach that is likely to result in a risk to the rights and freedoms of natural persons, the data importer shall without undue delay notify both the data exporter and the competent supervisory authority pursuant to Clause 13. Such notification shall contain i) a description of the nature of the breach (including, where possible, categories and approximate number of data subjects and personal data records concerned), ii) its likely consequences, iii) the measures taken or proposed to address the breach, and iv) the details of a contact point from whom more information can be obtained. To the extent it is not possible for the data importer to provide all the information at the same time, it may do so in phases without undue further delay.

\(f\) In case of a personal data breach that is likely to result in a high risk to the rights and freedoms of natural persons, the data importer shall also notify without undue delay the data subjects concerned of the personal data breach and its nature, if necessary in cooperation with the data exporter, together with the information referred to in paragraph (e), points ii) to iv), unless the data importer has implemented measures to significantly reduce the risk to the rights or freedoms of natural persons, or notification would involve disproportionate efforts. In the latter case, the data importer shall instead issue a public communication or take a similar measure to inform the public of the personal data breach.

\(g\) The data importer shall document all relevant facts relating to the personal data breach, including its effects and any remedial action taken, and keep a record thereof.

**8.6 Sensitive data**

Where the transfer involves personal data revealing racial or ethnic origin, political opinions, religious or philosophical beliefs, or trade union membership, genetic data, or biometric data for the purpose of uniquely identifying a natural person, data concerning health or a person’s sex life or sexual orientation, or data relating to criminal convictions or offences (hereinafter ‘sensitive data’), the data importer shall apply specific restrictions and/or additional safeguards adapted to the specific nature of the data and the risks involved. This may include restricting the personnel permitted to access the personal data, additional security measures (such as pseudonymisation) and/or additional restrictions with respect to further disclosure.

**8.7 Onward transfers**

The data importer shall not disclose the personal data to a third party located outside the European Union (in the same country as the data importer or in another third country, hereinafter ‘onward transfer’) unless the third party is or agrees to be bound by these Clauses, under the appropriate Module. Otherwise, an onward transfer by the data importer may only take place if:

\(i\) it is to a country benefitting from an adequacy decision pursuant to Article 45 of Regulation (EU) 2016/679 that covers the onward transfer;

\(ii\) the third party otherwise ensures appropriate safeguards pursuant to Articles 46 or 47 of Regulation (EU) 2016/679 with respect to the processing in question;

\(iii\) the third party enters into a binding instrument with the data importer ensuring the same level of data protection as under these Clauses, and the data importer provides a copy of these safeguards to the data exporter;

\(iv\) it is necessary for the establishment, exercise or defence of legal claims in the context of specific administrative, regulatory or judicial proceedings;

\(v\) it is necessary in order to protect the vital interests of the data subject or of another natural person; or

\(vi\) where none of the other conditions apply, the data importer has obtained the explicit consent of the data subject for an onward transfer in a specific situation, after having informed him/her of its purpose(s), the identity of the recipient and the possible risks of such transfer to him/her due to the lack of appropriate data protection safeguards. In this case, the data importer shall inform the data exporter and, at the request of the latter, shall transmit to it a copy of the information provided to the data subject.

Any onward transfer is subject to compliance by the data importer with all the other safeguards under these Clauses, in particular purpose limitation.

**8.8 Processing under the authority of the data importer**

The data importer shall ensure that any person acting under its authority, including a processor, processes the data only on its instructions.

**8.9 Documentation and compliance**

\(a\) Each Party shall be able to demonstrate compliance with its obligations under these Clauses. In particular, the data importer shall keep appropriate documentation of the processing activities carried out under its responsibility.

\(b\) The data importer shall make such documentation available to the competent supervisory authority on request.
​

***Clause 9:* Use of sub-processors ** N/A
​

***Clause 10:* Data subject rights**

\(a\) The data importer, where relevant with the assistance of the data exporter, shall deal with any enquiries and requests it receives from a data subject relating to the processing of his/her personal data and the exercise of his/her rights under these Clauses without undue delay and at the latest within one month of the receipt of the enquiry or request. The data importer shall take appropriate measures to facilitate such enquiries, requests and the exercise of data subject rights. Any information provided to the data subject shall be in an intelligible and easily accessible form, using clear and plain language.

\(b\) In particular, upon request by the data subject the data importer shall, free of charge:

\(i\) provide confirmation to the data subject as to whether personal data concerning him/her is being processed and, where this is the case, a copy of the data relating to him/her and the information in Annex I; if personal data has been or will be onward transferred, provide information on recipients or categories of recipients (as appropriate with a view to providing meaningful information) to which the personal data has been or will be onward transferred, the purpose of such onward transfers and their ground pursuant to Clause 8.7; and provide information on the right to lodge a complaint with a supervisory authority in accordance with Clause 12(c)(i);

\(ii\) rectify inaccurate or incomplete data concerning the data subject;

\(iii\) erase personal data concerning the data subject if such data is being or has been processed in violation of any of these Clauses ensuring third-party beneficiary rights, or if the data subject withdraws the consent on which the processing is based.

\(c\) Where the data importer processes the personal data for direct marketing purposes, it shall cease processing for such purposes if the data subject objects to it.

\(d\) The data importer shall not make a decision based solely on the automated processing of the personal data transferred (hereinafter ‘automated decision’), which would produce legal effects concerning the data subject or similarly significantly affect him/her, unless with the explicit consent of the data subject or if authorised to do so under the laws of the country of destination, provided that such laws lays down suitable measures to safeguard the data subject’s rights and legitimate interests. In this case, the data importer shall, where necessary in cooperation with the data exporter:

\(i\) inform the data subject about the envisaged automated decision, the envisaged consequences and the logic involved; and

\(ii\) implement suitable safeguards, at least by enabling the data subject to contest the decision, express his/her point of view and obtain review by a human being.

\(e\) Where requests from a data subject are excessive, in particular because of their repetitive character, the data importer may either charge a reasonable fee taking into account the administrative costs of granting the request or refuse to act on the request.

\(f\) The data importer may refuse a data subject’s request if such refusal is allowed under the laws of the country of destination and is necessary and proportionate in a democratic society to protect one of the objectives listed in Article 23(1) of Regulation (EU) 2016/679.

\(g\) If the data importer intends to refuse a data subject’s request, it shall inform the data subject of the reasons for the refusal and the possibility of lodging a complaint with the competent supervisory authority and/or seeking judicial redress.
​

***Clause 11:* Redress**

\(a\) The data importer shall inform data subjects in a transparent and easily accessible format, through individual notice or on its website, of a contact point authorised to handle complaints. It shall deal promptly with any complaints it receives from a data subject.

\(b\) In case of a dispute between a data subject and one of the Parties as regards compliance with these Clauses, that Party shall use its best efforts to resolve the issue amicably in a timely fashion. The Parties shall keep each other informed about such disputes and, where appropriate, cooperate in resolving them.

\(c\) Where the data subject invokes a third-party beneficiary right pursuant to Clause 3, the data importer shall accept the decision of the data subject to:

\(i\) lodge a complaint with the supervisory authority in the Member State of his/her habitual residence or place of work, or the competent supervisory authority pursuant to Clause 13;

\(ii\) refer the dispute to the competent courts within the meaning of Clause 18.

\(d\) The Parties accept that the data subject may be represented by a not-for-profit body, organisation or association under the conditions set out in Article 80(1) of Regulation (EU) 2016/679.

\(e\) The data importer shall abide by a decision that is binding under the applicable EU or Member State law.

\(f\) The data importer agrees that the choice made by the data subject will not prejudice his/her substantive and procedural rights to seek remedies in accordance with applicable laws.

***Clause 12:* Liability**

\(a\) Each Party shall be liable to the other Party/ies for any damages it causes the other Party/ies by any breach of these Clauses.

\(b\) Each Party shall be liable to the data subject, and the data subject shall be entitled to receive compensation, for any material or non-material damages that the Party causes the data subject by breaching the third-party beneficiary rights under these Clauses. This is without prejudice to the liability of the data exporter under Regulation (EU) 2016/679.

\(c\) Where more than one Party is responsible for any damage caused to the data subject as a result of a breach of these Clauses, all responsible Parties shall be jointly and severally liable and the data subject is entitled to bring an action in court against any of these Parties.

\(d\) The Parties agree that if one Party is held liable under paragraph (c), it shall be entitled to claim back from the other Party/ies that part of the compensation corresponding to its/their responsibility for the damage.

\(e\) The data importer may not invoke the conduct of a processor or sub-processor to avoid its own liability.
​

***Clause 13:* Supervision**

\(a\) \[Where the data exporter is established in an EU Member State:\] The supervisory authority with responsibility for ensuring compliance by the data exporter with Regulation (EU) 2016/679 as regards the data transfer, as indicated in Annex I.C, shall act as competent supervisory authority.

\[Where the data exporter is not established in an EU Member State, but falls within the territorial scope of application of Regulation (EU) 2016/679 in accordance with its Article 3(2) and has appointed a representative pursuant to Article 27(1) of Regulation (EU) 2016/679:\] The supervisory authority of the Member State in which the representative within the meaning of Article 27(1) of Regulation (EU) 2016/679 is established, as indicated in Annex I.C, shall act as competent supervisory authority.

\[Where the data exporter is not established in an EU Member State, but falls within the territorial scope of application of Regulation (EU) 2016/679 in accordance with its Article 3(2) without however having to appoint a representative pursuant to Article 27(2) of Regulation (EU) 2016/679:\] The supervisory authority of one of the Member States in which the data subjects whose personal data is transferred under these Clauses in relation to the offering of goods or services to them, or whose behaviour is monitored, are located, as indicated in Annex I.C, shall act as competent supervisory authority.

\(b\) The data importer agrees to submit itself to the jurisdiction of and cooperate with the competent supervisory authority in any procedures aimed at ensuring compliance with these Clauses. In particular, the data importer agrees to respond to enquiries, submit to audits and comply with the measures adopted by the supervisory authority, including remedial and compensatory measures. It shall provide the supervisory authority with written confirmation that the necessary actions have been taken.
​

**SECTION III – LOCAL LAWS AND OBLIGATIONS IN CASE OF ACCESS BY PUBLIC AUTHORITIES**
​

***Clause 14:* Local laws and practices affecting compliance with the Clauses**

\(a\) The Parties warrant that they have no reason to believe that the laws and practices in the third country of destination applicable to the processing of the personal data by the data importer, including any requirements to disclose personal data or measures authorising access by public authorities, prevent the data importer from fulfilling its obligations under these Clauses. This is based on the understanding that laws and practices that respect the essence of the fundamental rights and freedoms and do not exceed what is necessary and proportionate in a democratic society to safeguard one of the objectives listed in Article 23(1) of Regulation (EU) 2016/679, are not in contradiction with these Clauses.

\(b\) The Parties declare that in providing the warranty in paragraph (a), they have taken due account in particular of the following elements:

\(i\) the specific circumstances of the transfer, including the length of the processing chain, the number of actors involved and the transmission channels used; intended onward transfers; the type of recipient; the purpose of processing; the categories and format of the transferred personal data; the economic sector in which the transfer occurs; the storage location of the data transferred;

\(ii\) the laws and practices of the third country of destination– including those requiring the disclosure of data to public authorities or authorising access by such authorities – relevant in light of the specific circumstances of the transfer, and the applicable limitations and safeguards;

\(iii\) any relevant contractual, technical or organisational safeguards put in place to supplement the safeguards under these Clauses, including measures applied during transmission and to the processing of the personal data in the country of destination.

\(c\) The data importer warrants that, in carrying out the assessment under paragraph (b), it has made its best efforts to provide the data exporter with relevant information and agrees that it will continue to cooperate with the data exporter in ensuring compliance with these Clauses.

\(d\) The Parties agree to document the assessment under paragraph (b) and make it available to the competent supervisory authority on request.

\(e\) The data importer agrees to notify the data exporter promptly if, after having agreed to these Clauses and for the duration of the contract, it has reason to believe that it is or has become subject to laws or practices not in line with the requirements under paragraph (a), including following a change in the laws of the third country or a measure (such as a disclosure request) indicating an application of such laws in practice that is not in line with the requirements in paragraph (a).

\(f\) Following a notification pursuant to paragraph (e), or if the data exporter otherwise has reason to believe that the data importer can no longer fulfil its obligations under these Clauses, the data exporter shall promptly identify appropriate measures (e.g. technical or organisational measures to ensure security and confidentiality) to be adopted by the data exporter and/or data importer to address the situation. The data exporter shall suspend the data transfer if it considers that no appropriate safeguards for such transfer can be ensured, or if instructed by the competent supervisory authority to do so. In this case, the data exporter shall be entitled to terminate the contract, insofar as it concerns the processing of personal data under these Clauses. If the contract involves more than two Parties, the data exporter may exercise this right to termination only with respect to the relevant Party, unless the Parties have agreed otherwise. Where the contract is terminated pursuant to this Clause, Clause 16(d) and (e) shall apply.
​

***Clause 15:* Obligations of the data importer in case of access by public authorities15.1 Notification**

\(a\) The data importer agrees to notify the data exporter and, where possible, the data subject promptly (if necessary with the help of the data exporter) if it:

\(i\) receives a legally binding request from a public authority, including judicial authorities, under the laws of the country of destination for the disclosure of personal data transferred pursuant to these Clauses; such notification shall include information about the personal data requested, the requesting authority, the legal basis for the request and the response provided; or

\(ii\) becomes aware of any direct access by public authorities to personal data transferred pursuant to these Clauses in accordance with the laws of the country of destination; such notification shall include all information available to the importer.

\(b\) If the data importer is prohibited from notifying the data exporter and/or the data subject under the laws of the country of destination, the data importer agrees to use its best efforts to obtain a waiver of the prohibition, with a view to communicating as much information as possible, as soon as possible. The data importer agrees to document its best efforts in order to be able to demonstrate them on request of the data exporter.

\(c\) Where permissible under the laws of the country of destination, the data importer agrees to provide the data exporter, at regular intervals for the duration of the contract, with as much relevant information as possible on the requests received (in particular, number of requests, type of data requested, requesting authority/ies, whether requests have been challenged and the outcome of such challenges, etc.).

\(d\) The data importer agrees to preserve the information pursuant to paragraphs (a) to (c) for the duration of the contract and make it available to the competent supervisory authority on request.

\(e\) Paragraphs (a) to (c) are without prejudice to the obligation of the data importer pursuant to Clause 14(e) and Clause 16 to inform the data exporter promptly where it is unable to comply with these Clauses.

**15.2 Review of legality and data minimisation**

\(a\) The data importer agrees to review the legality of the request for disclosure, in particular whether it remains within the powers granted to the requesting public authority, and to challenge the request if, after careful assessment, it concludes that there are reasonable grounds to consider that the request is unlawful under the laws of the country of destination, applicable obligations under international law and principles of international comity. The data importer shall, under the same conditions, pursue possibilities of appeal. When challenging a request, the data importer shall seek interim measures with a view to suspending the effects of the request until the competent judicial authority has decided on its merits. It shall not disclose the personal data requested until required to do so under the applicable procedural rules. These requirements are without prejudice to the obligations of the data importer under Clause 14(e).

\(b\) The data importer agrees to document its legal assessment and any challenge to the request for disclosure and, to the extent permissible under the laws of the country of destination, make the documentation available to the data exporter. It shall also make it available to the competent supervisory authority on request.

\(c\) The data importer agrees to provide the minimum amount of information permissible when responding to a request for disclosure, based on a reasonable interpretation of the request.
​

**SECTION IV – FINAL PROVISIONS**
​

***Clause 16:* Non-compliance with the Clauses and termination**

\(a\) The data importer shall promptly inform the data exporter if it is unable to comply with these Clauses, for whatever reason.

\(b\) In the event that the data importer is in breach of these Clauses or unable to comply with these Clauses, the data exporter shall suspend the transfer of personal data to the data importer until compliance is again ensured or the contract is terminated. This is without prejudice to Clause 14(f).

\(c\) The data exporter shall be entitled to terminate the contract, insofar as it concerns the processing of personal data under these Clauses, where:

\(i\) the data exporter has suspended the transfer of personal data to the data importer pursuant to paragraph (b) and compliance with these Clauses is not restored within a reasonable time and in any event within one month of suspension;

\(ii\) the data importer is in substantial or persistent breach of these Clauses; or

\(iii\) the data importer fails to comply with a binding decision of a competent court or supervisory authority regarding its obligations under these Clauses.

In these cases, it shall inform the competent supervisory authority of such non-compliance. Where the contract involves more than two Parties, the data exporter may exercise this right to termination only with respect to the relevant Party, unless the Parties have agreed otherwise.

\(d\) Personal data that has been transferred prior to the termination of the contract pursuant to paragraph (c) shall at the choice of the data exporter immediately be returned to the data exporter or deleted in its entirety. The same shall apply to any copies of the data. The data importer shall certify the deletion of the data to the data exporter. Until the data is deleted or returned, the data importer shall continue to ensure compliance with these Clauses. In case of local laws applicable to the data importer that prohibit the return or deletion of the transferred personal data, the data importer warrants that it will continue to ensure compliance with these Clauses and will only process the data to the extent and for as long as required under that local law.

\(e\) Either Party may revoke its agreement to be bound by these Clauses where (i) the European Commission adopts a decision pursuant to Article 45(3) of Regulation (EU) 2016/679 that covers the transfer of personal data to which these Clauses apply; or (ii) Regulation (EU) 2016/679 becomes part of the legal framework of the country to which the personal data is transferred. This is without prejudice to other obligations applying to the processing in question under Regulation (EU) 2016/679.
​

***Clause 17:* Governing law**

These Clauses shall be governed by the law of one of the EU Member States, provided such law allows for third-party beneficiary rights. The Parties agree that this shall be the law of \_\_\_\_\_\_ (*specify Member State*).
​

***Clause 18*: Choice of forum and jurisdiction**

\(a\) Any dispute arising from these Clauses shall be resolved by the courts of an EU Member State.

\(b\) The Parties agree that those shall be the courts of \_\_\_\_\_ (*specify Member State*).

\(c\) A data subject may also bring legal proceedings against the data exporter and/or data importer before the courts of the Member State in which he/she has his/her habitual residence.

\(d\) The Parties agree to submit themselves to the jurisdiction of such courts

​

**ANNEX IA. LIST OF PARTIESData exporter(s):**

Name: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Address: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Contact person’s name, position and contact details: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Activities relevant to the data transferred under these Clauses: Use of the Disqus Commenting Platform

Signature and date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Role (controller/processor): Co-Controller

**Data importer(s):**

Name: Disqus, Inc.

Address: 3 Park Avenue, 33rd Floor, New York, NY 10016

Contact person’s name, position and contact details: Steven Stein, steven.stein@disqus.com

Activities relevant to the data transferred under these Clauses: Providing the Disqus Commenting Platform.

Signature and date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Role (controller/processor): Co-Controller

​

**B. DESCRIPTION OF TRANSFER**

*Categories of data subjects whose personal data is transferred*

Users that use the Disqus comment function on a publisher’s website

*Categories of personal data transferred*

Email addresses, username, IP address, other online identifier, information revealed in user comments.

*Sensitive data transferred (if applicable) and applied restrictions or safeguards that fully take into consideration the nature of the data and the risks involved, such as for instance strict purpose limitation, access restrictions (including access only for staff having followed specialised training), keeping a record of access to the data, restrictions for onward transfers or additional security measures.*

NA

*The frequency of the transfer*

Depends on user’s use of the Disqus comment function

*Nature of the processing*

Disqus provides a commenting application service (“Disqus Comments”) to publisher for use as a comment forum on the publisher’s website. Disqus collects personal data from users commenting in the Disqus Comments on publisher’s website. Disqus provides publisher with access to the comments so that publisher may act as moderator on its website, and to meet its relevant obligations under applicable laws.

*Purpose(s) of the data transfer and further processing*

To fulfill user’s request to comment on publisher’s website.

*The period for which the personal data will be retained, or, if that is not possible, the criteria used to determine that period*

To fulfill the purposes of processing personal data.

*For transfers to (sub-) processors, also specify subject matter, nature and duration of the processing*

**C. COMPETENT SUPERVISORY AUTHORITYANNEX IITECHNICAL AND ORGANISATIONAL MEASURES INCLUDING TECHNICAL AND ORGANISATIONAL MEASURES TO ENSURE THE SECURITY OF THE DATAA. Technical Measures1. Information Security Policy**

1.1. Disqus maintains a written information security policy which shall include, at a minimum, the approach adopted by Disqus to address the confidentiality, integrity, and availability of Disqus, its affiliates’ and representatives’, and its customers’ confidential information, as applicable, which at least meets the minimum standards of (a) International Standard ISO.IEC 27001 and 27002 or (b) a similar industry-standard framework.

**2. Testing and Scanning Procedures**

2.1. Penetration testing: Upon request, Disqus provides Publisher with an executive summary of the results of penetration testing.

2.2. Information security certification: Upon request, Disqus provides Publisher with a third-party information security certification such as ISO 27K and SOC 2 Type II from an industry-recognized third party as of such then-completed year.

2.3. Vulnerability scanning: Upon request, Disqus provides Publisher with executive summaries of external, internal, and web application vulnerability scanning. If Publisher identifies elevated risks present within provided information, Disqus will promptly remediate identified risks at Disqus’ expense. Disqus adheres to OWASP coding principles for web application code and uses static or dynamic application vulnerability scanning or manual code review, when appropriate. Disqus ensures that vulnerability scans periodically are conducted on devices and software present in its internal and external network environments and web applications to identify and remediate or have remediated any vulnerabilities within a documented timeframe. Disqus will provide Publisher with summaries of such scans and results upon request.

2.4. Penetration tests: Disqus conducts annual penetration tests on all its externally facing critical systems and applications from an industry-recognized independent third party at Disqus ‘expense. Disqus conducts similar such tests after significant changes are made to its network. Such penetration tests shall:

\(a\) be based on industry-accepted penetration testing approaches (e.g. NIST SP800-115),

\(b\) include testing from inside and outside the network,

\(c\) include testing to validate segmentation, and

\(d\) include network-layer, operating system, and application layer testing such that, at a minimum, they test against vulnerabilities identified in industry standards (e.g. OWASP Guide, SANS CWE Top 25, CERT Secure Coding).

**3. Backups**

Disqus maintains secure, usable, and traceable data/information backups to ensure that backups can be used when necessary.

**4. Internal Hardware Protection**

4.1. Disqus ensures that all computing and storage devices on its network, including but not limited to, workstations, servers, and network devices have endpoint protection in place consistent with industry standards, such as anti-malware, email, and application scanning, or antivirus.

4.2. Disqus has a tiered network architecture, which includes preventive and detective devices and where highly sensitive non-public information is in a secured and segregated network.

4.3. Disqus ensures that the network devices, servers, and workstations where Publisher information is located are hardened and continuously subject to minimum security baselines.

4.4. Disqus maintains an inventory of authorized devices that can be connected to its network environment and ensures that such inventory is reconciled periodically.

4.5. Disqus maintains an inventory of authorized software required for its network devices, servers, and workstations present in its network environment and ensures that such inventory is reconciled periodically.

4.6. Disqus maintains a change management methodology that ensures only approved changes are released and deployed in the production environment.

4.7. Disqus conducts periodic reviews of its cloud computing use based on the cloud security alliance risks and controls structure and addresses any elevated risks identified in a timely manner.

**5. Periodic Reviews and Updates**

5.1. Disqus conducts periodic reviews of its cloud computing use based on the cloud security alliance risks and controls structure and addresses any elevated risks identified in a timely manner.

5.2. Disqus promptly applies the latest firmware/security patches and updates on devices and software present in its network environment, expediting the application of critical and high-risk security patches and updates.

**6. Encryption**

6.1. Disqus ensures that all communications being initiated by Disqus or handling sensitive data are encrypted using industry-standard secure protocols.

**B. Physical Security Measures7. Data Center Security Measures**

Disqus ensures appropriate data center physical security and data center environmental controls. Disqus utilizes the following physical security measures:

7.1. a closed-circuit television monitoring system with redundant power sources that provides recognizable images and usable recordings of entrances, exits, loading docks, and other high-security areas, and which maintains all images for at least 30 days and incident images indefinitely; The media portion is kept in a secured area.

7.2. distribution logs and for all issued access devices (including keys), secured storage areas for unissued devices, and regular audits of each of the foregoing;

7.3. access control alarms that actively are monitored by appropriate personnel;

7.4. identification (relying on governmentally issued credentials) and logging of individuals accessing Disqus’ facilities (including visitors), as well as restriction of access to Publisher’s assets (including intangible assets) to individuals authorized by Publisher.

**C. Organizational Measures8. Internal Employee Procedures and Policies**

8.1. Role-based access control: Disqus maintains an up-to-date role-based access control based on data classification and job roles of employees, using the principle of least privilege and granting access only on a need-to-know basis.

8.2. Segregation of duties: Disqus maintains a segregation of duties, such that individuals performing application development are different from individuals managing production environments. Disqus employs technical and procedural controls to prevent developers and system administrators from obtaining access to production information.

8.3. Background checks: Disqus only utilizes personnel, including employees, contractors, and subcontractors, after performing background checks on them.

8.4. Security awareness training: Disqus ensures that its employees, contractors, and subcontractors receive appropriate security awareness training on a periodic basis.

8.5. Publisher’s Standard of Conduct: If non-escorted access or access to Publisher systems is required, Disqus causes its representatives to comply with Publisher’s standard of conduct.

**9. Incident Response Plan**

9.1. Disqus maintains an incident response plan that ensures that Disqus is adequately prepared to handle an incident, is able to accurately identify a Security Event as an incident, is able to contain the impact of the incident, has procedures in place to remediate the incident, has the ability to successfully recover from an incident and performs a root cause analysis of the incident.

**10. Security Event**

10.1. “Security Event” shall mean an instance of Disqus learning or having reason to believe that Publisher’s confidential information has been accessed by an unauthorized person or disclosed in a manner not permitted by Disqus’ agreement with Publisher, or that an incursion in any systems, processes, hardware or software used to store, transmit or that otherwise affect Publisher’s confidential information has occurred.

10.2. In a Security Event, Disqus will:

10.2.1. as soon as reasonably practicable and in no event more than seventy-two hours after becoming aware of such Security Event, provide details of the same to Publisher including (i) the date of the Security Event, (ii) details concerning the data compromised, (iii) the method of the Security Event, (iv) appropriate Disqus security personnel contacts and security personnel contacts of its representatives, (v) the name of any person or entity assisting Disqus with the investigation of the suspected or actual Security Event, (vi) a list of all parties known to have gained unauthorized access to confidential information for the limited purpose of assessing Publisher’s exposure and (vii) any other information which Publisher reasonably requests from Disqus or its representatives concerning such suspected or Security Event, including any forensics reports;

10.2.2. grant access to Publisher’s representatives or another person or entity agreed to by Publisher and Disqus (with each acting in good faith in the selection of such other person or entity) to Disqus’ systems and premises to allow such representatives or such other person or entity to perform an investigation (including the installation of any monitoring or diagnostic software) deemed necessary by Publisher to locate the source of such breach; and

10.2.3. immediately take appropriate steps to ensure that any actual data security breach does not continue.

10.3. Public Authorities: Disqus will not notify law enforcement or federal or state regulatory authorities of any Security Event or other matter related to Publisher’s security requirements without prior notice to Publisher unless otherwise required by applicable law.

10.4. Press Releases: Disqus will not issue any press release or other public announcement concerning a Security Event without prior approval of Publisher.

10.5. Read-only logs: Disqus maintains usable read-only logs of critical systems and events on network devices, key security systems, and workstation and server operating systems, and ensures that any suspicious activity is monitored and investigated and appropriate actions are taken subsequent to its detection.

**11. Data Classification**

11.1. Classes: Disqus defines classes of data/information based on applicable legal requirements and sensitivity levels of the related data/information and treats such data/information according to that classification.

**12. Collaboration with Publisher**

12.1. Publisher’s security requirements: If Publisher believes that Disqus’ or any of its representatives’ security procedures in connection with the services provided to Publisher do not comply with Publisher’s security requirements, Disqus will cooperate with Publisher to ensure that security measures and procedures that comply with Publisher’s security requirements are promptly implemented.

12.2. Notification: Disqus will promptly notify Publisher if Disqus learns or has reason to believe that it or any of its representatives are not in compliance with any of Publisher’s security requirements, whether or not a Security Event has occurred.

​

#### **STANDARD CONTRACTUAL CLAUSES**

Controller to Processor

**SECTION I*Clause 1:* Purpose and scope**

\(a\) The purpose of these standard contractual clauses is to ensure compliance with the requirements of Regulation (EU) 2016/679 of the European Parliament and of the Council of 27 April 2016 on the protection of natural persons with regard to the processing of personal data and on the free movement of such data (General Data Protection Regulation) for the transfer of data to a third country.

\(b\) The Parties:

\(i\) the natural or legal person(s), public authority/ies, agency/ies or other body/ies (hereinafter ‘entity/ies’) transferring the personal data, as listed in Annex I.A (hereinafter each ‘data exporter’), and

\(ii\) the entity/ies in a third country receiving the personal data from the data exporter, directly or indirectly via another entity also Party to these Clauses, as listed in Annex I.A (hereinafter each ‘data importer’)

have agreed to these standard contractual clauses (hereinafter: ‘Clauses’).

\(c\) These Clauses apply with respect to the transfer of personal data as specified in Annex I.B.

\(d\) The Appendix to these Clauses containing the Annexes referred to therein forms an integral part of these Clauses.

***Clause 2:* Effect and invariability of the Clauses**

\(a\) These Clauses set out appropriate safeguards, including enforceable data subject rights and effective legal remedies, pursuant to Article 46(1) and Article 46(2)(c) of Regulation (EU) 2016/679 and, with respect to data transfers from controllers to processors and/or processors to processors, standard contractual clauses pursuant to Article 28(7) of Regulation (EU) 2016/679, provided they are not modified, except to select the appropriate Module(s) or to add or update information in the Appendix. This does not prevent the Parties from including the standard contractual clauses laid down in these Clauses in a wider contract and/or to add other clauses or additional safeguards, provided that they do not contradict, directly or indirectly, these Clauses or prejudice the fundamental rights or freedoms of data subjects.

\(b\) These Clauses are without prejudice to obligations to which the data exporter is subject by virtue of Regulation (EU) 2016/679.

***Clause 3:* Third-party beneficiaries**

\(a\) Data subjects may invoke and enforce these Clauses, as third-party beneficiaries, against the data exporter and/or data importer, with the following exceptions:

\(i\) Clause 1, Clause 2, Clause 3, Clause 6, Clause 7;

\(ii\) Clause 8.1(b), 8.9(a), (c), (d) and (e);

\(iii\) Clause 9(a), (c), (d) and (e);

\(iv\) Clause 12(a), (d) and (f);

\(v\) Clause 13;

\(vi\) Clause 15.1(c), (d) and (e);

\(vii\) Clause 16(e);

\(viii\) Clause 18(a) and (b).

\(b\) Paragraph (a) is without prejudice to rights of data subjects under Regulation (EU) 2016/679.

***Clause 4:* Interpretation**

\(a\) Where these Clauses use terms that are defined in Regulation (EU) 2016/679, those terms shall have the same meaning as in that Regulation.

\(b\) These Clauses shall be read and interpreted in the light of the provisions of Regulation (EU) 2016/679.

\(c\) These Clauses shall not be interpreted in a way that conflicts with rights and obligations provided for in Regulation (EU) 2016/679.

***Clause 5:* Hierarchy**

In the event of a contradiction between these Clauses and the provisions of related agreements between the Parties, existing at the time these Clauses are agreed or entered into thereafter, these Clauses shall prevail.

***Clause 6:* Description of the transfer(s)**

The details of the transfer(s), and in particular the categories of personal data that are transferred and the purpose(s) for which they are transferred, are specified in Annex I.B.

***Clause 7 – Optional:* Docking clause**

\(a\) An entity that is not a Party to these Clauses may, with the agreement of the Parties, accede to these Clauses at any time, either as a data exporter or as a data importer, by completing the Appendix and signing Annex I.A.

\(b\) Once it has completed the Appendix and signed Annex I.A, the acceding entity shall become a Party to these Clauses and have the rights and obligations of a data exporter or data importer in accordance with its designation in Annex I.A.

\(c\) The acceding entity shall have no rights or obligations arising under these Clauses from the period prior to becoming a Party.

**SECTION II – OBLIGATIONS OF THE PARTIES*Clause 8:* Data protection safeguards**

The data exporter warrants that it has used reasonable efforts to determine that the data importer is able, through the implementation of appropriate technical and organisational measures, to satisfy its obligations under these Clauses.

**8.1 Instructions**

\(a\) The data importer shall process the personal data only on documented instructions from the data exporter. The data exporter may give such instructions throughout the duration of the contract.

\(b\) The data importer shall immediately inform the data exporter if it is unable to follow those instructions.

**8.2 Purpose limitation**

The data importer shall process the personal data only for the specific purpose(s) of the transfer, as set out in Annex I.B, unless on further instructions from the data exporter.

**8.3 Transparency**

On request, the data exporter shall make a copy of these Clauses, including the Appendix as completed by the Parties, available to the data subject free of charge. To the extent necessary to protect business secrets or other confidential information, including the measures described in Annex II and personal data, the data exporter may redact part of the text of the Appendix to these Clauses prior to sharing a copy, but shall provide a meaningful summary where the data subject would otherwise not be able to understand the its content or exercise his/her rights. On request, the Parties shall provide the data subject with the reasons for the redactions, to the extent possible without revealing the redacted information. This Clause is without prejudice to the obligations of the data exporter under Articles 13 and 14 of Regulation (EU) 2016/679.

**8.4 Accuracy**

If the data importer becomes aware that the personal data it has received is inaccurate, or has become outdated, it shall inform the data exporter without undue delay. In this case, the data importer shall cooperate with the data exporter to erase or rectify the data.

**8.5 Duration of processing and erasure or return of data**

Processing by the data importer shall only take place for the duration specified in Annex I.B. After the end of the provision of the processing services, the data importer shall, at the choice of the data exporter, delete all personal data processed on behalf of the data exporter and certify to the data exporter that it has done so, or return to the data exporter all personal data processed on its behalf and delete existing copies. Until the data is deleted or returned, the data importer shall continue to ensure compliance with these Clauses. In case of local laws applicable to the data importer that prohibit return or deletion of the personal data, the data importer warrants that it will continue to ensure compliance with these Clauses and will only process it to the extent and for as long as required under that local law. This is without prejudice to Clause 14, in particular the requirement for the data importer under Clause 14(e) to notify the data exporter throughout the duration of the contract if it has reason to believe that it is or has become subject to laws or practices not in line with the requirements under Clause 14(a).

**8.6 Security of processing**

\(a\) The data importer and, during transmission, also the data exporter shall implement appropriate technical and organisational measures to ensure the security of the data, including protection against a breach of security leading to accidental or unlawful destruction, loss, alteration, unauthorised disclosure or access to that data (hereinafter ‘personal data breach’). In assessing the appropriate level of security, the Parties shall take due account of the state of the art, the costs of implementation, the nature, scope, context and purpose(s) of processing and the risks involved in the processing for the data subjects. The Parties shall in particular consider having recourse to encryption or pseudonymisation, including during transmission, where the purpose of processing can be fulfilled in that manner. In case of pseudonymisation, the additional information for attributing the personal data to a specific data subject shall, where possible, remain under the exclusive control of the data exporter. In complying with its obligations under this paragraph, the data importer shall at least implement the technical and organisational measures specified in Annex II. The data importer shall carry out regular checks to ensure that these measures continue to provide an appropriate level of security.

\(b\) The data importer shall grant access to the personal data to members of its personnel only to the extent strictly necessary for the implementation, management and monitoring of the contract. It shall ensure that persons authorised to process the personal data have committed themselves to confidentiality or are under an appropriate statutory obligation of confidentiality.

\(c\) In the event of a personal data breach concerning personal data processed by the data importer under these Clauses, the data importer shall take appropriate measures to address the breach, including measures to mitigate its adverse effects. The data importer shall also notify the data exporter without undue delay after having become aware of the breach. Such notification shall contain the details of a contact point where more information can be obtained, a description of the nature of the breach (including, where possible, categories and approximate number of data subjects and personal data records concerned), its likely consequences and the measures taken or proposed to address the breach including, where appropriate, measures to mitigate its possible adverse effects. Where, and in so far as, it is not possible to provide all information at the same time, the initial notification shall contain the information then available and further information shall, as it becomes available, subsequently be provided without undue delay.

\(d\) The data importer shall cooperate with and assist the data exporter to enable the data exporter to comply with its obligations under Regulation (EU) 2016/679, in particular to notify the competent supervisory authority and the affected data subjects, taking into account the nature of processing and the information available to the data importer.

**8.7 Sensitive data**

Where the transfer involves personal data revealing racial or ethnic origin, political opinions, religious or philosophical beliefs, or trade union membership, genetic data, or biometric data for the purpose of uniquely identifying a natural person, data concerning health or a person’s sex life or sexual orientation, or data relating to criminal convictions and offences (hereinafter ‘sensitive data’), the data importer shall apply the specific restrictions and/or additional safeguards described in Annex I.B.

**8.8 Onward transfers**

The data importer shall only disclose the personal data to a third party on documented instructions from the data exporter. In addition, the data may only be disclosed to a third party located outside the European Union (in the same country as the data importer or in another third country, hereinafter ‘onward transfer’) if the third party is or agrees to be bound by these Clauses, under the appropriate Module, or if:

\(i\) the onward transfer is to a country benefitting from an adequacy decision pursuant to Article 45 of Regulation (EU) 2016/679 that covers the onward transfer;

\(ii\) the third party otherwise ensures appropriate safeguards pursuant to Articles 46 or 47 Regulation of (EU) 2016/679 with respect to the processing in question;

\(iii\) the onward transfer is necessary for the establishment, exercise or defence of legal claims in the context of specific administrative, regulatory or judicial proceedings; or

\(iv\) the onward transfer is necessary in order to protect the vital interests of the data subject or of another natural person.

Any onward transfer is subject to compliance by the data importer with all the other safeguards under these Clauses, in particular purpose limitation.

**8.9 Documentation and compliance**

\(a\) The data importer shall promptly and adequately deal with enquiries from the data exporter that relate to the processing under these Clauses.

\(b\) The Parties shall be able to demonstrate compliance with these Clauses. In particular, the data importer shall keep appropriate documentation on the processing activities carried out on behalf of the data exporter.

\(c\) The data importer shall make available to the data exporter all information necessary to demonstrate compliance with the obligations set out in these Clauses and at the data exporter’s request, allow for and contribute to audits of the processing activities covered by these Clauses, at reasonable intervals or if there are indications of non-compliance. In deciding on a review or audit, the data exporter may take into account relevant certifications held by the data importer.

\(d\) The data exporter may choose to conduct the audit by itself or mandate an independent auditor. Audits may include inspections at the premises or physical facilities of the data importer and shall, where appropriate, be carried out with reasonable notice.

\(e\) The Parties shall make the information referred to in paragraphs (b) and (c), including the results of any audits, available to the competent supervisory authority on request.

***Clause 9*: Use of sub-processors**

\(a\) GENERAL WRITTEN AUTHORISATION The data importer has the data exporter’s general authorisation for the engagement of sub-processor(s) from an agreed list. The data importer shall specifically inform the data exporter in writing of any intended changes to that list through the addition or replacement of sub-processors at least \[*Specify time period*\] in advance, thereby giving the data exporter sufficient time to be able to object to such changes prior to the engagement of the sub-processor(s). The data importer shall provide the data exporter with the information necessary to enable the data exporter to exercise its right to object.

\(b\) Where the data importer engages a sub-processor to carry out specific processing activities (on behalf of the data exporter), it shall do so by way of a written contract that provides for, in substance, the same data protection obligations as those binding the data importer under these Clauses, including in terms of third-party beneficiary rights for data subjects. The Parties agree that, by complying with this Clause, the data importer fulfils its obligations under Clause 8.8. The data importer shall ensure that the sub-processor complies with the obligations to which the data importer is subject pursuant to these Clauses.

\(c\) The data importer shall provide, at the data exporter’s request, a copy of such a sub-processor agreement and any subsequent amendments to the data exporter. To the extent necessary to protect business secrets or other confidential information, including personal data, the data importer may redact the text of the agreement prior to sharing a copy.

\(d\) The data importer shall remain fully responsible to the data exporter for the performance of the sub-processor’s obligations under its contract with the data importer. The data importer shall notify the data exporter of any failure by the sub-processor to fulfil its obligations under that contract.

\(e\) The data importer shall agree a third-party beneficiary clause with the sub-processor whereby – in the event the data importer has factually disappeared, ceased to exist in law or has become insolvent – the data exporter shall have the right to terminate the sub-processor contract and to instruct the sub-processor to erase or return the personal data.

***Clause 10*: Data subject rights**

\(a\) The data importer shall promptly notify the data exporter of any request it has received from a data subject. It shall not respond to that request itself unless it has been authorised to do so by the data exporter.

\(b\) The data importer shall assist the data exporter in fulfilling its obligations to respond to data subjects’ requests for the exercise of their rights under Regulation (EU) 2016/679. In this regard, the Parties shall set out in Annex II the appropriate technical and organisational measures, taking into account the nature of the processing, by which the assistance shall be provided, as well as the scope and the extent of the assistance required.

\(c\) In fulfilling its obligations under paragraphs (a) and (b), the data importer shall comply with the instructions from the data exporter.

***Clause 11*: Redress**

\(a\) The data importer shall inform data subjects in a transparent and easily accessible format, through individual notice or on its website, of a contact point authorised to handle complaints. It shall deal promptly with any complaints it receives from a data subject.

\[OPTION: The data importer agrees that data subjects may also lodge a complaint with an independent dispute resolution body at no cost to the data subject. It shall inform the data subjects, in the manner set out in paragraph (a), of such redress mechanism and that they are not required to use it, or follow a particular sequence in seeking redress.\]

\(b\) In case of a dispute between a data subject and one of the Parties as regards compliance with these Clauses, that Party shall use its best efforts to resolve the issue amicably in a timely fashion. The Parties shall keep each other informed about such disputes and, where appropriate, cooperate in resolving them.

\(c\) Where the data subject invokes a third-party beneficiary right pursuant to Clause 3, the data importer shall accept the decision of the data subject to:

\(i\) lodge a complaint with the supervisory authority in the Member State of his/her habitual residence or place of work, or the competent supervisory authority pursuant to Clause 13;

\(ii\) refer the dispute to the competent courts within the meaning of Clause 18.

\(d\) The Parties accept that the data subject may be represented by a not-for-profit body, organisation or association under the conditions set out in Article 80(1) of Regulation (EU) 2016/679.

\(e\) The data importer shall abide by a decision that is binding under the applicable EU or Member State law.

\(f\) The data importer agrees that the choice made by the data subject will not prejudice his/her substantive and procedural rights to seek remedies in accordance with applicable laws.

***Clause 12*: Liability**

\(a\) Each Party shall be liable to the other Party/ies for any damages it causes the other Party/ies by any breach of these Clauses.

\(b\) The data importer shall be liable to the data subject, and the data subject shall be entitled to receive compensation, for any material or non-material damages the data importer or its sub-processor causes the data subject by breaching the third-party beneficiary rights under these Clauses.

\(c\) Notwithstanding paragraph (b), the data exporter shall be liable to the data subject, and the data subject shall be entitled to receive compensation, for any material or non-material damages the data exporter or the data importer (or its sub-processor) causes the data subject by breaching the third-party beneficiary rights under these Clauses. This is without prejudice to the liability of the data exporter and, where the data exporter is a processor acting on behalf of a controller, to the liability of the controller under Regulation (EU) 2016/679 or Regulation (EU) 2018/1725, as applicable.

\(d\) The Parties agree that if the data exporter is held liable under paragraph (c) for damages caused by the data importer (or its sub-processor), it shall be entitled to claim back from the data importer that part of the compensation corresponding to the data importer’s responsibility for the damage.

\(e\) Where more than one Party is responsible for any damage caused to the data subject as a result of a breach of these Clauses, all responsible Parties shall be jointly and severally liable and the data subject is entitled to bring an action in court against any of these Parties.

\(f\) The Parties agree that if one Party is held liable under paragraph (e), it shall be entitled to claim back from the other Party/ies that part of the compensation corresponding to its/their responsibility for the damage.

\(g\) The data importer may not invoke the conduct of a sub-processor to avoid its own liability.

***Clause 13:* Supervision**

\(c\) \[Where the data exporter is established in an EU Member State:\] The supervisory authority with responsibility for ensuring compliance by the data exporter with Regulation (EU) 2016/679 as regards the data transfer, as indicated in Annex I.C, shall act as competent supervisory authority.

\[Where the data exporter is not established in an EU Member State, but falls within the territorial scope of application of Regulation (EU) 2016/679 in accordance with its Article 3(2) and has appointed a representative pursuant to Article 27(1) of Regulation (EU) 2016/679:\] The supervisory authority of the Member State in which the representative within the meaning of Article 27(1) of Regulation (EU) 2016/679 is established, as indicated in Annex I.C, shall act as competent supervisory authority.

\[Where the data exporter is not established in an EU Member State, but falls within the territorial scope of application of Regulation (EU) 2016/679 in accordance with its Article 3(2) without however having to appoint a representative pursuant to Article 27(2) of Regulation (EU) 2016/679:\] The supervisory authority of one of the Member States in which the data subjects whose personal data is transferred under these Clauses in relation to the offering of goods or services to them, or whose behaviour is monitored, are located, as indicated in Annex I.C, shall act as competent supervisory authority.

\(b\) The data importer agrees to submit itself to the jurisdiction of and cooperate with the competent supervisory authority in any procedures aimed at ensuring compliance with these Clauses. In particular, the data importer agrees to respond to enquiries, submit to audits and comply with the measures adopted by the supervisory authority, including remedial and compensatory measures. It shall provide the supervisory authority with written confirmation that the necessary actions have been taken.

**SECTION III – LOCAL LAWS AND OBLIGATIONS IN CASE OF ACCESS BY PUBLIC AUTHORITIES*Clause 14:* Local laws and practices affecting compliance with the Clauses**

\(a\) The Parties warrant that they have no reason to believe that the laws and practices in the third country of destination applicable to the processing of the personal data by the data importer, including any requirements to disclose personal data or measures authorising access by public authorities, prevent the data importer from fulfilling its obligations under these Clauses. This is based on the understanding that laws and practices that respect the essence of the fundamental rights and freedoms and do not exceed what is necessary and proportionate in a democratic society to safeguard one of the objectives listed in Article 23(1) of Regulation (EU) 2016/679, are not in contradiction with these Clauses.

\(b\) The Parties declare that in providing the warranty in paragraph (a), they have taken due account in particular of the following elements:

\(i\) the specific circumstances of the transfer, including the length of the processing chain, the number of actors involved and the transmission channels used; intended onward transfers; the type of recipient; the purpose of processing; the categories and format of the transferred personal data; the economic sector in which the transfer occurs; the storage location of the data transferred;

\(ii\) the laws and practices of the third country of destination– including those requiring the disclosure of data to public authorities or authorising access by such authorities – relevant in light of the specific circumstances of the transfer, and the applicable limitations and safeguards;

\(iii\) any relevant contractual, technical or organisational safeguards put in place to supplement the safeguards under these Clauses, including measures applied during transmission and to the processing of the personal data in the country of destination.

\(c\) The data importer warrants that, in carrying out the assessment under paragraph (b), it has made its best efforts to provide the data exporter with relevant information and agrees that it will continue to cooperate with the data exporter in ensuring compliance with these Clauses.

\(d\) The Parties agree to document the assessment under paragraph (b) and make it available to the competent supervisory authority on request.

\(e\) The data importer agrees to notify the data exporter promptly if, after having agreed to these Clauses and for the duration of the contract, it has reason to believe that it is or has become subject to laws or practices not in line with the requirements under paragraph (a), including following a change in the laws of the third country or a measure (such as a disclosure request) indicating an application of such laws in practice that is not in line with the requirements in paragraph (a).

\(f\) Following a notification pursuant to paragraph (e), or if the data exporter otherwise has reason to believe that the data importer can no longer fulfil its obligations under these Clauses, the data exporter shall promptly identify appropriate measures (e.g. technical or organisational measures to ensure security and confidentiality) to be adopted by the data exporter and/or data importer to address the situation. The data exporter shall suspend the data transfer if it considers that no appropriate safeguards for such transfer can be ensured, or if instructed by the competent supervisory authority to do so. In this case, the data exporter shall be entitled to terminate the contract, insofar as it concerns the processing of personal data under these Clauses. If the contract involves more than two Parties, the data exporter may exercise this right to termination only with respect to the relevant Party, unless the Parties have agreed otherwise. Where the contract is terminated pursuant to this Clause, Clause 16(d) and (e) shall apply.

***Clause 15:* Obligations of the data importer in case of access by public authorities15.1 Notification**

\(a\) The data importer agrees to notify the data exporter and, where possible, the data subject promptly (if necessary with the help of the data exporter) if it:

\(i\) receives a legally binding request from a public authority, including judicial authorities, under the laws of the country of destination for the disclosure of personal data transferred pursuant to these Clauses; such notification shall include information about the personal data requested, the requesting authority, the legal basis for the request and the response provided; or

\(ii\) becomes aware of any direct access by public authorities to personal data transferred pursuant to these Clauses in accordance with the laws of the country of destination; such notification shall include all information available to the importer.

\(b\) If the data importer is prohibited from notifying the data exporter and/or the data subject under the laws of the country of destination, the data importer agrees to use its best efforts to obtain a waiver of the prohibition, with a view to communicating as much information as possible, as soon as possible. The data importer agrees to document its best efforts in order to be able to demonstrate them on request of the data exporter.

\(c\) Where permissible under the laws of the country of destination, the data importer agrees to provide the data exporter, at regular intervals for the duration of the contract, with as much relevant information as possible on the requests received (in particular, number of requests, type of data requested, requesting authority/ies, whether requests have been challenged and the outcome of such challenges, etc.).

\(d\) The data importer agrees to preserve the information pursuant to paragraphs (a) to (c) for the duration of the contract and make it available to the competent supervisory authority on request.

\(e\) Paragraphs (a) to (c) are without prejudice to the obligation of the data importer pursuant to Clause 14(e) and Clause 16 to inform the data exporter promptly where it is unable to comply with these Clauses.

**15.2 Review of legality and data minimisation**

\(a\) The data importer agrees to review the legality of the request for disclosure, in particular whether it remains within the powers granted to the requesting public authority, and to challenge the request if, after careful assessment, it concludes that there are reasonable grounds to consider that the request is unlawful under the laws of the country of destination, applicable obligations under international law and principles of international comity. The data importer shall, under the same conditions, pursue possibilities of appeal. When challenging a request, the data importer shall seek interim measures with a view to suspending the effects of the request until the competent judicial authority has decided on its merits. It shall not disclose the personal data requested until required to do so under the applicable procedural rules. These requirements are without prejudice to the obligations of the data importer under Clause 14(e).

\(b\) The data importer agrees to document its legal assessment and any challenge to the request for disclosure and, to the extent permissible under the laws of the country of destination, make the documentation available to the data exporter. It shall also make it available to the competent supervisory authority on request.

\(c\) The data importer agrees to provide the minimum amount of information permissible when responding to a request for disclosure, based on a reasonable interpretation of the request.

**SECTION IV – FINAL PROVISIONS*Clause 16:* Non-compliance with the Clauses and termination**

\(a\) The data importer shall promptly inform the data exporter if it is unable to comply with these Clauses, for whatever reason.

\(b\) In the event that the data importer is in breach of these Clauses or unable to comply with these Clauses, the data exporter shall suspend the transfer of personal data to the data importer until compliance is again ensured or the contract is terminated. This is without prejudice to Clause 14(f).

\(c\) The data exporter shall be entitled to terminate the contract, insofar as it concerns the processing of personal data under these Clauses, where:

\(i\) the data exporter has suspended the transfer of personal data to the data importer pursuant to paragraph (b) and compliance with these Clauses is not restored within a reasonable time and in any event within one month of suspension;

\(ii\) the data importer is in substantial or persistent breach of these Clauses; or

\(iii\) the data importer fails to comply with a binding decision of a competent court or supervisory authority regarding its obligations under these Clauses.

In these cases, it shall inform the competent supervisory authority of such non-compliance. Where the contract involves more than two Parties, the data exporter may exercise this right to termination only with respect to the relevant Party, unless the Parties have agreed otherwise.

\(d\) Personal data that has been transferred prior to the termination of the contract pursuant to paragraph (c) shall at the choice of the data exporter immediately be returned to the data exporter or deleted in its entirety. The same shall apply to any copies of the data. The data importer shall certify the deletion of the data to the data exporter. Until the data is deleted or returned, the data importer shall continue to ensure compliance with these Clauses. In case of local laws applicable to the data importer that prohibit the return or deletion of the transferred personal data, the data importer warrants that it will continue to ensure compliance with these Clauses and will only process the data to the extent and for as long as required under that local law.

\(e\) Either Party may revoke its agreement to be bound by these Clauses where (i) the European Commission adopts a decision pursuant to Article 45(3) of Regulation (EU) 2016/679 that covers the transfer of personal data to which these Clauses apply; or (ii) Regulation (EU) 2016/679 becomes part of the legal framework of the country to which the personal data is transferred. This is without prejudice to other obligations applying to the processing in question under Regulation (EU) 2016/679.

***Clause 17*: Governing law**

These Clauses shall be governed by the law of one of the EU Member States, provided such law allows for third-party beneficiary rights. The Parties agree that this shall be the law of \_\_\_\_\_\_\_ (*specify Member State*).

***Clause 18*: Choice of forum and jurisdiction**

\(a\) Any dispute arising from these Clauses shall be resolved by the courts of an EU Member State.

\(b\) The Parties agree that those shall be the courts of \_\_\_\_\_ (*specify Member State*).

\(c\) A data subject may also bring legal proceedings against the data exporter and/or data importer before the courts of the Member State in which he/she has his/her habitual residence.

\(d\) The Parties agree to submit themselves to the jurisdiction of such courts

​

**ANNEX IA. LIST OF PARTIESData exporter(s):**

Name: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Address: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Contact person’s name, position and contact details: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Activities relevant to the data transferred under these Clauses: Use of the Disqus Commenting Platform.

Signature and date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Role (controller): Controller

**Data importer(s):**

Name: Disqus, Inc.

Address: 3 Park Avenue, 33rd Floor, New York, NY 10016

Contact person’s name, position and contact details: Steven Stein, steven.stein@disqus.com

Activities relevant to the data transferred under these Clauses: Providing the Disqus Commenting Platform.

Signature and date: \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_

Role (controller/processor): Processor

​

**B. DESCRIPTION OF TRANSFER**

*Categories of data subjects whose personal data is transferred:* Users that use the Disqus comment function on a publisher’s website

*Categories of personal data transferred*: SSO user data

*Sensitive data transferred (if applicable) and applied restrictions or safeguards that fully take into consideration the nature of the data and the risks involved, such as for instance strict purpose limitation, access restrictions (including access only for staff having followed specialised training), keeping a record of access to the data, restrictions for onward transfers or additional security measures.*

NA

*The frequency of the transfer:* Depends on user’s use of the Disqus comment function.

*Nature of the processing*

Disqus provides a commenting application service (“Disqus Comments”) to publisher for use as a comment forum on the publisher’s website. Disqus collects personal data from users commenting in the Disqus Comments on publisher’s website. Disqus provides publisher with access to the comments so that publisher may act as moderator on its website, and to meet its relevant obligations under applicable laws.

*Purpose(s) of the data transfer and further processing:* To fulfil user’s request to comment on publisher’s website.

*The period for which the personal data will be retained, or, if that is not possible, the criteria used to determine that period:* As set forth in the Agreement.

**C. COMPETENT SUPERVISORY AUTHORITY**

​

**ANNEX IITECHNICAL AND ORGANISATIONAL MEASURES INCLUDING TECHNICAL AND ORGANISATIONAL MEASURES TO ENSURE THE SECURITY OF THE DATAA. Technical Measures1. Information Security Policy**

1.1 Disqus maintains a written information security policy which shall include, at a minimum, the approach adopted by Disqus to address the confidentiality, integrity, and availability of Disqus, its affiliates’ and representatives’, and its customers’ confidential information, as applicable, which at least meets the minimum standards of (a) International Standard ISO.IEC 27001 and 27002 or (b) a similar industry-standard framework.

**2. Testing and Scanning Procedures**

2.1 Penetration testing: Upon request, Disqus provides Publisher with an executive summary of the results of penetration testing.

2.2 Information security certification: Upon request, Disqus provides Publisher with a third-party information security certification such as ISO 27K and SOC 2 Type II from an industry-recognized third party as of such then-completed year.

2.3 Vulnerability scanning: Upon request, Disqus provides Publisher with executive summaries of external, internal, and web application vulnerability scanning. If Publisher identifies elevated risks present within provided information, Disqus will promptly remediate identified risks at Disqus’ expense. Disqus adheres to OWASP coding principles for web application code and uses static or dynamic application vulnerability scanning or manual code review, when appropriate. Disqus ensures that vulnerability scans periodically are conducted on devices and software present in its internal and external network environments and web applications to identify and remediate or have remediated any vulnerabilities within a documented timeframe. Disqus will provide Publisher with summaries of such scans and results upon request.

2.4 Penetration tests: Disqus conducts annual penetration tests on all its externally facing critical systems and applications from an industry-recognized independent third party at Disqus ‘expense. Disqus conducts similar such tests after significant changes are made to its network. Such penetration tests shall:

\(a\) be based on industry-accepted penetration testing approaches (e.g. NIST SP800-115),

\(b\) include testing from inside and outside the network,

\(c\) include testing to validate segmentation, and

\(d\) include network-layer, operating system, and application layer testing such that, at a minimum, they test against vulnerabilities identified in industry standards (e.g. OWASP Guide, SANS CWE Top 25, CERT Secure Coding).

**3. Backups**

Disqus maintains secure, usable, and traceable data/information backups to ensure that backups can be used when necessary.

**4. Internal Hardware Protection**

4.1 Disqus ensures that all computing and storage devices on its network, including but not limited to, workstations, servers, and network devices have endpoint protection in place consistent with industry standards, such as anti-malware, email, and application scanning, or antivirus.

4.2 Disqus has a tiered network architecture, which includes preventive and detective devices and where highly sensitive non-public information is in a secured and segregated network.

4.3 Disqus ensures that the network devices, servers, and workstations where Publisher information is located are hardened and continuously subject to minimum security baselines.

4.4 Disqus maintains an inventory of authorized devices that can be connected to its network environment and ensures that such inventory is reconciled periodically.

4.5 Disqus maintains an inventory of authorized software required for its network devices, servers, and workstations present in its network environment and ensures that such inventory is reconciled periodically.

4.6 Disqus maintains a change management methodology that ensures only approved changes are released and deployed in the production environment.

4.7 Disqus conducts periodic reviews of its cloud computing use based on the cloud security alliance risks and controls structure and addresses any elevated risks identified in a timely manner.

**5. Periodic Reviews and Updates**

5.1 Disqus conducts periodic reviews of its cloud computing use based on the cloud security alliance risks and controls structure and addresses any elevated risks identified in a timely manner.

5.2 Disqus promptly applies the latest firmware/security patches and updates on devices and software present in its network environment, expediting the application of critical and high-risk security patches and updates.

**6. Encryption**

6.1 Disqus ensures that all communications being initiated by Disqus or handling sensitive data are encrypted using industry-standard secure protocols.

**B. Physical Security Measures7. Data Center Security Measures**

Disqus ensures appropriate data center physical security and data center environmental controls. Disqus utilizes the following physical security measures:

7.1 a closed-circuit television monitoring system with redundant power sources that provides recognizable images and usable recordings of entrances, exits, loading docks, and other high-security areas, and which maintains all images for at least 30 days and incident images indefinitely; The media portion is kept in a secured area.

7.2 distribution logs and for all issued access devices (including keys), secured storage areas for unissued devices, and regular audits of each of the foregoing;

7.3 access control alarms that actively are monitored by appropriate personnel;

7.4 identification (relying on governmentally issued credentials) and logging of individuals accessing Disqus’ facilities (including visitors), as well as restriction of access to Publisher’s assets (including intangible assets) to individuals authorized by Publisher.

**C. Organizational Measures8. Internal Employee Procedures and Policies**

8.1 Role-based access control: Disqus maintains an up-to-date role-based access control based on data classification and job roles of employees, using the principle of least privilege and granting access only on a need-to-know basis.

8.2 Segregation of duties: Disqus maintains a segregation of duties, such that individuals performing application development are different from individuals managing production environments. Disqus employs technical and procedural controls to prevent developers and system administrators from obtaining access to production information.

8.3 Background checks: Disqus only utilizes personnel, including employees, contractors, and subcontractors, after performing background checks on them.

8.4 Security awareness training: Disqus ensures that its employees, contractors, and subcontractors receive appropriate security awareness training on a periodic basis.

8.5 Publisher’s Standard of Conduct: If non-escorted access or access to Publisher systems is required, Disqus causes its representatives to comply with Publisher’s standard of conduct.

**9. Incident Response Plan**

9.1 Disqus maintains an incident response plan that ensures that Disqus is adequately prepared to handle an incident, is able to accurately identify a Security Event as an incident, is able to contain the impact of the incident, has procedures in place to remediate the incident, has the ability to successfully recover from an incident and performs a root cause analysis of the incident.

**10. Security Event**

10.1“Security Event” shall mean an instance of Disqus learning or having reason to believe that Publisher’s confidential information has been accessed by an unauthorized person or disclosed in a manner not permitted by Disqus’ agreement with Publisher, or that an incursion in any systems, processes, hardware or software used to store, transmit or that otherwise affect Publisher’s confidential information has occurred.

10.2In a Security Event, Disqus will:

10.2.1 as soon as reasonably practicable and in no event more than seventy-two hours after becoming aware of such Security Event, provide details of the same to Publisher including (i) the date of the Security Event, (ii) details concerning the data compromised, (iii) the method of the Security Event, (iv) appropriate Disqus security personnel contacts and security personnel contacts of its representatives, (v) the name of any person or entity assisting Disqus with the investigation of the suspected or actual Security Event, (vi) a list of all parties known to have gained unauthorized access to confidential information for the limited purpose of assessing Publisher’s exposure and (vii) any other information which Publisher reasonably requests from Disqus or its representatives concerning such suspected or Security Event, including any forensics reports;

10.2.2 grant access to Publisher’s representatives or another person or entity agreed to by Publisher and Disqus (with each acting in good faith in the selection of such other person or entity) to Disqus’ systems and premises to allow such representatives or such other person or entity to perform an investigation (including the installation of any monitoring or diagnostic software) deemed necessary by Publisher to locate the source of such breach; and

10.2.3 immediately take appropriate steps to ensure that any actual data security breach does not continue.

10.3Public Authorities: Disqus will not notify law enforcement or federal or state regulatory authorities of any Security Event or other matter related to Publisher’s security requirements without prior notice to Publisher unless otherwise required by applicable law.

10.4Press Releases: Disqus will not issue any press release or other public announcement concerning a Security Event without prior approval of Publisher.

10.5Read-only logs: Disqus maintains usable read-only logs of critical systems and events on network devices, key security systems, and workstation and server operating systems, and ensures that any suspicious activity is monitored and investigated and appropriate actions are taken subsequent to its detection.

**11. Data Classification**

11.1Classes: Disqus defines classes of data/information based on applicable legal requirements and sensitivity levels of the related data/information and treats such data/information according to that classification.

**12. Collaboration with Publisher**

12.1Publisher’s security requirements: If Publisher believes that Disqus’ or any of its representatives’ security procedures in connection with the services provided to Publisher do not comply with Publisher’s security requirements, Disqus will cooperate with Publisher to ensure that security measures and procedures that comply with Publisher’s security requirements are promptly implemented.

12.2Notification: Disqus will promptly notify Publisher if Disqus learns or has reason to believe that it or any of its representatives are not in compliance with any of Publisher’s security requirements, whether or not a Security Event has occurred.

​

**ANNEX IIILIST OF SUB-PROCESSORS**

The controller has authorised the use of the following sub-processors:

1\. Name: …

Address: …

Contact person’s name, position and contact details: …

Description of processing (including a clear delimitation of responsibilities in case several sub-processors are authorised): …

​

Exhibit C CPRA ADDENDUM

**CPRA AddendumCompliance with the California Consumer Privacy Act & Consumer Privacy Rights Act Regulations**

Disqus Inc. (“**Disqus**”) and the publisher identified in the main agreement (the “**Pulisher**”) have one or more written agreements (collectively, “the **Agreement**”) pursuant to which Disqus provides services to Publisher as a “Service Provider,” a “Contractor,” or a “Third Party” (as defined below). This addendum (“**CPRA Addendum**”) shall apply to the extent that Disqus provides services to Publisher that fall under the scope of the CA Privacy Laws (as defined below).

It is the intent of the parties that Disqus acts as a Service Provider and/or a Contractor (as appropriate) for Publisher when it provides the services to Pulisher under the Agreement, provided that Disqus is a Service Provider or Contractor under the CA Privacy Laws. Disqus acts as a Third Party for Publisher when providing Cross-Contextual Behavioral Advertising or other services that CA Privacy Laws consider Third Party services.

This CPRA Addendum sets forth the requirements for contracts imposed upon the parties by the CA Privacy Laws (as defined below). This CPRA Addendum is hereby incorporated by reference into each Agreement to demonstrate the parties’ compliance with the CA Privacy Laws.

1\. Definitions.

\(a\) “CA Privacy Laws” means, collectively, the California Consumer Privacy Act of 2018 (“**CCPA**”, codified at Civil Code section 1798.100 *et seq*.), the California Privacy Rights Act (“**CPRA**”), and all applicable regulations issued by competent authorities that implement CCPA and CPRA. Words and phrases in this CPRA Addendum shall, to the greatest extent possible, have the meanings given to them in the CA Privacy Laws.

\(b\) *“Contractor” has the meaning given to it in S*ection 1798.140(j) of the California Civil Code.

\(c\) *“Service Provider” has the meaning given to it in S*ection 1798.140(ag) of the California Civil Code.

\(d\) “Third Party” has the meaning given to it in Section 1798.140(ai) of the California Civil Code.

\(e\) “Cross-Contextual Behavioral Advertising” has the meaning given to it in Section 1798.140(k) of the California Civil Code.

2\. In accordance with § 7051 of the CPRA Regulations (Contract Requirements for Service Providers and Contractors), the following terms are incorporated by reference into the Agreement to the extent that Disqus acts as a Service Provider or Contractor:

\(a\) Disqus is prohibited from selling or sharing personal information it collects pursuant to the Agreement. Disqus shall only process Publisher’s personal information for the specific business purpose(s) set forth in the Agreement and for the specific business purposes listed below:

· Providing advertising and marketing services and public relations of the Publisher’s own business or activity, goods or services.

· Providing a commenting platform to Publisher for Publisher’s website.

\(b\) Disqus is prohibited from retaining, using, or disclosing the personal information that Disqus collected pursuant to the Agreement with the Publisher for any purposes other than those specified in this CPRA Addendum, the Agreement or as otherwise permitted by the CA Privacy Laws.

\(c\) Disqus is prohibited from retaining, using, or disclosing the personal information Disqus collected pursuant to the Agreement with the Publisher for any commercial purpose other than the business purposes specified in the Agreement, including in the servicing of a different business, unless expressly permitted by the CA Privacy Laws.

\(d\) Disqus is prohibited from retaining, using, or disclosing the personal information that Disqus collected pursuant to the Agreement with the Disqus outside the direct business relationship between Disqus and Publisher unless expressly permitted by the CA Privacy Laws. For example, Disqus may not combine or update personal information Disqus collected pursuant to the Agreement with the Publisher with personal information that it received from another source or collected from its own interaction with a consumer unless expressly permitted by the CA Privacy Laws.

\(e\) Disqus shall comply with all applicable sections of the CA Privacy Laws, including providing the same level of privacy protection as required by Publisher, by cooperating with Publisher in responding to and complying with consumers’ requests made pursuant to the CA Privacy Laws, and implementing reasonable security procedures and practices appropriate to the nature of the personal information to protect the personal information from unauthorized or illegal access, destruction, use, modification, or disclosure in accordance with California Civil Code section 1798.81.5.

\(f\) Disqus grants Publisher the right to take reasonable and appropriate steps to ensure that Disqus uses the personal information in a manner consistent with the Publisher’s obligations under the CA Privacy Laws. Reasonable and appropriate steps may include ongoing manual reviews and automated scans of Disqus’s system and regular internal or third-party assessments, audits, or other technical and operational testing at least once every 12 months.

\(g\) Disqus shall notify Publisher if Disqus can no longer meet its obligations under the CA Privacy Laws.

\(h\) Disqus grants Publisher the right, upon notice, to take reasonable and appropriate steps to stop and remediate Disqus’s unauthorized use of personal information. Publisher may require Disqus to provide documentation that verifies that Disqus no longer retains or uses the personal information of consumers that have made a valid request to delete with the Publisher.

\(i\) Disqus shall enable Publisher to comply with consumer requests and Publisher shall notify Disqus of any consumer request made pursuant to the CA Privacy Laws that it must comply with and provide the information necessary for Disqus to comply with the request.

\(j\) To the extent that Disqus subcontracts with another person in providing services to Publisher, Disqus shall have a contract with the subcontractor that complies with the CA Privacy Laws.

3\. In accordance with § 7053 of the CPRA Regulations (Contract Requirements for Third Parties), the following terms are incorporated by reference into the Agreement to the extent that Disqus acts as a Third Party:

\(a\) Disqus shall only process Publisher’s personal information for the limited and specified business purpose(s) set forth in the Agreement and below.

· Cross-Context Behavioral Advertising: targeting of advertising to a consumer based on the consumer’s personal Information obtained from the consumer’s activity across businesses, distinctly-branded websites, applications, or services, other than the Publisher’s distinctly-branded website, application, or service with which the consumer intentionally interacts.

\(b\) Disqus shall comply with the CA Privacy Laws. Disqus shall provide the same level of privacy protection as required of Publisher. Disqus shall comply with a consumer’s request to opt-out of sale/sharing forwarded to Disqus by a Publisher. Disqus shall implement reasonable security procedures and practices appropriate to the nature of the personal information to protect the personal information from unauthorized or illegal access, destruction, use, modification, or disclosure in accordance with Civil Code section 1798.81.5.

\(c\) Disqus grants Publisher the right to take reasonable and appropriate steps to ensure that Disqus uses Publisher data in a manner consistent with the Publisher’s obligations under the CA Privacy Laws. Publisher may require the Disqus to attest that Disqus treats Publisher data in the same manner that Publisher is obligated to treat it under the CA Privacy Laws.

\(d\) Disqus grants Publisher the right, upon notice, to take reasonable and appropriate steps to stop and remediate unauthorized use of personal information. Publisher may require Disqus to provide documentation that verifies that Disqus no longer retains or uses the personal information of consumers who have had their request to opt-out of sale/sharing forwarded to them by Publisher.

\(e\) Disqus shall notify Publisher if Disqus can no longer meet its obligations under the CA Privacy Laws.

4\. Each party shall maintain records needed to demonstrate compliance with the applicable provisions of the CA Privacy Laws.

5\. The CPRA Addendum shall remain in force so long as the Agreement is in force and shall terminate when the Agreement is terminated.
​

### Disqus - Publisher Terms of Service Agreement for Ad Management Solutions {#disqus-publisher-terms-of-service-agreement-for-ad-managemen}

**PUBLISHER TERMS OF SERVICE AGREEMENT FOR AD MANAGEMENT SOLUTIONS**
​
**This Publisher Terms of Service Agreement (the, "Agreement") is entered into by and between Disqus, Inc. (“Disqus”) and the publisher ("Publisher") as of the first date the Disqus and Publisher execute an insertion order for ad management solutions ("Effective Date"). Therefore, in consideration of the mutual covenants of the parties and other valuable considerations, the sufficiency and receipt of which is hereby acknowledged, the parties agree as follows:1. Services.** Disqus will, among other things, manage Publisher's ad stack across the agreed upon Publisher website(s). In doing so, Disqus will provide ad tags (i.e., programming code to enable the display of a digital advertisement) for Publisher to include on the applicable website(s) each an (“Applicable Site”) and Disqus will thereby source the demand to fill impressions.

**2. Access and Use.**

2.1. *Access.* Publisher authorizes Disqus to place ad tags and advertisements on Publisher’s applicable website(s). Publisher may not modify any ad tags in such a way as to adversely impact delivery of the advertisement or an end-users ability to view an advertisement.. Publisher may add Applicable Sites not set forth in the Service Order upon execution of an additional Service Order which shall be governed by this Agreement. Publisher shall not in any way deliver, transfer, or otherwise provide access to or make available the Service to any third parties except as specifically permitted by this Agreement. Publisher is solely responsible for the activity that occurs on Publisher’s Applicable Site.

2.2. *Use.* Publisher shall use the Service in accordance with the terms of this Agreement and Disqus’s privacy policy. Publisher shall be solely responsible for maintaining its own equipment and establishing its own connection via the Internet. In no event shall Publisher, or any third party, use Disqus’s APIs to “harvest” or read in bulk the contents of the data files used in the Service, expose or otherwise make available Disqus’s APIs, including pass-through of the APIs to third parties, nor repackage the APIs to make available their functionality to third parties. Publisher shall not take any action to interfere with the Service or any other user's use of the Service, Disqus’s host or network, including, without limitation, via means of overloading, “flooding”, “mailbombing” or “crashing” the Service.

2.3*. Updates.* The parties agree that Disqus may make updates, modifications or improvements (collectively, “Updates”) to the Service from time to time in its sole discretion.

2.4. *License to Use Service*. Disqus reserves the right to revoke your license to use the Service at any time and for any reason. Disqus may also modify or discontinue the Services or any of its features at any time in our sole discretion without any responsibility or liability to you.

2.5. *Google Adsense.* Participation in Services is subject to your acceptance and continued compliance with this agreement and with the Google Adsense Advertising Policies outlined at 0. You agree that Disqus is not, and can not, under any circumstances be held responsible for the removal or banning of your site by Google Adsense or any other online advertising program. You authorize Disqus to represent you and/or act as an agent on your behalf in dealings with ad networks. Disqus may represent your website’s inventory, setup your website and domain(s) and complete other such actions in order to get approved and to get ads running from such network.

3\. *Advertising; Revenue Share.* Publisher agrees that Disqus may include advertisements and/or content provided by Disqus and/or a third party (collectively “Ads”) as part of the Service. Disqus, in its sole discretion, determines whether the Publisher’s Applicable Site(s) are eligible to receive payments for running advertisements ("Revenue Share"). Publisher agrees to comply with any specifications that may be required by Disqus from time to time to enable proper delivery, display, tracking and/or reporting of Ads. As a prerequisite to earning Revenue Share, Publisher must adhere to Disqus’ Ads.txt policy, and Publisher shall be required to submit valid payment information and relevant tax forms via Disqus’s publisher dashboard. Disqus shall have no obligation to pay Publisher in the event Disqus has not received payment from its advertisers. Disqus shall not be liable for any payment based on: (a) any amounts which result from invalid queries, invalid Referral events, or invalid clicks or impressions on Ads generated by any person, bot, automated program or similar device, as reasonably determined by Disqus, including without limitation through any clicks or impressions (i) originating from Publisher’s IP addresses or computers under Publisher’s control, (ii) solicited by payment of money, false representation, or request for end users to click on Ads, or (iii) solicited by payment of money, false representation, or any illegal or otherwise invalid request for end users to complete referral events; (b) ads delivered to end users whose browsers have JavaScript disabled; (c) ads benefiting charitable organizations and other placeholder or transparent ads that Disqus may deliver; or (d) clicks co-mingled with a significant number of invalid clicks described in (a) above, or as a result of any breach of this Agreement by Publisher for any applicable pay period. Disqus reserves the right to withhold payment or charge back Payee’s account due to any of the foregoing or any breach of this agreement by Publisher, pending Disqus’s reasonable investigation of any of the foregoing or any breach of this Agreement by Publisher, or in the event that an advertiser or ad network whose ads are displayed in connection with Publisher’s site defaults on payment for such ads to Disqus. Disqus shall pay Publisher the Revenue Share due to Publisher ninety (90) days from the end of each calendar month that Ads are running on the Applicable Site(s). Payment will be distributed through Tipalti. Disqus shall not distribute Revenue Share to Publisher if the amount due to Publisher is less than US\$100. Publisher shall be required to claim Revenue Share from Disqus within three (3) months of the date Revenue Share was distributed to Publisher. In the event Publisher does not claim Revenue Share within such time period, Disqus shall have the right to reclaim such Revenue Share. Disqus reserves the right, in its sole discretion, not to run Ads on the Applicable Site(s) for any reason, or no reason, including, but not limited to, quality of the content or content requirements from Disqus’s advertisers. Publisher shall promptly notify Disqus if Publisher has any legal obligations to show specific content or advertisements on Publisher’s site and Publisher will indemnify and hold Disqus, and its subsidiaries, affiliates, officers, agents and employees, harmless from any claim or demand, including reasonable attorneys’ fees, arising out of the removal or failure to display such content or advertisements.

**4. Data Privacy.**

4.1. *License to Use Disqus Personal Data.* “Disqus Personal Data” means all personal data that is collected, transmitted, displayed, uploaded, or exchanged by or through the Service. Disqus hereby grants Publisher a limited, non-exclusive, and revocable license to use Disqus Personal Data for comment moderation, ad placement and tracking, and analytics purposes only (the “Permitted Purpose”).

4.2. *Data Processing.* For the purposes of this clause, the terms "controller", "data subjects", "personal data", "processor", "processing", and “supervisory authority” shall have the meaning given to them by the European Regulation 2016/679 (“GDPR”). Disqus and Publisher shall each be controllers of Disqus Personal Data, and both parties shall process Disqus Personal Data only in accordance with the Permitted Purpose. If Publisher is required to process Disqus Personal Data for any other purpose by a law to which Publisher is subject, Publisher shall inform Disqus of this requirement before the processing, unless prohibited by applicable law. Publisher shall ensure that (i) its personnel and subcontractors who have access to Disqus Personal Data have committed themselves to confidentiality and are aware of and comply with Publisher's duties and their personal duties and obligations under this Agreement (ii) implement appropriate technical and organizational security measures to ensure a level of security appropriate to the risks that are presented by the processing of Disqus Personal Data. In case of a personal data breach which affects Disqus Personal Data, Publisher will notify Disqus without undue delay after becoming aware of it, (iii) taking into account the nature of the processing, assist Disqus by appropriate technical and organizational measures insofar as it is possible to fulfill Disqus's obligations to respond to requests from data subjects exercising their rights; (iv) taking into account the nature of the processing and the information available to Publisher, assist Disqus, at Disqus's cost, to ensure compliance with the obligations under applicable privacy law with respect to security, breach notifications, impact assessments and consultations with supervisory authorities or regulators; (v) upon termination of this Agreement or upon Disqus's request, destroy or return all Disqus Personal Data to Disqus (unless a law requires storage of Disqus Personal Data), and (vi) make available to Disqus all information reasonably necessary to demonstrate compliance with the obligations laid down in this section and allow for and contribute to audits, including inspections, conducted by Disqus or an auditor mandated by Disqus. Disqus acknowledges and agrees that Publisher may retain its affiliates and other third parties as sub-processors (all together "Sub-Processors") in connection with the provision of the Services having imposed on such Sub-Processors the same data protection obligations as are imposed on Publisher under this Agreement. Publisher will be liable to Disqus for the performance of the Sub-Processors' obligations. Publisher will inform Disqus in advance of any changes concerning the addition or replacement of Sub-Processors.

4.3. *Cookies.* Disqus shall be permitted to place or recognize a cookie on the visitors to the Applicable Sites for the purpose of collecting Disqus Personal Data relating to the visitor’s activity and interaction with the Service, or content on the Applicable Sites, and information about the visitor’s device ID, browser type, environmental or location information, or other similar information, as set forth in Disqus’s privacy policy (“Disqus Cookie Data”). To the extent that Cookie Tracking is turned on, and subject to its compliance with applicable Privacy Laws (as defined below), Disqus will also cause third-party cookies to be served. Publishers may choose to turn off Cookie Tracking at any time, however, Publisher shall not be eligible to for Ad Revenue unless Cookie Tracking is turned on. Publisher further agrees that, to the extent Cookie Tracking is turned on, and to the extent required by Privacy Laws, the Applicable Sites contain a mechanism to obtain the user’s consent for the collection of Disqus Cookie Data for GDPR or other applicable legal purposes and a “Do Not Sell” or “Privacy Choices” button or link available on, at a minimum, the home page of each Publisher-owned site that utilizes Disqus, and each web page where the Disqus comment section or Disqus logo is displayed. Publisher will be solely responsible for obtaining user consent for the placement of cookies at the Publisher’s website to the extent required by applicable Privacy Laws.

4.4. *Compliance with Privacy Laws.* Both Disqus and Publisher shall comply fully with all applicable laws, rules, regulations, and government orders relating to data protection and data privacy, including, but not limited to, the GDPR, the CCPA, and other U.S. state or federal privacy laws, regulations, court precedent, and regulatory agency orders, (collectively “Privacy Laws”), and will only collect, use and disclose Disqus Personal Data collected through the Service and the Applicable Site(s) as set forth in this Agreement and in compliance with applicable Privacy Laws. Publisher will ensure that each of its Applicable Sites includes a privacy policy that complies with all Privacy Laws and specifically (i) discloses that the site shares personal data with Disqus, (ii) discloses the usage of third-party technology; and to the extent Cookie Tracking is turned on, the data collection and usage by Disqus; and (iii) contains a conspicuous live hyperlink to give users the ability to opt out of interest-based or cross context behavioral advertising through the Service. Publisher and Disqus agree to comply with the obligations set out in the Standard Contractual Clauses, which are incorporated herein by reference. “Standard Contractual Clauses” means the applicable module(s) of the European Commission’s standard contractual clauses for the transfer of personal data to third countries pursuant to Regulation (EU) 2016/679 of the European Parliament and of the Council, as set out in the Annex to Commission Implementing Decision (EU) 2021/914 (“Standard Contractual Clauses”). The Controller-to-Controller Standard Contractual Clauses shall apply in all cases where Disqus Personal Data that relates to residents of a Restricted Country (as defined below) is processed by Disqus. In particular, and without limiting the above obligations: (i) Publisher and Disqus agree that their respective obligations under the Standard Contractual Clauses shall be governed by the law(s) of the Member State(s) (or Switzerland or the United Kingdom) in which users are established; and (ii) the details of the appendices applicable to the Standard Contractual Clauses are set out in Exhibit B to the data processing agreement, which is incorporated herein by reference. “Restricted Country” means a member state of the European Economic Area, Argentina, Brazil, Canada, Chile, China, Costa Rica, Ghana, Hong Kong, Israel, Malaysia, Mexico, Morocco, Russia, Saudi Arabia, Singapore, Switzerland, Tunisia, Turkey, the United Kingdom, or Uruguay.

**5. Intellectual Property.** Notwithstanding anything to the contrary in this agreement, all intellectual property rights (a) owned or licensed by a party before the date of this agreement and (b) created, developed or licensed by that party after the date of this Agreement independently of this Agreement shall continue to vest in that party. Publisher acknowledges that all intellectual property rights in the Service (including any improvements, enhancements and modifications thereto), are Disqus’s Confidential Information (as defined below) and any other of Disqus’ software, data, or information provided or made available to Publisher under this Agreement (together the “Disqus’s Intellectual Property”) shall belong to Disqus and Publisher shall have no rights in or to Disqus’s Intellectual Property other than the right to use it in accordance with the terms of this Agreement. Unless otherwise agreed to in writing, Publisher shall not remove or obscure any copyright, trademark or patent notice that appears on the Service.

**6. Confidential Information**

6.1. *Confidential Information.* In connection with this Agreement, each party may disclose, or may learn of or have access to, certain confidential proprietary information owned by the other party (“Confidential Information”). Confidential Information means any non-public data or information, oral or written, that relates to a party, or any of its business activities, technology, developments, inventions, processes, trade secrets, know how, source code, plans, financial information, Publisher and supplier lists, forecasts, and projections. Notwithstanding the foregoing, Confidential Information is deemed not to include information that: (i) is publicly available or in the public domain at the time disclosed; (ii) is or becomes publicly available or enters the public domain through no fault of the receiving party; (iii) is rightfully communicated to the receiving party by persons not bound by confidentiality obligations with respect thereto; (iv) is already in the receiving party's possession free of any confidentiality obligations with respect thereto; (v) can be documented as independently developed by a party without use of any Confidential Information of the other party; or (vi) is approved for release or disclosure by the disclosing party without restriction. Each party shall use reasonable measures to maintain the Confidential Information of the other party in confidence and shall not disclose, publish or copy any part of such Confidential Information, to any third party. Each party shall only use the Confidential Information of the other party for the purpose of this Agreement and shall limit disclosures to any employees on a strict need-to-know basis. Notwithstanding the foregoing, a party may disclose Confidential Information of the other party pursuant to the order or requirement of a court, administrative agency, or other governmental body, provided that such party gives reasonable prior notice (if permissible) to the other party to contest such order or requirement. Upon request, each party shall return to the other party, or certify the destruction of, all Confidential Information of the other party.

**7. Representations and Warranties.**

7.1. *Mutual Representations.* Each party represents and warrants to the other party that: (i) it has the full corporate right, power and authority to enter into this Agreement and to perform the acts required of it hereunder; (ii) the execution of this Agreement and the performance of its obligations hereunder, do not and will not violate any agreement to which it is a party or by which it is bound; and (iii) when executed and delivered, this Agreement will constitute the legal, valid and binding obligation of such party, enforceable against it in accordance with its terms.

7.2. *Publisher Representations.* Publisher represents and warrants to Disqus that: (i) it owns, operates, or controls all Applicable Sites; (ii) the Applicable Sites do not contain materials that infringe or violate any third party proprietary rights including, but not limited to, third party intellectual property rights, or materials that violate any applicable laws, rules, or regulations and Privacy Laws; and (iii) the Applicable Sites do not contain any harmful or disabling software code, including without limitation any virus, time-bomb or trojan horse.

7.3. *Disclaimer of Warranties.* Except for the express warranties provided for herein, the service, and any support services are provided to Publisher “as is” and Disqus expressly disclaims all warranties, express, implied or statutory, including but not limited to the implied warranties of merchantability, fitness for a particular purpose, and noninfringement, and any warranties arising out of course of dealing, usage, or trade. Disqus does not warrant that the service or any updates will meet Publisher's specific requirements or that the operation of the service or updates will be completely error-free or uninterrupted. Disqus shall not be liable to Publisher for any inoperability of the service or for any loss of information or other injury, damage or disruption of any kind. Disqus makes no guarantee regarding the level of pageviews or ad impressions or clicks or the amount of payment made to you under this agreement.

**8. Limitation of Liability.** IN NO EVENT WILL EITHER PARTY BE LIABLE TO THE OTHER FOR ANY SPECIAL, INDIRECT, INCIDENTAL OR CONSEQUENTIAL DAMAGES (INCLUDING WITHOUT LIMITATION LOSS OF USE, DATA, BUSINESS OR PROFITS OR COSTS OF COVER) ARISING OUT OF OR IN CONNECTION WITH THIS AGREEMENT OR THE USE OR PERFORMANCE OF THE SERVICE AND/OR UPDATE(S), WHETHER SUCH LIABILITY ARISES FROM ANY CLAIM BASED UPON CONTRACT, WARRANTY, TORT (INCLUDING NEGLIGENCE), PRODUCT LIABILITY OR OTHERWISE, AND WHETHER OR NOT DISQUS HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH LOSS OR DAMAGE. IN NO EVENT SHALL DISQUS’S CUMULATIVE LIABILITY TO THE OTHER EXCEED THE FEES PAID TO DISQUS BY PUBLISHER DURING TWELVE (12) MONTHS PRECEDING THE INCIDENT GIVING RISE TO SUCH LIABILITY.

**9. Indemnification.**

9.1. *Disqus.* Disqus shall indemnify, defend and hold harmless Publisher and its affiliates, and their respective shareholders, officers, directors, employees, agents, successors and assigns from and against any and all third party claims for losses, liabilities, costs, expenses (including amounts paid in settlement and reasonable attorneys’ fees and expenses), penalties, judgments and damages (“Losses”) resulting from any claim by a third party that the Services or infringe or violate the intellectual property rights of any third party, provided, in each case, that (i) Disqus is promptly notified in writing of the claim; (ii) Disqus has sole control of the defense and any negotiations for the settlement of such claim; and (iii) the indemnified party provides to Disqus, at Disqus’s expense, with all reasonable assistance, information, and authority necessary to perform the above. Should the Services in Disqus's opinion, be likely to become, the subject of a claim of infringement, Disqus may, at its option and expense, either procure for Publisher the right to continue using the Services or replace or modify the Services or Work Product in order to make them non-infringing.

9.2. *Publisher.* Publisher agrees to indemnify, defend and hold harmless Disqus, its affiliates and their respective officers, directors, and employees from and against any and all Losses to the extent that such is based upon any third party claim in connection with (i) Publisher’s breach of any of its representations or warranties made hereunder; (ii) Publisher’s violation of any applicable laws, rules or regulations, including, but not limited to, any data protection and data privacy laws and regulations and industry association guidelines; or (iii) Publisher’s violation of any third party intellectual property right.

**10. Term and Termination**

10.1. *Term.* This Agreement shall commence on the Effective Date and shall continue for an initial term of twelve (12) months following the Effective Date (the “Initial Term”). After the expiration of the Initial Term, this Agreement shall automatically renew for additional twelve (12) month periods unless either party gives not less than ninety (90) days’ prior written notice of its intention not to renew (the initial term and any Renewal Term collectively referred to as the “Term”).

10.2. *Termination.* This Agreement shall terminate: (i) by a party thirty (30) business days after the other party’s receipt of written notice that such party is in material breach of any of the terms or conditions set forth in this Agreement, unless such party cures such breach within said thirty (30) business days period or (ii) upon written notice if the other party becomes insolvent, makes a general assignment for the benefit of creditors, files a voluntary petition of bankruptcy, suffers or permits the appointment of a receiver for its business or assets, becomes subject to any proceedings under any bankruptcy or insolvency law, whether domestic or foreign, or has wound up or liquidated its business voluntarily or otherwise, and same has not been discharged or terminated within ninety (90) days. Notwithstanding the foregoing, Disqus may immediately and without prior notice terminate or suspend Publisher’s access to the Service in the event Disqus reasonably believes that continued Publisher access or storage may harm the Service, expose Disqus to liability or is necessary to comply with applicable law.

10.3. *Upon Termination.* Upon the effective date of expiration or termination of this Agreement for any reason, whether by Publisher or Disqus, Publisher’s right to use the Service shall immediately cease. It is Publisher’s sole responsibility to download Disqus Personal Data; Disqus has no obligation to make any data available to the Publisher following the date of termination. Publisher can request a copy of Disqus Personal Data from Disqus only for additional cost determined by Disqus. Disqus has the right to deny such request at its sole discretion. Promptly upon expiration or termination of this Agreement for any reason, Publisher shall pay any unpaid and outstanding Fees due to Disqus that have accrued as of the date of expiration or termination and Publisher shall return to Disqus, or certify the destruction of, all copies of Disqus’s Confidential Information.

**11. General Provisions**

11.1. *Severability and Waiver.* If any provision of this Agreement is held to be void, invalid or inoperative, the remaining provisions of this Agreement shall continue in effect and the invalid portion of any provision shall be deemed modified to the least degree necessary to remedy such invalidity while retaining the original intent of the parties. The failure of either party to partially or fully exercise any rights or the waiver of either party of any breach shall not prevent a subsequent exercise of such right or be deemed a waiver of any subsequent breach of the same or any other term of this Agreement.

11.2. *Independent Contractors.* Each party to this Agreement is an independent contractor in relation to the other party with respect to all matters arising under this Agreement. Nothing herein shall be deemed to establish a partnership, joint venture, association or employment relationship between the parties. Publisher may not assign any of its rights or obligations under this Agreement to any other entity without the prior written consent of Disqus, which shall not be unreasonably withheld.
​

11.3. *Assignment.* Neither party may, or shall have the power to, assign this Agreement without the prior written consent of the other; provided, however, that either party may assign its rights and obligations under this Agreement without the approval of the other party to any subsidiary or Affiliate or successor in connection with a merger, consolidation, sale of all of the equity interests of the party, or a sale of all or substantially all of the assets of the party to which this Agreement relates; provided, that in no event shall such assignment relieve such party of its obligations under this Agreement. Subject to the foregoing, this Agreement shall be binding on the parties hereto and their respective successors and assigns.

11.4. *Entire Agreement.* This Agreement, including any exhibits and schedules attached hereto, constitutes the entire agreement between the parties on this subject matter and supersedes all prior negotiations, understandings and agreements between the parties concerning this subject matter. Neither Party will be bound by, and each party specifically objects to, any term, condition, or other provision which is different from or in addition to the provisions of this Agreement (whether or not it would materially alter this agreement). No amendment or modification of this Agreement shall be made except by a writing signed by both parties.

11.5. *Survival*. The provisions of this Agreement, which by their nature are intended to survive after termination or expiration of this Agreement shall so survive the expiration or termination of this Agreement regardless of the reason or reasons therefore.

11.6. *Freedom of Action*. Either party is free to enter into similar agreements with others and may design, develop, manufacture, acquire or market competitive products or services. Either party may assign and re-assign its employees in any way it may choose and neither party is restricted in any way from hiring or soliciting employees of the other.

11.7. *Counterparts Acceptable*. This Agreement may be executed in any number of counterparts, each of which shall be an original and all of which together shall constitute one and the same document.

11.8. *Publicity.* Disqus shall be entitled, without prior consultation with or approval of the Publisher, to make press releases or other public disclosures with respect to this transaction. Publisher grants Disqus a non-exclusive license during the Term to use its name and trademarks in marketing materials, website or customer lists; provided, that Publisher has the right to notify Disqus in writing if it does not agree to any of the foregoing uses of its name and trademarks.

11.9. *Force Majeure*. Except for payment obligations, neither party shall be in breach of this Agreement or responsible for damages caused by delay or failure to perform, in full or in part, its obligations hereunder, provided that there is due diligence in attempted performance under the circumstances and that such delay or failure is due to fire, earthquake, unusually severe weather, strikes, government sanctioned embargo, flood, act of God, act of war or terrorism, act of any public authority or sovereign government, civil disorder, delay or destruction caused by public carrier, or any other circumstance substantially beyond the control of the party to be charged.

11.10. *Governing Law; Jurisdiction.* The validity, interpretation, performance and enforcement of this Agreement shall be governed by the laws of the State of California and each party irrevocably submits to exclusive jurisdiction and venue in the courts located in Santa Clara County, California. The United Nations Convention on contracts for the International Sales of Goods shall not apply. The remedies under this Agreement shall be cumulative and not alternative and the election of one remedy for a breach shall not preclude pursuit of other remedies unless expressly provided otherwise in this Agreement. Disqus shall be entitled to collect its reasonable attorney’s fees, costs and expenses in any action brought to seek amounts past due or to otherwise enforce rights hereunder.

11.11. *Notice*. All notices and other communications hereunder shall be in writing and shall be deemed to have been duly given when delivered in person (including by overnight courier) or three days after being mailed by registered or certified mail (postage prepaid, return receipt requested) or sent by email, and on the date the notice is sent when sent by verified facsimile or email, in each case to the respective Parties at the address first set forth hereto.

### Disqus Informativa Sulla Riservatezza {#disqus-informativa-sulla-riservatezza}

**Aggiornato** il 10° luglio 2026

Questa Informativa sulla Privacy ti spiega come Disqus raccoglie, utilizza, vendi, divulga e protegge i dati relativi a te (l'"Utente") in relazione al nostro Servizio (come definito di seguito), nonché le tue scelte riguardo alla raccolta e all'uso di questi dati.

#### 1. INTRODUZIONE

**Panoramica**

Disqus offre una piattaforma online di condivisione di commenti e opinioni pubbliche dove gli utenti accedono e creano profili per partecipare a conversazioni con i colleghi e godersi un'esperienza interattiva nelle sezioni commenti, nei sondaggi e in altre funzionalità interattive offerte su questo sito, oltre che integrate in siti di terze parti. L'uso della nostra piattaforma e software, e l'interazione con i nostri cookie o tecnologie di tracciamento simili (collettivamente il "Servizio"), sia su questo sito che su un sito di terze parti, sono soggetti ai termini di questa Informativa sulla Privacy. Il Servizio è una piattaforma pubblica e Disqus o altri possono cercare, vedere, utilizzare o ripubblicare qualsiasi tuo Contenuto Utente (come definito nei nostri Termini di Utilizzo) che pubblichi tramite il Servizio. Disqus è anche una società di marketing e dati, e utilizza e condivide dati personali raccolti da siti di terze parti dove il nostro Servizio è abilitato per scopi di marketing, inclusa la pubblicità comportamentale cross-contest. Per ulteriori informazioni sulle nostre attività di marketing, consulta la Sezione 4: Partner pubblicitari e pubblicitari mirati qui sotto.
​

**Applicabilità a siti web e servizi di terze parti**

Disqus offre un servizio di coinvolgimento online che altri siti web utilizzano per facilitare discussioni e interattività tra i loro utenti. Questa Informativa sulla Privacy si applica ai dati raccolti da Disqus sugli Utenti del Servizio e tramite cookie sui siti web abilitati al Servizio, e non si applica alle pratiche indipendenti di raccolta dati di qualsiasi sito web che utilizza il Servizio o altri siti web collegati dal Servizio. Per informazioni su come i siti web di terze parti raccolgono e utilizzano le tue informazioni personali, consulta le politiche sulla privacy di quei siti.
​

**I tuoi diritti alla privacy**

Hai diritti sulle tue informazioni personali. Questi diritti sono descritti in modo più dettagliato nella Sezione 9: I tuoi diritti, qui sotto. Puoi esercitare i tuoi diritti sulla privacy dei dati presso Ecco.
​

**Privacy dei bambini**

Il servizio non è destinato all'uso da parte di bambini sotto i 18 anni. Non raccogliamo né vendiamo consapevolmente informazioni personali di minori di 18 anni né permettiamo consapevolmente a tali persone di registrarsi per un account sul servizio. Nel caso venga a sapere di aver raccolto informazioni personali da un bambino sotto i 18 anni, le cancelleremo. Se ritieni che potremmo aver raccolto informazioni personali da un minore di 18 anni, ti preghiamo di contattarci o di presentare una richiesta di diritti sulla privacy dei dati all'indirizzo Ecco.
​

**Il nostro uso dell'intelligenza artificiale**

Disqus utilizza intelligenza artificiale e apprendimento automatico in due modi. Innanzitutto, l'IA viene utilizzata per aiutare a moderare i contenuti sulla piattaforma — rilevando spam e contenuti che violano le nostre linee guida della comunità, così che il Servizio possa essere mantenuto sicuro e funzionante per gli utenti. In secondo luogo, l'IA viene utilizzata per offrirti pubblicità più pertinente in base ai tuoi interessi e alle tue attività online. In entrambi i casi, Disqus non utilizza l'IA per prendere decisioni che abbiano effetti legali o di altrettanta importanza sugli individui. Disqus non utilizza né vende dati personali dei consumatori per addestrare grandi modelli linguistici. I nostri sistemi di IA operano in conformità con i nostri impegni sulla privacy come descritto in questa politica.

#### 2. I DATI CHE RACCOGLIAMO SU DI TE

Dati personali, o informazioni personali, indicano qualsiasi informazione su una persona che possa ragionevolmente essere collegata, direttamente o indirettamente, a una persona o a un nucleo familiare specifico.

Raccogliamo, e abbiamo raccolto negli ultimi dodici (12) mesi, i seguenti tipi di dati personali sugli Utenti:

1\. **Identificatori ("Dati di Identità")** come nome, cognome, nome utente o identificatore simile, indirizzo IP (protocollo internet), ID cookie univoco, ID dispositivo, data di nascita, indirizzo email, numero di telefono e indirizzo postale;

2\. **Le categorie di Informazioni Personali elencate nello Statuto dei Registri dei Clienti della California (Codice Civile della California § 1798.80(e)),** che includono qualsiasi informazione che identifichi, si riferisca, descriva, possa essere associata a un individuo particolare ma non sia pubblicamente disponibile al pubblico generale tramite registri governativi federali, statali o locali;

3\. **Caratteristiche di classificazione protette** secondo la legge californiana o federale, come razza, genere o età;

4\. **Informazioni sull'attività di Internet o di altre reti elettroniche,** come la cronologia di navigazione e commenti, feedback e risposte ai sondaggi, i tuoi dati di accesso, tipo e versione del browser, impostazione e posizione del fuso orario, tipi e versioni dei plug-in del browser, sistema operativo e piattaforma e altre tecnologie sui dispositivi che usi per accedere al Servizio;

5\. **Informazioni professionali o legate all'occupazione,** nella misura in cui le includi nel tuo profilo o nei commenti, o che possano essere dedotte dalle pagine che visualizzi;

6\. **Informazioni educative,** nella misura in cui le includi nel tuo profilo o nei commenti, o che possono essere dedotte dalle pagine che visualizzi;

7\. **Informazioni personali sensibili.** Non raccogliamo intenzionalmente dati personali sulla tua razza o etnia, credenze religiose o filosofiche, vita sessuale, orientamento sessuale, opinioni politiche, appartenenza a sindacati, informazioni sulla tua salute o dati genetici o biometrici, o informazioni su condanne e reati penali. Tuttavia, se fai commenti utilizzando il Servizio che includono tali dati su di te, questi saranno disponibili pubblicamente e potranno essere trattati da Disqus o da altri. Inoltre, raccogliamo e condividiamo informazioni raccolte tramite cookie o tecnologie di tracciamento simili sulle pagine web che hai consultato, il che potrebbe permettere alle terze parti con cui condividiamo le tue informazioni di trarre deduzioni su di te che potrebbero costituire informazioni personali sensibili.
​

Possiamo anche combinare, de-identificare o aggregare qualsiasi delle informazioni raccolte tramite il nostro Servizio per uno qualsiasi degli scopi descritti di seguito.

#### 3. COME VENGONO RACCOLTI I TUOI DATI PERSONALI?

Utilizziamo diversi metodi per raccogliere dati da e su di te, inclusi tra:

**Interazioni dirette**

Questo include i dati personali che fornisci quando crei un account o lasci un commento.
​

**Tecnologie o interazioni automatizzate**

Mentre interagisci con il nostro Servizio, possiamo raccogliere automaticamente Dati Tecnici sulle tue apparecchiature, sulle azioni di navigazione e sui pattern. In particolare, nei seguenti modi:
​

**Biscotti**

Un cookie è un piccolo file digitale posizionato sull'hard disk del tuo computer. Puoi rifiutarti di accettare i cookie del browser attivando le impostazioni del browser o, in alcuni casi, interagendo con banner pop-up per cookie. Utilizziamo cookie inseriti su siti web di terze parti su cui il Servizio è abilitato per raccogliere informazioni su come interagisci con quei siti anche se non lasci commenti, non rispondi ai sondaggi o non interagisci direttamente con il Servizio su quei siti, oltre a informazioni sugli altri siti che visiti. Disqus utilizza cookie e permette ai partner di impostare anche i cookie tramite il Servizio per facilitare la pubblicità comportamentale cross-contest. In pratica, questo significa che utilizziamo i cookie per aiutare a determinare quali annunci vedi online, registrando le tue visite su molti siti web di inserzionisti o brand, e poi mostrandoti annunci per prodotti e servizi simili.
​

Disqus utilizza cookie di 'autenticazione', ad esempio sessionid, disqusauth e disqusauths, per mantenerti connesso dal browser web e personalizzare la tua esperienza Disqus.
​

Disqus utilizza cookie 'unici', ad esempio disqus_unique e \_jid, per associare le attività web al caricamento di una pagina e a un browser web, e per comprendere i tuoi interessi e l'uso del prodotto.
​

Quando Disqus carica annunci, utilizziamo tecnologie di pubblicità di Google che possono impostare cookie per scopi di marketing personalizzato, associare annunci ad attività successive e limitare la frequenza con cui vengono mostrati specifici annunci.
​

**Informazioni sul file di registro**

I log dei server raccolgono dati tecnici come il tuo indirizzo IP, il tipo di browser e informazioni sul numero di clic e su come interagisci con i link sul Servizio, siti partner, nomi di dominio, landing page, pagine visualizzate e altre informazioni simili.
​

**Pixel e tracker simili**

Quando utilizzi il Servizio, utilizziamo pixel, gif chiari (noti anche come web beacon) che vengono utilizzati per raccogliere dati tecnici e informazioni come i modelli di utilizzo online. Utilizziamo anche gif chiare nelle email basate su HTML inviate ai nostri utenti per tracciare quali email vengono aperte e quali link o pubblicità vengono cliccati dai destinatari. Utilizziamo anche pixel e gif chiari su siti web di terze parti su cui il Servizio è abilitato per raccogliere informazioni su come interagisci con quei siti, anche se non lasci commenti, non rispondi ai sondaggi o interagisci direttamente con il Servizio su quei siti.
​

**Terze parti o fonti pubbliche**

Otteniamo o riceviamo dati personali su di te da fornitori di analisi come Google; Partner pubblicitari [Ecco](#cookies-and-data-recipients); e broker di dati terzi che vendono dati personali. Otteniamo o riceviamo dati personali da connessioni o accessi di terze parti tramite piattaforme social come Facebook Connect, Google o Twitter/X quando "segui", "metti mi piace" o colleghi il tuo account al Servizio. Si prega di notare che alcuni di questi fornitori, in particolare quelli di analisi come Google, possono elaborare dati provenienti da residenti dell'Area Economica Europea (EEE) al di fuori del SEE.
​

**Gestione dei cookie nel tuo browser**

Potresti essere in grado di modificare le impostazioni del browser per gestire le preferenze sui cookie, ad esempio impostando il browser per notificarti quando ricevi un cookie e darti la possibilità di decidere se accettarlo o meno. Se rifiuti i cookie, puoi comunque utilizzare il nostro sito, ma la funzionalità di alcune aree potrebbe essere limitata.
​

Di seguito sono riportati link a informazioni su come gestire le tue preferenze sui cookie nei browser comuni:

· Cookie di Google Chrome: Cookie su Google Chrome

· Cookie Mozilla Firefox: Mozilla Firefox Cookie

· Cookie su Internet Explorer: Cookie di Internet Explorer

· Biscotti Safari: Safari Cookies

· Google Analytics: Google Analytics

#### 4. PARTNER PUBBLICITARI E PUBBLICITARI MIRATI

**Pubblicità mirata**

La pubblicità è il modo principale in cui Disqus guadagna. I ricavi pubblicitari permettono a Disqus di gestire, supportare e migliorare il Servizio. Disqus utilizza e condivide con partner pubblicitari terzi e affiliati, cookie ID, ID dei dispositivi (inclusi dispositivi mobili), indirizzi email hashati, indirizzi IP, provider Internet (ISP) e informazioni sul browser, dati demografici o di interesse, contenuti visualizzati e azioni intraprese sul Servizio, sui siti partner o su altri siti di terze parti. Questo include informazioni sui siti web che hai visitato e sulle pubblicità con cui hai interagito, al fine di offrirti pubblicità più pertinente e riguardata alle tue preferenze e interessi. Questo può derivare dalla tua interazione con il Servizio, i siti partner o altri siti di terze parti. Per un elenco dei partner pubblicitari di terze parti con cui Disqus sta attualmente lavorando, vedi [Ecco](#cookies-and-data-recipients).
​

**Email Marketing**

Disqus può anche inviarti newsletter via email e messaggi di email marketing se ci hai dato il permesso o hai acconsentito a ricevere tali email, come richiesto dalla giurisdizione in cui risiede. I messaggi di email marketing possono essere adattati ai tuoi interessi in base alle informazioni sopra descritte in questa sezione. Per informazioni su come rinunciare e per esercitare i propri diritti alla privacy, si prega di consultare [Ecco](#updating-your-account-settings).
​

**Partner pubblicitari e divulgazioni di terze parti**

Collaboriamo e condividiamo dati con terze parti che raccolgono informazioni su vari canali, sia offline che online, con l'obiettivo di offrire pubblicità più pertinente a te o alla tua azienda. I nostri partner utilizzano queste informazioni per riconoscerti attraverso diversi canali e piattaforme, nel tempo (inclusi, ma non limitati a, computer, dispositivi mobili, TV indirizzabili o altri media), per scopi di marketing, analisi, attribuzione e reportistica. Sebbene Disqus non tragga inferenze basate sui tuoi dati, i nostri partner pubblicitari possono trarre deduzioni dai tuoi dati per comprendere le tue preferenze, caratteristiche, tendenze psicologiche, predisposizioni, comportamenti, atteggiamenti, intelligenza, capacità e attitudine.
​

Negli ultimi dodici (12) mesi abbiamo condiviso o venduto i seguenti dati:

· Identificatori;

· Le categorie di informazioni personali elencate nello Statuto dei Registri dei Clienti della California (Codice Civile della California § 1798.80(e);

· Caratteristiche di classificazione protetta;

· informazioni sull'attività di Internet o altre reti elettroniche;

· Informazioni professionali o legate all'occupazione;

· Informazioni educative; e,

· Informazioni personali sensibili.

#### 5. COME UTILIZZIAMO I DATI PERSONALI E LA NOSTRA BASE LEGALE PER L'UTILIZZO

Nel SEE, nel Regno Unito (UK) e in Brasile, di solito ci affidiamo al tuo consenso per utilizzare i dati personali. In alcuni casi, utilizziamo i dati quando necessario per i nostri legittimi interessi (o quelli di terzi), ma solo quando i tuoi diritti e interessi non sono negativamente influenzati. Infine, possiamo anche utilizzarla quando dobbiamo rispettare un obbligo legale o normativo, o per proteggere la salute, la sicurezza o i diritti legali di una persona.
​

Non utilizziamo informazioni personali per sostenere decisioni esclusivamente automatizzate che producono effetti legali o simili significativi su di te, note come "profiling" secondo alcune leggi.
​

Di seguito abbiamo illustrato una descrizione dei modi in cui utilizziamo i dati personali e su quali basi legali ci affidiamo per farlo. Abbiamo anche identificato quali sono i nostri interessi legittimi quando appropriato.
​

Si noti che potremmo elaborare i tuoi dati personali per più di un motivo legale, a seconda dello scopo specifico per cui utilizziamo i tuoi dati. Contattaci se hai ulteriori domande.

-   **Scopo o attività**
    -   Tipo di dati personali
    -   Fonte dei dati personali
    -   Base dell'elaborazione
-   **Per registrarti come nuovo utente**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Attività su Internet o Rete Elettronica
    -   · Interazioni dirette · Tecnologie Automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Esecuzione di un contratto con te · Con il tuo consenso
-   **Per gestire il nostro rapporto con te, che può includere notificarti di modifiche ai nostri termini o all'informativa sulla privacy, chiederti di lasciare una recensione o partecipare a un sondaggio**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Esecuzione di un contratto con te · Necessari per rispettare un obbligo legale · Necessari per i nostri legittimi interessi (per mantenere aggiornati i nostri registri e studiare come i clienti utilizzano il nostro servizio) · Con il tuo consenso
-   **Gestire e proteggere la nostra attività e il Servizio (inclusi la risoluzione dei problemi, l'analisi dei dati, i test, il supporto, la reportistica e l'hosting dei dati)**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Necessari per rispettare un obbligo legale · Necessari per i nostri legittimi interessi (ad esempio per mantenere aggiornati i nostri registri e studiare come i clienti utilizzano il nostro servizio)
-   **Per consegnarti contenuti e pubblicità rilevanti e misurare o comprendere l'efficacia della pubblicità che ti offriamo**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica · Informazioni professionali o legate all'occupazione · Informazioni educative · \[Informazioni sensibili\]
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Consenso. Quando hai fornito un consenso esplicito all'uso dei dati personali per fornirti contenuti rilevanti e personalizzati · Necessaria per i nostri legittimi interessi
-   **Utilizzare l'analisi dei dati per migliorare il nostro servizio**
    -   · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Necessari per i nostri legittimi interessi (per mantenere il nostro sito web aggiornato e rilevante, per sviluppare il nostro business e per informare la nostra strategia di marketing)
-   **Per inviarti newsletter ed email promozionali che potrebbero interessarti**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica · Informazioni professionali o legate all'occupazione · Informazioni educative · \[Informazioni sensibili\]
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Consenso. Quando hai fornito un consenso esplicito all'uso dei dati personali per fornirti contenuti rilevanti e personalizzati · Necessari per i nostri legittimi interessi (per sviluppare il nostro servizio e far crescere il nostro business)
-   **Vendere o condividere dati personali con terze parti a fini di marketing e pubblicità**
    -   · Identificatori · Categorie di Informazioni Personali elencate nello Statuto sui Registri dei Clienti della California · Caratteristiche di Classificazione Protetta secondo la legge californiana o federale · Attività su Internet o Rete Elettronica · Informazioni professionali o legate all'occupazione · Informazioni educative · \[Informazioni sensibili\]
    -   · Interazioni dirette · Tecnologie automatizzate · Biscotti · Informazioni sul file di registro · Gif trasparenti · Terze parti o fonti pubbliche
    -   · Consenso. Quando hai fornito un consenso esplicito all'uso di dati personali per fornirti contenuti rilevanti e personalizzati. · Interesse legittimo. In particolare, al di fuori del SEE possiamo vendere o condividere dati come consentito dalle leggi locali applicabili, il che potrebbe non richiedere il tuo consenso per farlo.

#### 6. DIVULGAZIONE DEI TUOI DATI PERSONALI

Possiamo divulgare qualsiasi categoria di dati personali che raccogliamo da te e su di te nelle seguenti circostanze, per gli scopi indicati nella tabella sopra. Negli ultimi 12 mesi, abbiamo venduto e/o condiviso tutte le categorie di dati personali sopra elencate con Zeta Global e i nostri Business Partners (come descritto di seguito).
​

**Zeta Global**

I dati degli utenti Disqus vengono condivisi con la nostra società affiliata Zeta Global per scopi di marketing, inclusa la pubblicità comportamentale cross-contest. I dati Disqus vengono uniti da Zeta con dati provenienti da altre fonti online e offline per informare decisioni di marketing e messaggi a nome dei clienti inserzionisti di Zeta (marchi e agenzie). Zeta può anche condividere i dati ottenuti da Disqus con questi clienti inserzionisti.
​

**Partner Commerciali**

Condividiamo inoltre dati personali con terze parti esterne per scopi simili a quelli di Zeta Global. Queste terze parti includono altre aziende che partecipano alla pubblicità comportamentale cross-contest, e broker di dati che possono ulteriormente vendere o condividere i tuoi dati personali. Potremmo anche condividere dati personali con inserzionisti con cui hai già condiviso le categorie di interesse dati basate sulla tua attività online, a fini di marketing. Per un elenco di terze parti esterne, si veda [Ecco](#cookies-and-data-recipients).
​

**Fornitori di servizi**

Facciamo affidamento su fornitori di servizi (ad esempio aziende che offrono servizi di hosting web o di analisi) per il funzionamento del nostro sito web e servizio, e questi fornitori potrebbero avere accesso ai tuoi dati personali. I nostri fornitori di servizi sono contrattualmente vietati dall'utilizzare i tuoi dati personali per i propri scopi e sono tenuti a trattarli e a mantenerne la riservatezza nello stesso modo in cui noi.
​

**Transazioni commerciali**

Potremmo cercare di acquisire altre imprese o fonderci con loro. Se dovesse avvenire un cambiamento nella nostra attività, la nuova azienda potrà utilizzare i tuoi dati personali nello stesso modo indicato in questa Informativa sulla privacy. Potremmo anche impegnarci in una condivisione limitata dei dati per valutare e testare i nostri potenziali partner di dati. Quando tali test avvengono effettuati, i dati vengono criptati o codificati prima di essere condivisi. Inoltre, tale condivisione è soggetta a disposizioni contrattuali appropriate e i dati vengono prontamente cancellati dopo il completamento dei test.
​

**Requisiti legali**

Inoltre, possiamo accedere, preservare e divulgare i tuoi dati personali, se riteniamo che ciò sia richiesto dalla legge, dall'ordine del tribunale o da altri processi legali validi. Possiamo anche accedere, preservare e divulgare tali dati personali se riteniamo in buona fede che la divulgazione sia necessaria per proteggere i vostri diritti, la nostra o quella altrui, la vostra proprietà o la sicurezza, o per indagare sulle frodi.
​

Richiediamo a tutte le terze parti con cui condividiamo i dati di adottare le misure appropriate per garantire la sicurezza dei tuoi dati personali e trattarli in conformità con tutte le leggi applicabili.

#### 7. CONSERVAZIONE DEI DATI

Conserveremo i tuoi dati personali solo per il tempo necessario per soddisfare gli scopi per cui li abbiamo raccolti, incluso il rispetto di eventuali requisiti legali, contabili o di rendicontazione.
​

Per determinare il periodo di conservazione appropriato dei dati personali, consideriamo la quantità di dati, la natura e la sensibilità dei dati personali, il potenziale rischio di danni derivanti da un uso o divulgazione non autorizzata dei tuoi dati personali, gli scopi per cui trattiamo i tuoi dati personali e se possiamo raggiungere tali scopi con altri mezzi, nonché i requisiti legali applicabili.

#### 8. SICUREZZA

Disqus utilizza garanzie commercialmente ragionevoli per preservare l'integrità e la sicurezza di tutte le informazioni raccolte tramite il Servizio. Per proteggere la tua privacy e sicurezza, adottiamo misure ragionevoli (come richiedere una password unica) per verificare la tua identità prima di concederti l'accesso al tuo account. Sei responsabile di mantenere la segretezza della tua password uniche e delle informazioni dell'account, e di controllare l'accesso alle tue comunicazioni email da Disqus. Disqus non è responsabile della funzionalità o delle misure di sicurezza di terze parti.

#### 9. I TUOI DIRITTI

A seconda delle leggi in cui vivi, potresti avere uno o più dei seguenti diritti previsti dalle leggi locali. Disqus estende questi diritti a tutti gli individui, indipendentemente da dove vivano, incluso il diritto di:

1\. **Richiedi una copia** dei dati personali che abbiamo raccolto su di te;

2\. **Rinunciare alla nostra vendita o condivisione** di dati personali su di te;

3\. **Rinuncia al ricevere email** dal Servizio;

4\. **Chiediamo di eliminare** i dati raccolti su di te;

5\. **Chiediamo di correggere** dati errati (errate);

6\. **Chiedi che limitiamo l'uso o che cancelliamo dati personali sensibili** che potresti aver precedentemente fornito in un commento (a causa di limitazioni tecniche, tutte queste richieste saranno gestite come richieste di cancellazione dei dati).

Disqus non ti discriminerà per aver esercitato i tuoi diritti alla privacy. Prima di soddisfare una richiesta di copia delle tue informazioni personali, Disqus è tenuto a verificare ragionevolmente la tua identità, cosa che generalmente facciamo inviando un link di verifica all'indirizzo email associato alle tue informazioni personali. Secondo le leggi applicabili, puoi utilizzare un agente autorizzato per fare una richiesta per tuo conto, ma tale agente deve essere in grado di completare il nostro processo di verifica per dimostrare di essere autorizzato a fare la richiesta.
​

Hai inoltre il diritto in molti paesi di contattare l'autorità competente per la privacy o la protezione dei dati se ritieni che non stiamo rispettando le leggi sulla privacy e non siamo riusciti a risolvere la situazione a tua soddisfazione. Se sei residente nello Spazio Economico Europeo, puoi trovare l'autorità locale per la protezione dei dati Ecco. I residenti del Regno Unito possono contattare l'Ufficio del Commissario per l'Informazione Ecco.
​

Visita la nostra pagina Scelte sulla Privacy Pagina delle Scelte sulla Privacy Per maggiori dettagli su questi diritti o per esercitarli.
​

**Richieste di diritti**

Secondo la legge californiana, siamo tenuti a pubblicare statistiche su quante persone hanno esercitato il proprio diritto alla privacy nell'anno precedente. Le seguenti statistiche sono rivolte a chi ha fatto richieste relative a Disqus durante i dodici mesi terminati il 31 dicembre 2025. Queste statistiche coprono tutte le richieste ricevute da individui in tutto il mondo.
​

· RICHIESTE DI CONOSCENZA / ACCESSO AI DATI: Ricevute: 258 \| Rispettato: 258 \| Negato: 0 \| Tempo di risposta mediano: 1,19 giorni \| Tempo medio di risposta: 1,78 giorni

· RICHIESTE DI CANCELLAZIONE: Ricevute: 1.175 \| Rispettato: 1.175 \| Negato: 0 \| Tempo di risposta mediano: 0,83 giorni \| Tempo medio di risposta: 1,33 giorni

· RICHIESTE DI DISATTIVAZIONE PER LA VENDITA O LA CONDIVISIONE: Ricevute: 1.303 \| Completato: 1.303 \| Negato: 0 \| Tempo di risposta mediano: 0,85 giorni \| Tempo medio di risposta: 1,34 giorni

· RICHIESTE DI LIMITARE L'USO DI INFORMAZIONI PERSONALI SENSIBILI: Ricevute: 280 \| Rispettato: 280 \| Negato: 0 \| Tempo di risposta mediano: 1,00 giorno \| Tempo medio di risposta: 1,44 giorni

· RICHIESTE NON ESAMINATE PERCHÉ IL RICHIEDENTE NON HA COMPLETATO LA VERIFICA: 595

· RICHIESTE DI CANCELLAZIONE NON SODDISFATTE IN TUTTO O IN PARTE: 0 — Disqus non ha negato alcuna richiesta di cancellazione nel 2025.

#### 10. TRASFERIMENTI INTERNAZIONALI DI DATI

Per gli utenti con sede nello Spazio Economico Europeo (SEE) e nel Regno Unito, possiamo condividere i vostri dati personali all'interno del Gruppo Disqus o con terze parti esterne. Questo può comportare il trasferimento dei tuoi dati al di fuori del SEE. In particolare, i dati saranno elaborati da team di supporto tecnico negli Stati Uniti, in India e nelle Filippine.
​

Ogni volta che i tuoi dati personali vengono trattati al di fuori del SEE, li proteggiamo con contratti specifici approvati dalla Commissione Europea e/o le disposizioni contrattuali equivalenti approvate dall'Information Commissioner's Office del Regno Unito, che garantiscono che i dati personali ricevano la stessa protezione che hanno in Europa, indipendentemente dal luogo in cui vengono trattati.

#### 11. NON TRACCIARE / CONTROLLO GLOBALE DELLA PRIVACY ("GPC")

Disqus riconosce ed elabora le preferenze dell'utente come impostato nei segnali "Non tracciare" basati su browser tramite la funzione di Controllo Globale della Privacy ("GPC").

#### 12. GENERALE

**Contatto**

Se hai domande su questa Informativa sulla Privacy, inviaci una mail al privacy@disqus.com o contattaci per posta al 3 Park Avenue, 33rd Floor, New York, NY 10016.
​

**Modifiche all'Informativa sulla Privacy**

Disqus può, a sua esclusiva discrezione, modificare o aggiornare questa Informativa sulla Privacy di tanto in tanto, quindi dovresti consultare periodicamente questa pagina. Quando modificheremo la politica, aggiorneremo la data di 'ultima modifica' in cima a questa pagina. Il tuo utilizzo continuativo del Sito dopo la pubblicazione di qualsiasi modifica a questa politica significa che accetti tali modifiche.
​

Click [Ecco](#terms-of-service) per consultare i Termini di Servizio.

La nostra Privacy Policy è disponibile anche nelle seguenti lingue:
[Deutsch](#disqus-datenschutzrichtlinie)

[English](#disqus-privacy-policy)
[Español](#politica-de-privacidad-de-disqus)
[Français](#politique-de-confidentialite-de-disqus)
[Português](#politica-de-privacidade-do-disqus)

### Disqus Privacy Policy {#disqus-privacy-policy}

**Updated** July 10, 2026

This Privacy Policy tells you how Disqus collects, uses, sells, discloses and protects data relating to you (the “User”) in connection with our Service (as defined below), as well as your choices regarding our collection and use of this data.

#### 1. INTRODUCTION

**Overview**

Disqus offers an online public comment and opinion sharing platform where users login and create profiles to participate in conversations with peers and enjoy an interactive experience in Disqus comment sections, polls, and other interactive features that are provided on this site, as well as embedded in third-party sites. Use of our platform and software, and interaction with our cookies or similar tracking technologies (collectively the “Service”), whether on this site or on a third-party site, is subject to the terms of this Privacy Policy. The Service is a public platform and Disqus or others may search for, see, use, or re-post any of your User Content (as defined in our Terms of Use) that you post through the Service. Disqus is also a marketing and data company, and uses and shares personal data collected from third party sites where our Service is enabled for marketing purposes, including cross-context behavioral advertising. For more information on our marketing activities, please see Section 4: Targeted Advertising and Ad Partners below.
​

**Applicability to third-party websites and services**

Disqus offers an online engagement service that other websites use to enable discussion and interactivity among their users. This Privacy Policy applies to the data Disqus collects about Users of the Service and through cookies on Service-enabled websites, and does not apply to the independent data collection practices of any website that uses the Service or other website linked to from the Service. For information about how third-party websites collect and use your personal information, please refer to those websites’ privacy policies.
​

**Your Privacy Rights**

You have rights over your personal information. These rights are described in greater detail in Section 9: Your Rights, below. You can exercise your data privacy rights at here.
​

**Children’s Privacy**

The Service is not intended for use by children under the age of 18. We do not knowingly collect or sell personal information from children under 18 or knowingly allow such persons to register for an account on the service. In the event that we learn that we have collected personal information from a child under the age of 18, we will delete it. If you believe that we might have collected personal information from a child under 18, please contact us, or submit a data privacy rights request at here.
​

**Our Use of Artificial Intelligence**

Disqus uses artificial intelligence and machine learning in two ways. First, AI is used to help moderate content on the platform — detecting spam and content that violates our community guidelines so that the Service can be kept safe and functional for users. Second, AI is used to help deliver more relevant advertising to you based on your interests and online activity. In both cases, Disqus does not use AI to make decisions that have legal or similarly significant effects on individuals. Disqus does not use or sell consumer personal data for the purpose of training large language models. Our AI systems operate in accordance with our privacy commitments as described in this policy.

#### 2. THE DATA WE COLLECT ABOUT YOU

Personal data, or personal information, means any information about a person that can reasonably be linked, directly or indirectly, with a specific person or household.
​

We collect, and have collected in the past twelve (12) months, the following kinds of personal data about Users:

1\. **Identifiers (“Identity Data”)** such as first name, last name, username or similar identifier, internet protocol (IP) address, unique Cookie ID, Device ID, date of birth, email address, telephone number, and mailing address;

2\. **Personal Information categories listed in California Customer Records Statute (Cal. Civ. Code § 1798.80(e)),** which includes any information that identifies, relates to, describes, is capable of being associated with a particular individual but is not publicly available to the general public from federal, state, or local government records;

3\. **Protected classification characteristics** under California or federal law, such as race, gender, or age;

4\. **Internet or other electronic network activity information,** such as browsing and comment history, feedback and survey responses, your login data, browser type and version, time zone setting and location, browser plug-in types and versions, operating system and platform and other technology on the devices you use to access the Service;

5\. **Professional or employment-related information,** to the extent that you include it in your profile or comments, or that it can be inferred from the pages that you view;

6\. **Educational information,** to the extent that you include it in your profile or comments, or that it can be inferred from the pages that you view;

7\. **Sensitive Personal Information.** We do not intentionally collect any personal data about your race or ethnicity, religious or philosophical beliefs, sex life, sexual orientation, political opinions, trade union membership, information about your health or genetic or biometric data, or information about criminal convictions and offences. However, if you make comments using the Service that include such data about yourself it will be publicly available and may be processed by Disqus or others. Additionally, we collect and share information collected through cookies or similar tracking technology about the web pages you have viewed, which may allow the third-parties with whom we share your information to make inferences about you that could constitute sensitive personal information.
​

We may also combine, de-identify, or aggregate any of the information we collect through our Service for any of the purposes described below.

#### 3. HOW IS YOUR PERSONAL DATA COLLECTED?

We use different methods to collect data from and about you including via:
**Direct interactions**

This includes personal data you provide when you create an account or leave a comment.
​

**Automated technologies or interactions**

As you interact with our Service, we may automatically collect Technical Data about your equipment, browsing actions and patterns. Specifically, in the following ways:
​

**Cookies**

A cookie is a small digital file placed on the hard drive of your computer. You may refuse to accept browser cookies by activating settings on your browser, or, in some cases, by interacting with pop-up cookie banners. We use cookies placed on third party websites on which the Service is enabled to collect information about how you interact with those websites even if you do not leave comments, respond to polls, or otherwise directly interact with the Service on those websites, as well as information about the other websites that you visit. Disqus uses cookies and allows partners to also set cookies through the Service in order to facilitate cross-context behavioral advertising. What this means in practice is that we use cookies to help determine which ads you see online, by logging your visits to many advertiser / brand websites, and then showing you ads for similar products and services.
​

Disqus uses ‘authentication’ cookies, e.g., sessionid, disqusauth, and disqusauths, to keep you logged in from your web browser and personalize your Disqus experience.
​

Disqus uses ‘unique’ cookies, e.g., disqus_unique and \_jid, to associate web-based activities with a page load and with a web browser, and understand your interests and product usage.
​

When Disqus loads ads, we use ad serving technologies from Google that may set cookies for the purposes of personalized marketing, associating ads with later activities, and limiting how often you are shown specific ads.
​

**Log File Information**

Server logs collect technical data such as your IP address, browser type, and information about the number of clicks and how you interact with links on the Service, partner sites, domain names, landing pages, pages viewed, and other such information.
​

**Pixels and Similar Trackers**

When you use the Service, we employ pixels, clear gifs (also known as web beacons) which are used to collect Technical Data and information such as online usage patterns. We also use clear gifs in HTML-based emails sent to our users to track which emails are opened and which links or advertisements are clicked by recipients. We also use pixels and clear gifs on third party websites on which the Service is enabled to collect information about how you interact with those websites, even if you do not leave comments, respond to polls, or otherwise directly interact with the Service on those websites.
​

**Third parties or publicly available sources**

We obtain or receive personal data about you from analytics providers such as Google; advertising partners [here](#cookies-and-data-recipients); and third-party data brokers who sell personal data. We obtain or receive personal data from third party connections or log-ins through social media platforms such as Facebook Connect, Google or Twitter/X when you “follow,” “like,” or link your account to the Service. Please note that some of these providers, in particular analytics providers like Google, may process data from European Economic Area (EEA) residents outside the EEA.
​

**Managing Cookies in Your Browser**

You may be able to adjust your browser settings to manage your cookie preferences, such as setting your browser to notify you when you receive a cookie and give you the choice to decide whether or not to accept it. If you reject cookies, you may still use our site, but the functionality of some areas may be limited.

Below are links to information about managing your cookie preferences in common browsers:

· Google Chrome Cookies: Google Chrome Cookies

· Mozilla Firefox Cookies: Mozilla Firefox Cookies

· Internet Explorer Cookies: Internet Explorer Cookies

· Safari Cookies: Safari Cookies

· Google Analytics: Google Analytics

#### 4. TARGETED ADVERTISING AND AD PARTNERS

**Targeted Advertising**

Advertising is the predominant way Disqus makes money. Advertising revenue allows Disqus to operate, support and improve the Service. Disqus uses, and shares with third party ad partners and affiliates, cookie IDs, device IDs (including mobile), hashed email addresses, IP address, Internet Service Provider (ISP) and browser information, demographic or interest data, content viewed and actions taken on the Service, on partner sites, or on other third party sites. This includes information about the websites you’ve viewed and advertisements you’ve interacted with in order to provide you with more relevant advertising targeted to your preferences and interests. This may be derived from your interaction with the Service, partner sites or other third party websites. For a list of third-party ad partners that Disqus is currently working with see [here](#cookies-and-data-recipients).
​

**Email Marketing**

Disqus may also send you email newsletters and email marketing messages if you have provided us with permission, or consented to receive such emails, as required in the jurisdiction in which you reside. Email marketing messages may be tailored to your interests based on the information described above in this section. For information about how to opt-out and to exercise your privacy rights please see [here](#updating-your-account-settings).
​

**Ad Partners and Third Party Disclosures**

We partner and share data with third parties that collect information across various channels, including offline and online, for purposes of delivering more relevant advertising to you or your business. Our partners use this information to recognize you across different channels and platforms, over time, (including but not limited to, computers, mobile devices, addressable TV, or other media), for marketing, analytics, attribution, and reporting purposes. While Disqus does not make inferences based on your data, our advertising partners may draw inferences from your data to understand your preferences, characteristics, psychological trends, predispositions, behavior, attitudes, intelligence, abilities, and aptitudes.
​

In the past twelve (12) months, we have shared or sold the following data:

· Identifiers;

· Personal Information categories listed in California Customer Records Statute (Cal. Civ. Code § 1798.80(e);

· Protected classification characteristics;

· Internet or other electronic network activity information;

· Professional or employment-related information;

· Educational information; and,

· Sensitive Personal Information.

#### 5. HOW WE USE PERSONAL DATA AND OUR LEGAL BASIS FOR USE

In the EEA, United Kingdom (UK), and Brazil, we typically rely on your consent to use personal data. In some cases, we use data as necessary for our legitimate interests (or those of a third party), but only where your rights and interests are not negatively impacted. Finally, we may also use it where we need to comply with a legal or regulatory obligation, or to protect the health, safety, or legal rights of any person.
​

We do not utilize personal information in furtherance of solely automated decisions that produce legal or similarly significant effects concerning you, known as “profiling” under certain laws.
​

We have set out below a description of the ways we use personal data, and which of the legal bases we rely on to do so. We have also identified what our legitimate interests are where appropriate.
​

Note that we may process your personal data for more than one lawful ground depending on the specific purpose for which we are using your data. Please contact us if you have any additional questions.

-   **Purpose or Activity**
    -   Type of Personal Data
    -   Source of Personal Data
    -   Basis of Processing
-   **To register you as a new user**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Internet or Electronic Network Activity
    -   · Direct Interactions · Automated Technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Performance of a contract with you · With your consent
-   **To manage our relationship with you which may include notifying you about changes to our terms or privacy policy, asking you to leave a review or take a survey**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Performance of a contract with you · Necessary to comply with a legal obligation · Necessary for our legitimate interests (to keep our records updated and to study how customers use our Service) · With your consent
-   **To administer and protect our business and the Service (including troubleshooting, data analysis, testing, support, reporting and hosting of data)**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Necessary to comply with a legal obligation · Necessary for our legitimate interests (e.g. to keep our records updated and to study how customers use our Service)
-   **To deliver relevant content and advertisements to you and measure or understand the effectiveness of the advertising we serve you**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity · Professional or Employment Related Information · Educational Information · \[Sensitive Information\]
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Consent. When you have provided explicit consent to our use of personal data to provide you with relevant and personalized content · Necessary for our legitimate interests
-   **To use data analytics to improve our Service**
    -   · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Necessary for our legitimate interests (to keep our website updated and relevant, to develop our business and to inform our marketing strategy)
-   **To send you newsletters and promotional emails that may interest you**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity · Professional or Employment Related Information · Educational Information · \[Sensitive Information\]
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Consent. When you have provided explicit consent to our use of personal data to provide you with relevant and personalized content · Necessary for our legitimate interests (to develop our Service and grow our business)
-   **To sell or share personal data with third parties for marketing and advertising purposes**
    -   · Identifiers · Personal Information Categories listed in the California Customer Records Statute · Protected Classification Characteristics under California or federal law · Internet or Electronic Network Activity · Professional or Employment Related Information · Educational Information · \[Sensitive Information\]
    -   · Direct Interactions · Automated technologies · Cookies · Log File Information · Clear Gifs · Third parties or publicly available sources
    -   · Consent. When you have provided explicit consent to our use of personal data to provide you with relevant and personalized content. · Legitimate interest. In particular outside the EEA we may sell or share data as allowed under applicable local laws, which may not require your consent for us to do so.

#### 6. DISCLOSURES OF YOUR PERSONAL DATA

We may disclose any of the categories of personal data that we collect from and about you in the following circumstances, for the purposes set out in the table above. In the past 12 months, we have sold and/or shared all of the categories of personal data listed above with Zeta Global and our Business Partners (as described below).
​

**Zeta Global**

Disqus user data is shared with our affiliated company Zeta Global for marketing purposes, including cross-context behavioral advertising. Disqus data is merged by Zeta with data from other online and offline sources to inform marketing decisions and messages on behalf of Zeta’s advertiser customers (brands and agencies). Zeta may also share data obtained from Disqus with these advertiser customers.
​

**Business Partners**

We also share personal data with external third parties for similar purposes as with Zeta Global. These third parties include other companies who participate in cross-context behavioral advertising, and data brokers who may further sell or share your personal data. We may also share personal data with advertisers with whom you have already shared your data interest categories based on your online activity, for marketing purposes. For a list of external third parties, please see [here](#cookies-and-data-recipients).
​

**Service Providers**

We rely on service providers (e.g. companies that provide web hosting or analytics services) for the operation of our website and Service, and these service providers may have access to your personal data. Our service providers are contractually prohibited from using your personal data for their own purposes, and are required to treat your personal data and maintain its confidentiality in the same way as us.
​

**Business Transactions**

We may seek to acquire other businesses or merge with them. If a change happens to our business, then the new company may use your personal data in the same way as set out in this privacy notice. We may also engage in limited data sharing to evaluate and test our potential data partners. Where such testing takes place, the data is encrypted or encoded before it is shared. Furthermore, such sharing is subject to appropriate contractual provisions and the data is promptly deleted after the testing is completed.
​

**Legal Requirements**

In addition, we may access, preserve, and disclose your personal data, if we believe doing so is required by law, court order or other valid legal processes. We may also access, preserve, and disclose such personal data if we believe in good faith that disclosure is necessary to protect your, our or others’ rights, property, or safety, or to investigate fraud.
​

We require all third parties with whom we share data to take appropriate steps to ensure the security of your personal data and to treat it in accordance with all applicable laws.

#### 7. DATA RETENTION

We will only retain your personal data for as long as necessary to fulfill the purposes we collected it for, including for the purposes of satisfying any legal, accounting, or reporting requirements.

To determine the appropriate retention period for personal data, we consider the amount, nature, and sensitivity of the personal data, the potential risk of harm from unauthorized use or disclosure of your personal data, the purposes for which we process your personal data and whether we can achieve those purposes through other means, and the applicable legal requirements.

#### 8. SECURITY

Disqus uses commercially reasonable safeguards to preserve the integrity and security of all information collected through the Service. To protect your privacy and security, we take reasonable steps (such as requesting a unique password) to verify your identity before granting you access to your account. You are responsible for maintaining the secrecy of your unique password and account information, and for controlling access to your email communications from Disqus. Disqus is not responsible for the functionality or security measures of any third party.

#### 9. YOUR RIGHTS

Depending on the laws where you live, you may have one or more of the following rights under local laws. Disqus extends these rights to all individuals, regardless of where they live, including the right to:

1\. **Request a copy** of personal data we have collected about you;

2\. **Opt-out of our sale or sharing** of personal data about you;

3\. **Opt-out of receiving email** from the Service;

4\. **Request that we delete** data we have collected about you;

5\. **Request that we correct** incorrect data;

6\. **Request that we limit our use of, or delete sensitive personal data** that you may have previously provided in a comment (due to technical limitations, all of these requests will be handled as requests to delete data).

Disqus will not discriminate against you for exercising your privacy rights. Prior to fulfilling a request for a copy of your personal information, Disqus is required to reasonably verify your identity, which we generally do by sending a verification link to the email address associated with your personal information. Under applicable laws, you can use an authorized agent to make a request on your behalf, but that agent must be able to complete our verification process in order to demonstrate that they have been authorized to make the request.

You also have the right in many countries to contact the relevant privacy or data protection authority if you believe we are not complying with privacy laws, and we have not been able to resolve the situation to your satisfaction. If you are a resident of the European Economic Area, you can find your local data protection authority here. UK residents can contact the Information Commissioner's Office here.
​

Visit our Privacy Choices page Privacy Choices page for more details about these rights or to exercise them.
​

**Rights Requests**

Under California law, we are required to publish statistics on how many people exercised their privacy rights in the previous year. The following statistics are for people who made requests relating to Disqus during the twelve months ending December 31, 2025. These statistics cover all requests received from individuals around the world.
​

· REQUESTS TO KNOW / ACCESS DATA: Received: 258 \| Complied with: 258 \| Denied: 0 \| Median response time: 1.19 days \| Mean response time: 1.78 days

· REQUESTS TO DELETE: Received: 1,175 \| Complied with: 1,175 \| Denied: 0 \| Median response time: 0.83 days \| Mean response time: 1.33 days

· REQUESTS TO OPT-OUT OF SALE OR SHARING: Received: 1,303 \| Complied with: 1,303 \| Denied: 0 \| Median response time: 0.85 days \| Mean response time: 1.34 days

· REQUESTS TO LIMIT USE OF SENSITIVE PERSONAL INFORMATION: Received: 280 \| Complied with: 280 \| Denied: 0 \| Median response time: 1.00 day \| Mean response time: 1.44 days

· REQUESTS NOT ACTIONED BECAUSE REQUESTOR DID NOT COMPLETE VERIFICATION: 595

· DELETION REQUESTS NOT FULFILLED IN WHOLE OR IN PART: 0 — Disqus did not deny any deletion requests in 2025.

#### 10. INTERNATIONAL DATA TRANSFERS

For users based in the European Economic Area (EEA) and the UK we may share your personal data within the Disqus Group or with external third parties. This may involve transferring your data outside the EEA. Specifically, data will be processed by technical support teams in the United States, India, and the Philippines.

Whenever your personal data is processed outside of the EEA, we protect it using specific contracts approved by the European Commission and/or the equivalent contractual provisions approved by the UK’s Information Commissioner’s Office which ensures that personal data receives the same protection it has in Europe regardless of where it is processed.

#### 11. DO NOT TRACK / GLOBAL PRIVACY CONTROL (“GPC”)

Disqus recognizes and processes user preferences as set in the browser-based “Do Not Track” signals via the Global Privacy Control (“GPC”) feature.

#### 12. GENERAL

**Contact**

If you have any questions about this Privacy Policy, please email us at privacy@disqus.com, or contact us by mail at 3 Park Avenue, 33rd Floor, New York, NY 10016.
​

**Changes to Privacy Policy**

Disqus may, in its sole discretion, modify or update this Privacy Policy from time to time, and so you should review this page periodically. When we change the policy, we will update the ‘last modified’ date at the top of this page. Your continued use of the Site following the posting of any changes to this policy means you accept such changes.
​

Click [here](#terms-of-service) to view the Terms of Service.

Our Privacy Policy is also available in the following languages:
[Deutsch](#disqus-datenschutzrichtlinie)
[Español](#politica-de-privacidad-de-disqus)
[Français](#politique-de-confidentialite-de-disqus)
[Italiano](#disqus-informativa-sulla-riservatezza)
[Português](#politica-de-privacidade-do-disqus)

### Disqus-Datenschutzrichtlinie {#disqus-datenschutzrichtlinie}

Datenschutzrichtlinie von Disqus

**Aktualisiert** am 10. Juli 2026

Diese Datenschutzrichtlinie erklärt Ihnen, wie Disqus Daten, die sich auf Sie (den "Nutzer") im Zusammenhang mit unserem Service (wie unten definiert) beziehen, sammelt, verkauft, offenlegt und schützt, sowie welche Entscheidungen Sie bezüglich der Sammlung und Nutzung dieser Daten haben.

#### 1. EINLEITUNG

**Überblick**

Disqus bietet eine Online-Plattform für öffentliche Kommentare und Meinungsaustausch, auf der Nutzer sich anmelden und Profile erstellen, um an Gesprächen mit Gleichgesinnten teilzunehmen und eine interaktive Erfahrung in Disqus-Kommentarbereichen, Umfragen und anderen interaktiven Funktionen zu genießen, die auf dieser Seite sowie in Drittanbieter-Seiten eingebettet sind. Die Nutzung unserer Plattform und Software sowie die Interaktion mit unseren Cookies oder ähnlichen Tracking-Technologien (zusammen der "Service"), sei es auf dieser Seite oder auf einer Drittanbieterseite, unterliegt den Bedingungen dieser Datenschutzrichtlinie. Der Dienst ist eine öffentliche Plattform und Disqus oder andere können nach deinen Nutzerinhalten (wie in unseren Nutzungsbedingungen definiert) suchen, sehen, nutzen oder erneut posten, die du über den Dienst postest. Disqus ist außerdem ein Marketing- und Datenunternehmen und nutzt und teilt persönliche Daten, die von Drittanbieterseiten gesammelt werden, auf denen unser Service aktiviert ist, für Marketingzwecke, einschließlich kontextübergreifender verhaltensbezogener Werbung. Weitere Informationen zu unseren Marketingaktivitäten finden Sie unten unter Abschnitt 4: Zielgerichtete Werbung und Werbepartner.
​

**Anwendbarkeit auf Drittanbieter-Websites und -Dienste**

Disqus bietet einen Online-Engagement-Service, den andere Websites nutzen, um Diskussionen und Interaktivität unter ihren Nutzern zu ermöglichen. Diese Datenschutzrichtlinie gilt für die Daten, die Disqus über Nutzer des Dienstes und über Cookies auf dienstfähigen Websites sammelt, und gilt nicht für die unabhängigen Datenerfassungspraktiken irgendeiner Website, die den Dienst oder andere vom Dienst verlinkte Webseiten nutzt. Informationen darüber, wie Drittanbieter-Websites Ihre persönlichen Daten sammeln und verwenden, finden Sie bitte in den Datenschutzrichtlinien dieser Websites.
​

**Ihre Datenschutzrechte**

Du hast Rechte an deinen persönlichen Daten. Diese Rechte werden ausführlicher in Abschnitt 9: Ihre Rechte unten beschrieben. Sie können Ihre Datenschutzrechte unter Hier.
​

**Privatsphäre der Kinder**

Der Gottesdienst ist nicht für Kinder unter 18 Jahren gedacht. Wir sammeln oder verkaufen nicht wissentlich persönliche Informationen von Kindern unter 18 Jahren und erlauben diesen Personen nicht wissentlich, sich für ein Konto auf dem Dienst zu registrieren. Falls wir erfahren, dass wir persönliche Informationen von einem Kind unter 18 Jahren gesammelt haben, werden wir diese löschen. Wenn Sie glauben, dass wir persönliche Informationen von einem Kind unter 18 Jahren gesammelt haben könnten, kontaktieren Sie uns bitte oder stellen Sie einen Antrag auf Datenschutz unter Hier.
​

**Unsere Nutzung von künstlicher Intelligenz**

Disqus nutzt künstliche Intelligenz und maschinelles Lernen auf zwei Arten. Erstens wird KI eingesetzt, um Inhalte auf der Plattform zu moderieren – Spam und Inhalte, die gegen unsere Community-Richtlinien verstoßen, zu erkennen, damit der Dienst für Nutzer sicher und funktionsfähig bleibt. Zweitens wird KI eingesetzt, um Ihnen relevantere Werbung basierend auf Ihren Interessen und Online-Aktivitäten zu liefern. In beiden Fällen nutzt Disqus keine KI, um Entscheidungen zu treffen, die rechtliche oder ähnlich bedeutende Auswirkungen auf Einzelpersonen haben. Disqus verwendet oder verkauft keine persönlichen Verbraucherdaten zum Zweck des Trainings großer Sprachmodelle. Unsere KI-Systeme funktionieren gemäß unseren Datenschutzverpflichtungen wie in dieser Richtlinie beschrieben.

#### 2. DIE DATEN, DIE WIR ÜBER SIE SAMMELN

Personenbezogene Daten oder persönliche Informationen bedeuten alle Informationen über eine Person, die vernünftigerweise direkt oder indirekt mit einer bestimmten Person oder einem bestimmten Haushalt in Verbindung gebracht werden können.
​

Wir sammeln und haben in den letzten zwölf (12) Monaten die folgenden Arten von personenbezogenen Daten über Nutzer gesammelt:

1\. **Identifikatoren ("Identitätsdaten")** wie Vorname, Nachname, Benutzername oder ähnliche Kennung, Internetprotokoll-(IP)-Adresse, eindeutige Cookie-ID, Geräte-ID, Geburtsdatum, E-Mail-Adresse, Telefonnummer und Postadresse;

2\. **Kategorien personenbezogener Daten, die im California Customer Records Statute (Cal. Civ. Code § 1798.80(e)) aufgeführt sind,** welche alle Informationen umfassen, die eine bestimmte Person identifizieren, betreffen, beschreiben, mit einer bestimmten Person in Verbindung gebracht werden können, aber der Öffentlichkeit nicht öffentlich aus Bundes-, Landes- oder Kommunalregistern zugänglich sind;

3\. **Geschützte Klassifizierungsmerkmale** nach kalifornischem oder Bundesrecht, wie Rasse, Geschlecht oder Alter;

4\. **Internet- oder andere elektronische Netzwerkaktivitäten,** wie Browsing- und Kommentarverlauf, Rückmeldungen und Umfrageantworten, Ihre Anmeldedaten, Browsertyp und -version, Zeitzoneneinstellungen und -standorte, Browser-Plug-in-Arten und -versionen, Betriebssystem und Plattform sowie andere Technologien auf den Geräten, die Sie zum Zugriff auf den Dienst verwenden;

5\. **Berufliche oder arbeitsbezogene Informationen,** soweit Sie sie in Ihr Profil oder Ihre Kommentare aufnehmen oder aus den Seiten, die Sie ansehen, abgeleitet werden können;

6\. **Bildungsinformationen,** soweit Sie sie in Ihr Profil oder Ihre Kommentare aufnehmen oder aus den Seiten, die Sie ansehen, abgeleitet werden können;

7\. **Sensible persönliche Informationen.** Wir sammeln absichtlich keine persönlichen Daten zu Ihrer Rasse oder Ethnie, Ihren religiösen oder philosophischen Überzeugungen, Ihrem Sexualleben, Ihrer sexuellen Orientierung, politischen Ansichten, Gewerkschaftsmitgliedschaft, Informationen über Ihre Gesundheit oder genetische oder biometrische Daten oder Informationen zu strafrechtlichen Verurteilungen und Straftaten. Wenn Sie jedoch Kommentare über den Dienst abgeben, die solche Daten über sich selbst enthalten, sind diese öffentlich zugänglich und können von Disqus oder anderen verarbeitet werden. Zusätzlich sammeln und teilen wir Informationen, die wir über Cookies oder ähnliche Tracking-Technologien über die von Ihnen angesehenen Webseiten gesammelt haben, was es den Drittanbietern, mit denen wir Ihre Informationen teilen, ermöglichen kann, Schlussfolgerungen über Sie zu ziehen, die sensible persönliche Informationen darstellen könnten.

Wir können auch alle Informationen, die wir über unseren Service sammeln, für die unten beschriebenen Zwecke kombinieren, deidentifizieren oder aggregieren.

#### 3. WIE WERDEN IHRE PERSÖNLICHEN DATEN GESAMMELT?

Wir verwenden verschiedene Methoden, um Daten von und über Sie zu sammeln, unter anderem durch:

**Direkte Wechselwirkungen**

Dazu gehören persönliche Daten, die Sie angeben, wenn Sie ein Konto erstellen oder einen Kommentar hinterlassen.

**Automatisierte Technologien oder Interaktionen**

Während Sie mit unserem Service interagieren, können wir automatisch technische Daten zu Ihrer Ausrüstung, Browsing-Aktionen und -Mustern sammeln. Genauer gesagt auf folgende Weise:
​

**Kekse**

Ein Cookie ist eine kleine digitale Datei, die auf die Festplatte Ihres Computers gelegt wird. Sie können die Akzeptanz von Browser-Cookies ablehnen, indem Sie Einstellungen in Ihrem Browser aktivieren oder in manchen Fällen über Pop-up-Cookie-Banner interagieren. Wir verwenden Cookies, die auf Drittanbieter-Websites platziert werden, auf denen der Dienst aktiviert ist, um Informationen darüber zu sammeln, wie Sie mit diesen Websites interagieren, auch wenn Sie keine Kommentare hinterlassen, auf Umfragen reagieren oder anderweitig nicht direkt mit dem Dienst interagieren, sowie Informationen über die anderen von Ihnen besuchten Webseiten. Disqus verwendet Cookies und ermöglicht es Partnern, auch Cookies über den Dienst zu setzen, um kontextübergreifende Verhaltenswerbung zu ermöglichen. In der Praxis bedeutet das, dass wir Cookies verwenden, um zu bestimmen, welche Werbung Sie online sehen, indem wir Ihre Besuche auf vielen Werbe- und Markenwebsites protokollieren und Ihnen dann Anzeigen für ähnliche Produkte und Dienstleistungen anzeigen.

Disqus verwendet 'Authentifizierungs'-Cookies, z. B. sessionid, disqusauth und Disqusauths, um Sie im Webbrowser einloggen zu lassen und Ihr Disqus-Erlebnis zu personalisieren.
​

Disqus verwendet 'einzigartige' Cookies, z. B. disqus_unique und \_jid, um webbasierte Aktivitäten mit einem Seitenladen und einem Webbrowser zu verknüpfen und Ihre Interessen und Produktnutzung zu verstehen.

Wenn Disqus Anzeigen lädt, nutzen wir Ad-Serving-Technologien von Google, die Cookies für personalisiertes Marketing setzen, Werbung mit späteren Aktivitäten verknüpfen und die Häufigkeit bestimmter Anzeigen begrenzen.
​

**Logdatei-Informationen**

Serverprotokolle sammeln technische Daten wie Ihre IP-Adresse, Browsertyp sowie Informationen über die Anzahl der Klicks und wie Sie mit Links auf dem Dienst, Partnerseiten, Domainnamen, Landingpages, angesehenen Seiten und weiteren Informationen interagieren.
​

**Pixel und ähnliche Tracker**

Wenn Sie den Dienst nutzen, verwenden wir Pixel, klare GIFs (auch Web-Beacons genannt), die zur Sammlung technischer Daten und Informationen wie Online-Nutzungsmuster verwendet werden. Wir verwenden auch klare GIFs in HTML-basierten E-Mails, die an unsere Nutzer gesendet werden, um zu verfolgen, welche E-Mails geöffnet werden und welche Links oder Werbung von den Empfängern angeklickt werden. Wir verwenden außerdem Pixel und klare GIFs auf Drittanbieter-Webseiten, auf denen der Dienst Informationen darüber sammeln kann, wie Sie mit diesen Websites interagieren, selbst wenn Sie keine Kommentare hinterlassen, auf Umfragen reagieren oder auf diese Websites nicht direkt mit dem Dienst interagieren.
​

**Dritte oder öffentlich zugängliche Quellen**

Wir erhalten persönliche Daten über Sie von Analyseanbietern wie Google; Werbepartner [Hier](#cookies-and-data-recipients); und Drittanbieter-Datenhändler, die personenbezogene Daten verkaufen. Wir erhalten persönliche Daten von Drittanbieter-Verbindungen oder Anmeldungen über soziale Medienplattformen wie Facebook Connect, Google oder Twitter/X, wenn Sie Ihrem Konto folgen, "liken" oder mit dem Dienst verknüpfen. Bitte beachten Sie, dass einige dieser Anbieter, insbesondere Analyseanbieter wie Google, Daten von Einwohnern des Europäischen Wirtschaftsraums (EWR) außerhalb des EWR verarbeiten können.
​

**Cookies in Ihrem Browser verwalten**

Sie können möglicherweise Ihre Browsereinstellungen anpassen, um Ihre Cookie-Präferenzen zu verwalten, zum Beispiel so einzustellen, dass Ihr Browser Sie benachrichtigt, wenn Sie ein Cookie erhalten, und Ihnen die Möglichkeit geben, zu entscheiden, ob Sie es akzeptieren möchten oder nicht. Wenn Sie Cookies ablehnen, können Sie unsere Seite weiterhin nutzen, aber die Funktionalität einiger Bereiche kann eingeschränkt sein.

Nachfolgend finden Sie Links zu Informationen zur Verwaltung Ihrer Cookie-Präferenzen in gängigen Browsern:

· Google Chrome Cookies: Google Chrome Cookies

· Mozilla Firefox Cookies: Mozilla Firefox Cookies

· Internet Explorer Cookies: Internet Explorer Cookies

· Safari-Kekse: Safari-Cookies

· Google Analytics: Google Analytics
​

#### 4. GEZIELTE WERBE- UND WERBEPARTNER

**Gezielte Werbung**

Werbung ist die vorherrschende Art und Weise, wie Disqus Geld verdient. Werbeeinnahmen ermöglichen es Disqus, den Service zu betreiben, zu unterstützen und zu verbessern. Disqus nutzt und teilt mit Drittanbieter-Werbepartnern und Partnern sowie Partnern Cookie-IDs, Geräte-IDs (einschließlich Mobilgeräten), gehashten E-Mail-Adressen, IP-Adressen, Internetdienstanbieter- (ISP) und Browserdaten, demografische oder Interessendaten, angesehene Inhalte und ergriffene Maßnahmen auf dem Dienst, Partnerseiten oder anderen Drittanbieterseiten. Dazu gehören Informationen über die besuchten Websites und Anzeigen, mit denen du interagiert hast, um dir relevantere Werbung zu bieten, die auf deine Vorlieben und Interessen zugeschnitten ist. Dies kann sich aus Ihrer Interaktion mit dem Service, Partnerseiten oder anderen Drittanbieter-Websites ergeben. Eine Liste von Drittanbieter-Werbepartnern, mit denen Disqus derzeit zusammenarbeitet, finden Sie unter [Hier](#cookies-and-data-recipients).
​

**E-Mail-Marketing**

Disqus kann Ihnen auch E-Mail-Newsletter und Marketing-Nachrichten senden, wenn Sie uns die Erlaubnis erteilt oder zugestimmt haben, diese E-Mails zu erhalten, wie es in Ihrer Region vorgeschrieben ist. E-Mail-Marketing-Nachrichten können auf Grundlage der oben in diesem Abschnitt beschriebenen Informationen auf Ihre Interessen zugeschnitten werden. Informationen darüber, wie Sie sich abmelden und Ihre Datenschutzrechte ausüben können, finden Sie bitte unter [Hier](#updating-your-account-settings).
​

**Werbepartner und Drittanbieter-Offenlegungen**

Wir arbeiten mit Dritten zusammen und teilen Daten, die Informationen über verschiedene Kanäle, einschließlich offline und online, sammeln, um Ihnen oder Ihrem Unternehmen relevantere Werbung zu liefern. Unsere Partner nutzen diese Informationen, um Sie über verschiedene Kanäle und Plattformen hinweg zu erkennen (einschließlich, aber nicht beschränkt auf, Computer, mobile Geräte, adressierbares Fernsehen oder andere Medien), für Marketing-, Analyse-, Attributions- und Berichtszwecke. Obwohl Disqus keine Schlussfolgerungen auf Basis Ihrer Daten zieht, können unsere Werbepartner aus Ihren Daten Schlüsse ziehen, um Ihre Präferenzen, Eigenschaften, psychologischen Trends, Veranlagungen, Verhaltensweisen, Einstellungen, Intelligenz, Fähigkeiten und Fähigkeiten zu verstehen.

In den letzten zwölf (12) Monaten haben wir folgende Daten geteilt oder verkauft:

· Identifikatoren;

· Persönliche Informationskategorien, die im kalifornischen Kundenaktengesetz (Cal. Civ. Code § 1798.80(e) aufgeführt sind;

· Geschützte Klassifikationsmerkmale;

· Internet- oder andere elektronische Netzwerkaktivitäten;

· berufliche oder arbeitsbezogene Informationen;

· Bildungsinformationen; und,

· Sensible persönliche Informationen.

#### 5. WIE WIR PERSONENBEZOGENE DATEN VERWENDEN UND UNSERE RECHTLICHE NUTZUNGSGRUNDLAGE

Im EWR, im Vereinigten Königreich (UK) und in Brasilien verlassen wir uns in der Regel auf Ihre Zustimmung zur Nutzung personenbezogener Daten. In einigen Fällen verwenden wir Daten für unsere legitimen Interessen (oder die eines Dritten) erforderlich, jedoch nur, wenn Ihre Rechte und Interessen nicht negativ beeinträchtigt werden. Schließlich können wir es auch dort verwenden, wo wir einer gesetzlichen oder regulatorischen Verpflichtung nachkommen oder die Gesundheit, Sicherheit oder Rechte einer Person schützen müssen.

Wir verwenden personenbezogene Daten nicht ausschließlich für automatisierte Entscheidungen, die rechtliche oder ähnlich bedeutende Auswirkungen auf Sie haben, was unter bestimmten Gesetzen als "Profiling" bezeichnet wird.

Im Folgenden haben wir eine Beschreibung der Art und Weise dargelegt, wie wir personenbezogene Daten verwenden und auf welche der rechtlichen Grundlagen wir uns dafür stützen. Wir haben auch, wo es angebracht ist, unsere legitimen Interessen identifiziert.
​

Beachten Sie, dass wir Ihre personenbezogenen Daten je nach dem spezifischen Zweck, zu dem wir Ihre Daten verwenden, aus mehreren rechtlichen Gründen verarbeiten können. Bitte kontaktieren Sie uns, wenn Sie weitere Fragen haben.

-   **Zweck oder Tätigkeit**
    -   Art der personenbezogenen Daten
    -   Quelle der personenbezogenen Daten
    -   Grundlage der Verarbeitung
-   **Um dich als neuen Nutzer zu registrieren**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Internet- oder elektronische Netzwerkaktivitäten
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Erfüllung eines Vertrags mit Ihnen · Mit Ihrer Zustimmung
-   **Um unsere Beziehung zu Ihnen zu verwalten, was dazu führen kann, Sie über Änderungen unserer Bedingungen oder Datenschutzrichtlinie zu informieren, Sie zu bitten, eine Bewertung zu hinterlassen oder eine Umfrage auszufüllen.**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Erfüllung eines Vertrags mit Ihnen · Notwendig, um einer gesetzlichen Verpflichtung nachzukommen · Notwendig für unsere berechtigten Interessen (um unsere Unterlagen aktuell zu halten und zu untersuchen, wie Kunden unseren Service nutzen) · Mit Ihrer Zustimmung
-   **Verwaltung und Schutz unseres Geschäfts und des Dienstes (einschließlich Fehlerbehebung, Datenanalyse, Test, Support, Berichterstattung und Hosting von Daten)**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Notwendig, um einer gesetzlichen Verpflichtung nachzukommen · Notwendig für unsere legitimen Interessen (z. B. um unsere Unterlagen aktuell zu halten und zu untersuchen, wie Kunden unseren Service nutzen)
-   **Um Ihnen relevante Inhalte und Werbung zu liefern und die Wirksamkeit der Werbung, die wir Ihnen anbieten, zu messen oder zu verstehen oder zu verstehen**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten · Berufliche oder arbeitsbezogene Informationen · Bildungsinformationen · \[sensible Informationen\]
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Einwilligung. Wenn Sie ausdrücklich zugestimmt haben, dass wir personenbezogene Daten verwenden, um Ihnen relevante und personalisierte Inhalte bereitzustellen, · Notwendig für unsere legitimen Interessen
-   **Datenanalyse zu nutzen, um unseren Service zu verbessern**
    -   · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Notwendig für unsere berechtigten Interessen (um unsere Website aktuell und relevant zu halten, unser Geschäft zu entwickeln und unsere Marketingstrategie zu informieren)
-   **Ihnen Newsletter und Werbe-E-Mails zu senden, die Sie interessieren könnten**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten · Berufliche oder arbeitsbezogene Informationen · Bildungsinformationen · \[sensible Informationen\]
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Einwilligung. Wenn Sie ausdrücklich zugestimmt haben, dass wir personenbezogene Daten verwenden, um Ihnen relevante und personalisierte Inhalte bereitzustellen, · Notwendig für unsere legitimen Interessen (um unseren Service zu entwickeln und unser Geschäft auszubauen)
-   **Den Verkauf oder das Teilen personenbezogener Daten mit Dritten zu Marketing- und Werbezwecken**
    -   · Kennungen · Personenbezogene Datenkategorien, die im kalifornischen Kundenaktengesetz aufgeführt sind · Geschützte Klassifizierungsmerkmale nach kalifornischem oder bundesstaatlichem Recht · Internet- oder elektronische Netzwerkaktivitäten · Berufliche oder arbeitsbezogene Informationen · Bildungsinformationen · \[sensible Informationen\]
    -   · Direkte Wechselwirkungen · Automatisierte Technologien · Kekse · Logdatei-Informationen · Clear Gifs · Dritte oder öffentlich zugängliche Quellen
    -   · Einwilligung. Wenn Sie ausdrücklich zugestimmt haben, dass wir personenbezogene Daten verwenden, um Ihnen relevante und personalisierte Inhalte bereitzustellen. · Legitimes Interesse. Insbesondere außerhalb des EWR dürfen wir Daten verkaufen oder teilen, wie es den geltenden lokalen Gesetzen erlaubt sind, was möglicherweise nicht Ihre Zustimmung erfordert.

#### 6. OFFENLEGUNG IHRER PERSONENBEZOGENEN DATEN

Wir dürfen alle Kategorien personenbezogener Daten, die wir von und über Sie erheben, unter den folgenden Umständen für die oben in der Tabelle genannten Zwecke offenlegen. In den vergangenen 12 Monaten haben wir alle oben genannten Kategorien personenbezogener Daten mit Zeta Global und unseren Geschäftspartnern (wie unten beschrieben) verkauft und/oder geteilt.
​

**Zeta Global**

Disqus-Nutzerdaten werden zu Marketingzwecken, einschließlich kontextübergreifender Verhaltenswerbung, mit unserem verbundenen Unternehmen Zeta Global geteilt. Disqus-Daten werden von Zeta mit Daten anderer Online- und Offline-Quellen zusammengeführt, um Marketingentscheidungen und Botschaften im Namen der Werbekunden von Zeta (Marken und Agenturen) zu unterstützen. Zeta kann auch Daten, die von Disqus gewonnen wurden, mit diesen Werbekunden teilen.
​

**Geschäftspartner**

Wir teilen auch persönliche Daten mit externen Dritten zu ähnlichen Zwecken wie bei Zeta Global. Zu diesen Drittanbietern gehören weitere Unternehmen, die an kontextübergreifender Verhaltenswerbung teilnehmen, sowie Datenhändler, die Ihre persönlichen Daten weiterverkaufen oder weitergeben können. Wir können auch persönliche Daten mit Werbetreibenden teilen, mit denen Sie Ihre Dateninteressenkategorien basierend auf Ihrer Online-Aktivität bereits geteilt haben, zu Marketingzwecken. Eine Liste externer Drittanbieter finden Sie bitte unter [Hier](#cookies-and-data-recipients).
​

**Dienstanbieter**

Wir sind auf Dienstleister (z. B. Unternehmen, die Webhosting- oder Analysedienste anbieten) für den Betrieb unserer Website und unseres Dienstes angewiesen, und diese Anbieter haben möglicherweise Zugriff auf Ihre persönlichen Daten. Unsere Dienstleister sind vertraglich untersagt, Ihre personenbezogenen Daten für eigene Zwecke zu verwenden, und sind verpflichtet, Ihre persönlichen Daten genauso zu behandeln und deren Vertraulichkeit zu wahren wie wir.
​

**Geschäftstransaktionen**

Wir könnten versuchen, andere Unternehmen zu übernehmen oder mit ihnen zu fusionieren. Wenn sich unser Unternehmen ändert, kann das neue Unternehmen Ihre persönlichen Daten auf die gleiche Weise verwenden, wie in dieser Datenschutzerklärung festgelegt. Wir könnten auch begrenzte Daten teilen, um unsere potenziellen Datenpartner zu evaluieren und zu testen. Wo solche Tests stattfinden, werden die Daten verschlüsselt oder kodiert, bevor sie geteilt werden. Darüber hinaus unterliegt eine solche Weitergabe den entsprechenden vertraglichen Bestimmungen und die Daten werden nach Abschluss der Tests umgehend gelöscht.
​

**Rechtliche Anforderungen**

Darüber hinaus können wir auf Ihre persönlichen Daten zugreifen, diese bewahren und offenlegen, wenn wir glauben, dass dies gesetzlich, gerichtlich oder andere gültige rechtliche Verfahren erforderlich ist. Wir dürfen auch auf solche personenbezogenen Daten zugreifen, diese bewahren und offenlegen, wenn wir in gutem Glauben glauben, dass eine Offenlegung notwendig ist, um Ihre, unsere oder fremde Rechte, Eigentum oder Sicherheit zu schützen oder Betrug zu untersuchen.

Wir verlangen von allen Dritten, mit denen wir Daten teilen, dass sie angemessene Maßnahmen ergreifen, um die Sicherheit Ihrer personenbezogenen Daten zu gewährleisten und sie gemäß allen geltenden Gesetzen zu behandeln.

#### 7. DATENSPEICHERUNG

Wir bewahren Ihre persönlichen Daten nur so lange auf, wie es notwendig ist, um die Zwecke zu erfüllen, für die wir sie gesammelt haben, einschließlich der Erfüllung etwaiger rechtlicher, buchhalterischer oder berichterstattender Anforderungen.
​

Um die angemessene Aufbewahrungsfrist für personenbezogene Daten festzulegen, berücksichtigen wir die Menge, Art und Sensibilität der personenbezogenen Daten, das potenzielle Risiko durch unbefugte Nutzung oder Offenlegung Ihrer personenbezogenen Daten, die Zwecke, zu denen wir Ihre personenbezogenen Daten verarbeiten, ob wir diese Ziele auf anderem Wege erreichen können, sowie die geltenden gesetzlichen Anforderungen.

#### 8. SICHERHEIT

Disqus nutzt kommerziell angemessene Schutzmaßnahmen, um die Integrität und Sicherheit aller über den Service gesammelten Informationen zu bewahren. Um Ihre Privatsphäre und Sicherheit zu schützen, ergreifen wir angemessene Maßnahmen (wie die Anforderung eines einzigartigen Passworts), um Ihre Identität zu überprüfen, bevor wir Ihnen Zugang zu Ihrem Konto gewähren. Sie sind verantwortlich für die Geheimhaltung Ihres einzigartigen Passworts und Ihrer Kontoinformationen sowie für die Kontrolle des Zugriffs auf Ihre E-Mail-Kommunikation von Disqus. Disqus ist nicht verantwortlich für die Funktionalität oder Sicherheitsmaßnahmen eines Drittanbieters.

#### 9. DEINE RECHTE

Je nach den Gesetzen, in denen Sie wohnen, haben Sie möglicherweise eines oder mehrere der folgenden Rechte nach lokalen Gesetzen. Disqus erweitert diese Rechte auf alle Einzelpersonen, unabhängig davon, wo sie leben, einschließlich des Rechts:

1\. **Fordern Sie eine Kopie** der persönlichen Daten an, die wir über Sie gesammelt haben;

2\. **Deaktivieren Sie unseren Verkauf oder das Teilen** persönlicher Daten über Sie;

3\. **Abmelden Sie den Empfang von E-Mails** vom Service;

4\. **Bitte darum, dass wir** die über Sie gesammelten Daten löschen;

5\. **Bitte darum, dass wir** falsche Daten korrigieren;

6\. **Bitte darum, dass wir unsere Nutzung von sensiblen personenbezogenen Daten einschränken oder löschen,** die Sie zuvor in einem Kommentar angegeben haben könnten (aufgrund technischer Einschränkungen werden all diese Anfragen als Anfragen zur Löschung von Daten behandelt).

Disqus wird Sie nicht diskriminieren, wenn Sie Ihre Privatsphäre ausüben. Bevor Sie eine Anfrage nach einer Kopie Ihrer persönlichen Daten erfüllen, ist Disqus verpflichtet, Ihre Identität angemessen zu überprüfen, was wir in der Regel tun, indem wir einen Verifizierungslink an die mit Ihren persönlichen Daten verknüpfte E-Mail-Adresse senden. Nach geltenden Gesetzen können Sie einen autorisierten Vertreter beauftragen, um einen Antrag in Ihrem Namen zu stellen, aber dieser Bevollmächtigte muss unseren Verifizierungsprozess abschließen können, um nachzuweisen, dass er zur Anfrage befugt ist.
​

Sie haben in vielen Ländern auch das Recht, die zuständige Datenschutz- oder Datenschutzbehörde zu kontaktieren, wenn Sie der Meinung sind, dass wir die Datenschutzgesetze nicht einhalten und wir die Situation nicht zu Ihrer Zufriedenheit klären konnten. Wenn Sie Einwohner des Europäischen Wirtschaftsraums sind, finden Sie Ihre örtliche Datenschutzbehörde Hier.
​
Einwohner Großbritanniens können das Büro des Informationskommissars kontaktieren Hier.
​

Besuchen Sie unsere Seite mit Datenschutzoptionen Datenschutz-Choices-Seite für weitere Details zu diesen Rechten oder deren Ausübung.
​

**Rechte Anträge**

Nach kalifornischem Recht sind wir verpflichtet, Statistiken darüber zu veröffentlichen, wie viele Menschen im vergangenen Jahr ihre Privatsphärerechte ausgeübt haben. Die folgenden Statistiken beziehen sich auf Personen, die während der zwölf Monate bis zum 31. Dezember 2025 Anfragen zu Disqus gestellt haben. Diese Statistiken umfassen alle Anfragen, die von Einzelpersonen weltweit eingehen.
​

· ANFRAGEN ZUM WISSEN / ZUGRIFF AUF DATEN: Erhalten: 258 \| Erfüllt: 258 \| Abgelehnt: 0 \| Median Antwortzeit: 1,19 Tage \| Mittlere Antwortzeit: 1,78 Tage

· ANFRAGEN ZUR LÖSCHUNG: Erhalten: 1.175 \| Erfüllt: 1.175 \| Abgelehnt: 0 \| Median Antwortzeit: 0,83 Tage \| Mittlere Antwortzeit: 1,33 Tage

· ANFRAGEN ZUM ABMELDEN VON VERKAUF ODER TEILEN: Erhalten: 1.303 \| Erfüllt: 1.303 \| Abgelehnt: 0 \| Median Antwortzeit: 0,85 Tage \| Mittlere Antwortzeit: 1,34 Tage

· ANFRAGEN ZUR BEGRENZUNG DER NUTZUNG SENSIBLER PERSONENBEZOGENER DATEN: Erhalten: 280 \| Erfüllt: 280 \| Abgelehnt: 0 \| Median Antwortzeit: 1,00 Tag \| Mittlere Antwortzeit: 1,44 Tage

· ANTRÄGE NICHT BEARBEITET, WEIL DER ANTRAGSTELLER DIE VERIFIZIERUNG NICHT ABGESCHLOSSEN HAT: 595

· LÖSCHANFRAGEN GANZ ODER TEILWEISE NICHT ERFÜLLT: 0 — Disqus lehnte 2025 keine Löschanfragen ab.

#### 10. INTERNATIONALE DATENÜBERTRAGUNGEN

Für Nutzer mit Sitz im Europäischen Wirtschaftsraum (EWR) und im Vereinigten Königreich können wir Ihre persönlichen Daten innerhalb der Disqus-Gruppe oder mit externen Dritten teilen. Dies kann bedeuten, Ihre Daten außerhalb des EWR zu übertragen. Konkret werden die Daten von technischen Support-Teams in den Vereinigten Staaten, Indien und den Philippinen verarbeitet.
​

Wann immer Ihre personenbezogenen Daten außerhalb des EWR verarbeitet werden, schützen wir sie durch spezifische Verträge, die von der Europäischen Kommission und/oder den gleichwertigen vertraglichen Bestimmungen des britischen Informationskommissarbüros genehmigt wurden, die sicherstellen, dass personenbezogene Daten denselben Schutz genießen wie in Europa, unabhängig davon, wo sie verarbeitet werden.

#### 11. NICHT VERFOLGEN / GLOBALE DATENSCHUTZKONTROLLE ("GPC")

Disqus erkennt und verarbeitet Nutzerpräferenzen, wie sie in den browserbasierten "Do Not Track"-Signalen über die Global Privacy Control ("GPC")-Funktion gesetzt sind.

#### 12. ALLGEMEIN

**Kontakt**

Wenn Sie Fragen zu dieser Datenschutzrichtlinie haben, senden Sie uns bitte eine E-Mail an privacy@disqus.com oder kontaktieren Sie uns per Post unter 3 Park Avenue, 33rd Floor, New York, NY 10016.
​

**Änderungen der Datenschutzrichtlinie**

Disqus kann diese Datenschutzrichtlinie nach eigenem Ermessen von Zeit zu Zeit ändern oder aktualisieren, daher sollten Sie diese Seite regelmäßig lesen. Wenn wir die Richtlinie ändern, aktualisieren wir das "letzte geänderte" Datum oben auf dieser Seite. Ihre fortgesetzte Nutzung der Seite nach dem Veröffentlichen etwaiger Änderungen dieser Richtlinie bedeutet, dass Sie diese Änderungen akzeptieren.
​

Klick [Hier](#terms-of-service) um die Nutzungsbedingungen einzusehen.

Unsere Datenschutzrichtlinie ist auch in den folgenden Sprachen verfügbar:
[English](#disqus-privacy-policy)
[Español](#politica-de-privacidad-de-disqus)
[Français](#politique-de-confidentialite-de-disqus)
[Italiano](#disqus-informativa-sulla-riservatezza)
[Português](#politica-de-privacidade-do-disqus)

### General Security Tips {#general-security-tips}

**What Government Agencies Provide ID Theft Resources?**

-   *U.S. Federal Trade Commission (FTC):* The FTC has helpful information about how to avoid and protect against ID theft. Write to: Consumer Response Center, 600 Pennsylvania Ave., NW, H-130, Washington, D.C. 20580. Call Toll-Free: 1-877-IDTHEFT (438-4338); or Visit: 0

-   *State Attorney General Offices:* You may contact the Attorney General’s office in the state in which you reside for more information about preventing and managing ID theft.

For IOWA Residents: You may contact local law enforcement or the Iowa Attorney General’s Office at 1305 E. Walnut St., Des Moines, IA 50319; Tel: (515) 281-5164; or 0
​
For MARYLAND Residents: You may obtain information about preventing identity theft from the FTC or the Maryland Attorney General’s Office at 200 St. Paul Place, Baltimore, MD 21202; Tel: (888) 743-0023; or 0
​
For NEW MEXICO Residents: You have a right to place a security freeze on your credit report or submit a declaration of removal with a consumer reporting agency pursuant to the Fair Credit Reporting and Identity Security Act. Please see below for more information on security freezes.
​
For NORTH CAROLINA Residents: You may obtain information about preventing identity theft from the FTC or the North Carolina Attorney General’s Office at 9001 Mail Service Center, Raleigh, NC 27699-9001; Tel: (919) 716-6400; Fax: (919) 716-6750; or 0
​
For RHODE ISLAND Residents: You may obtain information about preventing identity theft from the FTC or the Rhode Island Attorney General’s Office at 150 South Main Street, Providence, RI 02903; Tel: (401) 274-4400; or 0
​
​
**How Do I Get A Free Credit Report?** You may obtain one (1) free copy of your credit report once every 12 months, and may purchase additional copies. Call Toll-Free: 1-877-322-8228; or Visit: 0; or contact: Equifax, P.O. Box 740241, Atlanta, GA 30374-0241 (800) 685-1111 (1); Experian P.O. Box 2002, Allen, TX 75013, (888) 397-3742 (2) TransUnion, P. O. Box 1000, Chester, PA 19022, (800) 888-4213 (3).
​
**What is a “Fraud Alert”?** You may have the right to place a fraud alert in your file to alert potential creditors that you may be a victim of identity theft. Creditors must then follow certain procedures to protect you. You should know that a fraud alert may delay your ability to obtain credit. An “initial fraud alert” stays in your file for at least 90 days. An “extended fraud alert” stays in your file for 7 years, and will require an identity theft report, which is usually a filed police report. You may place a fraud alert by calling any one of the three national consumer reporting agencies: Equifax: 1-800-525-6285; Experian: 1-888-397-3742; TransUnion: 1-800-680-7289
​
**What is a “Security Freeze”?** Certain U.S. state laws, including Massachusetts, allow a security freeze, which prevents approval for credit, loans or services in your name without your consent. A security freeze can interfere or delay your ability to obtain credit.

-   To place a freeze, send a request by mail to each consumer reporting agency (addresses below) with the following (for each individual): (1) Full name, middle initial and any suffixes; (2) Social Security Number; (3) Date of Birth; (4) proof of current address (such as a utility bill or telephone bill) and list of previous addresses for past five years; (5) copy of government issued ID card, and (6) copy of police report, investigative report or complaint to law enforcement regarding ID theft. You may be charged a fee up to \$5.00 to place, lift, and/or remove a freeze, unless you are a victim of ID theft or the spouse of a victim, and you have submitted a valid police report relating to the ID theft incident to the consumer reporting agency. The consumer reporting agencies have three business days after receiving your letter to place a security freeze on your credit report. The credit bureaus must also send written confirmation to you within five (5) business days and provide you a unique PIN or password that can be used by you to authorize the removal or lifting of the security freeze.

-   To lift the security freeze to allow a specific entity or individual access to your credit report, you must call or send a written request to the consumer reporting agencies by mail and include proper identification (name, address, and SSN) and the PIN number or password provided to you when you placed the security freeze as well as the identities of entities or individuals you would like to receive your credit report or the specific period of time you want the credit report available. The consumer reporting agencies have three business days after receiving your request to lift the security freeze for the identified entities or specified time period.

-   To remove the security freeze, you must send a written request to each of the three credit bureaus by mail and include proper identification (name, address, and SSN) and the PIN number or password provided to you when you placed the security freeze. The credit bureaus have three business days after receiving your request to remove the freeze. *Equifax Security Freeze*: P.O. Box 105788, Atlanta, Georgia 30348; *Experian Security Freeze*: P.O. Box 9554, Allen, TX 75013; *TransUnion (Fraud Victim Assistance Division)*: P.O. Box 6790, Fullerton, CA 92834-6790.

### How to Edit Your Data Sharing Settings {#how-to-edit-your-data-sharing-settings}

On Disqus, Data Sharing can be disabled or opted out of at any time, regardless of what plan you are on. If you opt out as a commenter, you will not be tracked by Disqus on any sites you comment on. If tracking is disabled at the site level, no users will be tracked on your site, though they may be tracked elsewhere if they have not opted out of Data Sharing at their account level.

Sharing settings can be adjusted on the Data Sharing Settings page. This page is also available via the Privacy link at the bottom of every commenting embed.

Alternatively, you can opt out by turning on your browser's "Do Not Track" (DNT) setting and Disqus will recognize these requests.
​
Note that Disqus does not recognize the DNT requests in Safari 7, but you can still opt out using the primary method on the Data Sharing Settings page

#### Publisher

Forums can disable data tracking for their forum via the **Tracking** setting on their forum's Settings > Advanced page.

We continue to strive to make Disqus the best way to participate in rich, relevant discussion experiences across the web, while ensuring trust, safety and control are always top of mind. For more information on Privacy see our Privacy Policy.

### How to Report Abuse {#how-to-report-abuse}

If you've become the target of abusive activity or are seeing it take place in a discussion or profile powered by Disqus, there are a few things you can do to raise awareness of the issue to the right people, and hopefully, find a resolution.

#### #1. Do not engage with the person being abusive

#### #2. Block the user (recommended)

#### #3. Tell the site moderators that abuse is happening

#### #4. Check the Basic Rules

*Ask yourself, is this user breaking the Basic Rules?*

-   *If* ***YES,* *report the user* *from their profile*

-   *If* ***NO, block the user and be merry.* ***(still recommended)**Note: Offensive content is tolerated and Disqus does not moderate individual comments, that is the responsibility of site moderators. Disqus takes action on the most extreme reports in which the* *Terms of Service* *have been violated. Disqus does not mediate content or intervene in disputes between users.*

#### What Happens After Reporting a User?

User Reporting is incorporated into reputation scores and helps site moderators make better decisions about who to remove from their community. Moderation decisions are made by site moderators, not Disqus.

#### Unwanted followers

If you've picked up an unwanted follower that's bothering you across multiple sites, you can make your profile activity private to prevent all other users from being able to follow you or see your comments in your profile; this can help isolate issues with certain people to certain communities. If they already follow you, you can remove them and/or block them.

#### Cyber Harassment and Illegal Activity

If you or someone you care about is experiencing cyber harassment online, please seek out additional resources such as ReachOut as well as reporting to site moderators and Disqus.
​
If you encounter a direct threat of suicide or self-harm on Disqus, visit How to report threats of suicide or self-harm.

If you feel a crime has been committed or a credible threat has been made against you or someone else via Disqus, please contact your local law enforcement in addition to reporting the user to Disqus.

#### Related Documents:

-   User Blocking

-   Basic Rules of Disqus

-   Flagging Users

-   Flagging Comments

-   How to Contact a Site Moderator

### How to report threats of suicide or self-harm {#how-to-report-threats-of-suicide-or-self-harm}

If you encounter a direct threat of suicide or self-harm on Disqus, please contact local law enforcement as soon as possible.
​
You can submit reports of suicidal content to Disqus by flagging the user for "**Threat - posted directly threatening content**" by following: 0
​
After we have evaluated a report of self-harm or suicide, Disqus will contact the reported user privately via email to provide them with online resources where they can seek help.
​
You can also make use of the following online and prevention resources:
​
Crisis Text Line 0
Live Chats: 0 (2pm-2am ET) or 1
United States: National Suicide Prevention Lifeline at 1.800.273.TALK (8255) or 0
Worldwide: Befrienders: 0

### Is someone else posting using my account? {#is-someone-else-posting-using-my-account}

#### This can happen when...

-   You have a fairly generic email address, e.g., bruce@wayne.com

-   Someone posted a guest comment using a fake email address without realizing the email address they used (yours) is real.

Another person has a similar email address to yours and misspelled theirs when posting a guest comment or creating an account; or

Someone posted a guest comment using a fake email address without realizing the email address they used (yours) is real.

#### What do I do if someone made a full Disqus account with my email address?

-   Enter the email address into the box at our password recovery page. You should get an email within minutes at the address you entered into the box. If you don't, check your spam filter and make sure you're not filtering emails from Disqus in any way.

-   Using the link in the email, choose a new password. This prevents the person from using your email to continue to make posts.

-   Delete any comments which were made with the account before you took control of it. Keep in mind that this will remove identifiable information from the comment (no name, email, or avatar) but will not remove the comment from the discussion thread. Visit Removing and Editing Your Comments for more information.

**Note:** We recommend not deleting the account. You're not required to use the account but keeping it registered will prevent others from re-registering an account with your email address (whether intentional or not).

-   Once you've taken control of the account, you can delete each comment manually, which will prevent it from appearing on the page at all, or you can delete the account, which will anonymize the comments, stripping them of all user information. See Removing and Editing your Comments.

#### How is someone else able to post with my account?

Guest users can post under any name, as display names do not have to be unique. More information on the difference between usernames and display names can be found here.

#### How can I block others from commenting using my email address?

It's not currently possible to block guest comments posted using your email address. However, Guest comments are separate from accounts registered with the same email, so these comments shouldn't affect your Disqus account.
​
If your email address is being used for a Disqus account that you did not register, we recommend taking control of the account through a password reset email and leaving an account in your control so that the email address cannot be re-registered.

### Parody Accounts {#parody-accounts}

Although impersonating or portraying another person in a confusing or deceptive manner is against our terms of service, parody accounts are okay - as long as they meet the requirements below.

#### Avatar

The avatar must not be exactly the same as the account it is parodying. This includes the default avatar.

#### Display Name

The display name must not be exactly the same as the account it is parodying without an additional distinguishing word such as "not" or "fake". If the display name is different only in subtle changes of spelling, it must also be accompanied by a distinguishing word.

#### Bio

The bio must include a statement that distinguishes it from the account it is parodying. For instance, a bio could say, "this a parody account" or "not affiliated with..."

#### How do I report impersonation?

Click the flag icon in the profile of the user who is impersonating your account and complete a short survey. Please be sure to include the email address associated with the account being impersonated for verification purposes.

#### What happens when an impersonation complaint is reported?

The case will be reviewed and if the account doesn't comply with the above requirements (or is otherwise found to be portraying another person in a confusing or deceptive manner) the profile of the account will be set back to default and the user will have an opportunity to update their profile information abiding by the above requirements. Repeated violations may result in an account getting banned or deactivated.

#### Have questions?

Drop us a line at our Support Form.

### Politique de Confidentialité de Disqus {#politique-de-confidentialite-de-disqus}

Politique de Confidentialité de Disqus

**Mis à jour** le 10 juillet 2026

Cette politique de confidentialité vous indique comment Disqus collecte, utilise, vend, divulgue et protège les données vous concernant (l'« Utilisateur ») en lien avec notre service (tel que défini ci-dessous), ainsi que vos choix concernant notre collecte et notre utilisation de ces données.

#### 1. INTRODUCTION

**Aperçu**

Disqus propose une plateforme en ligne de partage public de commentaires et d'opinions où les utilisateurs se connectent et créent des profils pour participer à des conversations avec leurs pairs et profiter d'une expérience interactive dans les sections de commentaires, sondages et autres fonctionnalités interactives proposées sur ce site, ainsi que intégrées sur des sites tiers. L'utilisation de notre plateforme et de nos logiciels, ainsi que l'interaction avec nos cookies ou technologies de suivi similaires (collectivement le « Service »), que ce soit sur ce site ou sur un site tiers, sont soumises aux termes de cette Politique de confidentialité. Le Service est une plateforme publique et Disqus ou d'autres peuvent rechercher, voir, utiliser ou republier tout contenu utilisateur (tel que défini dans nos Conditions d'utilisation) que vous publiez via le Service. Disqus est également une société de marketing et de données, et utilise et partage des données personnelles collectées sur des sites tiers où notre service est activé à des fins marketing, y compris la publicité comportementale intercontextuelle. Pour plus d'informations sur nos activités marketing, veuillez consulter la section 4 : Publicité ciblée et partenaires publicitaires ci-dessous.
​

**Applicabilité aux sites web et services tiers**

Disqus propose un service d'engagement en ligne que d'autres sites web utilisent pour permettre la discussion et l'interactivité entre leurs utilisateurs. Cette politique de confidentialité s'applique aux données que Disqus collecte sur les utilisateurs du Service et via les cookies sur les sites web compatibles avec le Service, et ne s'applique pas aux pratiques de collecte indépendante de données de tout site web utilisant le Service ou tout autre site web lié au Service. Pour des informations sur la manière dont les sites tiers collectent et utilisent vos informations personnelles, veuillez consulter les politiques de confidentialité de ces sites.
​

**Vos droits à la vie privée**

Vous avez des droits sur vos informations personnelles. Ces droits sont décrits plus en détail dans la Section 9 : Vos droits, ci-dessous. Vous pouvez exercer vos droits à la confidentialité des données à l'adresse suivante Tiens.
​

**Vie privée des enfants**

Le service n'est pas destiné aux enfants de moins de 18 ans. Nous ne collectons ni ne vendons sciemment des informations personnelles auprès d'enfants de moins de 18 ans, ni ne permettons sciemment à ces personnes de s'enregistrer sur ce service. Si nous apprendrons que nous avons collecté des informations personnelles auprès d'un enfant de moins de 18 ans, nous les supprimerons. Si vous pensez que nous avons recueilli des informations personnelles auprès d'un enfant de moins de 18 ans, veuillez nous contacter ou déposer une demande de droits à la protection des données à l'adresse suivante Tiens.
​

**Notre utilisation de l'intelligence artificielle**

Disqus utilise l'intelligence artificielle et l'apprentissage automatique de deux manières. Tout d'abord, l'IA est utilisée pour modérer le contenu sur la plateforme — détectant le spam et les contenus qui enfreignent nos règles communautaires afin que le service puisse être sécurisé et fonctionnel pour les utilisateurs. Deuxièmement, l'IA est utilisée pour vous fournir des publicités plus pertinentes en fonction de vos centres d'intérêt et de votre activité en ligne. Dans les deux cas, Disqus n'utilise pas l'IA pour prendre des décisions ayant des effets juridiques ou d'une signification similaire sur les individus. Disqus n'utilise ni ne vend de données personnelles des consommateurs dans le but d'entraîner de grands modèles de langage. Nos systèmes d'IA fonctionnent conformément à nos engagements en matière de confidentialité tels que décrits dans cette politique.

#### 2. LES DONNÉES QUE NOUS COLLECTONS À VOTRE SUJET

Les données personnelles, ou informations personnelles, désignent toute information concernant une personne pouvant raisonnablement être liée, directement ou indirectement, à une personne ou un ménage spécifique.

Nous collectons, et avons collecté au cours des douze (12) derniers mois, les types de données personnelles suivantes concernant les utilisateurs:

1\. **Identifiants (« Données d'identité »)** tels que prénom, nom de famille, nom d'utilisateur ou identifiant similaire, adresse IP (protocole internet), identifiant unique de cookie, identifiant d'appareil, date de naissance, adresse e-mail, numéro de téléphone et adresse postale ;

2\. **Les catégories d'informations personnelles listées dans le Statut californien sur les dossiers clients (Code civil de Californie § 1798.80(e)),** qui incluent toute information identifiant, se rapportant, décrivant, pouvant être associée à un individu particulier mais qui n'est pas accessible au public à partir des documents fédéraux, étatiques ou locaux ;

3\. **Des caractéristiques de classification protégées** par la loi californienne ou fédérale, telles que la race, le sexe ou l'âge ;

4\. **Informations sur l'activité Internet ou d'autres réseaux électroniques,** telles que l'historique de navigation et de commentaires, les retours et réponses aux sondages, vos données de connexion, type et version du navigateur, réglage et emplacement du fuseau horaire, types et versions des plug-ins du navigateur, système d'exploitation et plateforme ainsi que d'autres technologies sur les appareils utilisés pour accéder au Service ;

5\. **Informations professionnelles ou liées à l'emploi,** dans la mesure où vous les incluez dans votre profil ou vos commentaires, ou qu'elles peuvent être déduites des pages que vous consultez ;

6\. **Les informations éducatives, dans** la mesure où vous les incluez dans votre profil ou vos commentaires, ou qu'elles peuvent être déduites des pages que vous consultez ;

7\. **Informations personnelles sensibles.** Nous ne collectons intentionnellement aucune donnée personnelle sur votre race ou ethnie, vos croyances religieuses ou philosophiques, votre vie sexuelle, votre orientation sexuelle, vos opinions politiques, votre appartenance syndicale, vos informations sur votre santé ou données génétiques ou biométriques, ni sur des condamnations et infractions pénales. Cependant, si vous faites des commentaires sur le Service incluant de telles informations à votre sujet, elles seront publiques et pourront être traitées par Disqus ou d'autres. De plus, nous collectons et partageons des informations recueillies via des cookies ou une technologie de suivi similaire concernant les pages web que vous avez consultées, ce qui peut permettre aux tiers avec lesquels nous partageons vos informations de tirer des conclusions à votre sujet pouvant constituer des informations personnelles sensibles.
​

Nous pouvons également combiner, désidentifier ou agréger toutes les informations que nous collectons via notre Service pour l'une des finalités décrites ci-dessous.

#### 3. COMMENT VOS DONNÉES PERSONNELLES SONT-ELLES COLLECTÉES?

Nous utilisons différentes méthodes pour collecter des données de vous et à votre sujet, notamment via:

**Interactions directes**

Cela inclut les données personnelles que vous fournissez lorsque vous créez un compte ou laissez un commentaire.
​

**Technologies ou interactions automatisées**

Lorsque vous interagissez avec notre Service, nous pouvons automatiquement collecter des données techniques sur votre équipement, vos actions et schémas de navigation. Plus précisément, de la manière suivante:
​

**Cookies**

Un cookie est un petit fichier numérique placé sur le disque dur de votre ordinateur. Vous pouvez refuser d'accepter les cookies du navigateur en activant les paramètres de votre navigateur ou, dans certains cas, en interagissant avec des bannières de cookies contextuels. Nous utilisons des cookies placés sur des sites tiers où le Service est activé pour recueillir des informations sur la façon dont vous interagissez avec ces sites, même si vous ne laissez pas de commentaires, ne répondez pas à des sondages ou n'interagissez pas directement avec le Service sur ces sites, ainsi que des informations sur les autres sites que vous visitez. Disqus utilise des cookies et permet aux partenaires de les configurer via le Service afin de faciliter la publicité comportementale intercontextuelle. Cela signifie en pratique que nous utilisons des cookies pour aider à déterminer quelles publicités vous voyez en ligne, en enregistrant vos visites sur de nombreux sites d'annonceurs / marques, puis en vous montrant des publicités pour des produits et services similaires.
​

Disqus utilise des cookies d'« authentification », par exemple sessionid, disqusauth et disqusauths, pour vous maintenir connecté depuis votre navigateur web et personnaliser votre expérience Disqus.
​

Disqus utilise des cookies « uniques », par exemple disqus_unique et \_jid, pour associer des activités web à une charge de page et à un navigateur web, et pour comprendre vos centres d'intérêt et votre utilisation du produit.

Lorsque Disqus charge des publicités, nous utilisons des technologies de diffusion de publicités de Google qui peuvent créer des cookies à des fins de marketing personnalisé, associant les publicités à des activités ultérieures, et limitant la fréquence à laquelle vous affichez des publicités spécifiques.
​

**Informations sur le fichier journal**

Les journaux serveur recueillent des données techniques telles que votre adresse IP, le type de navigateur, ainsi que des informations sur le nombre de clics et la manière dont vous interagissez avec les liens du Service, les sites partenaires, les noms de domaine, les pages d'atterrissage, les pages consultées et d'autres informations similaires.
​

**Pixels et trackers similaires**

Lorsque vous utilisez le Service, nous utilisons des pixels, des gifs clairs (également appelés balises web) qui servent à collecter des données techniques et des informations telles que les schémas d'utilisation en ligne. Nous utilisons également des gifs clairs dans les emails HTML envoyés à nos utilisateurs pour suivre quels e-mails sont ouverts et quels liens ou publicités sont cliqués par les destinataires. Nous utilisons également des pixels et des gifs clairs sur des sites tiers où le Service est activé pour collecter des informations sur votre interaction avec ces sites, même si vous ne laissez pas de commentaires, ne répondez pas aux sondages ou n'interagissez pas directement avec le Service sur ces sites.
​

**Tiers ou sources publiques**

Nous obtenons ou recevons des données personnelles à votre sujet auprès de fournisseurs d'analyses tels que Google ; Partenaires publicitaires [Tiens](#cookies-and-data-recipients); et des courtiers tiers en données qui vendent des données personnelles. Nous obtenons ou recevons des données personnelles via des connexions ou des connexions tierces via des plateformes de réseaux sociaux telles que Facebook Connect, Google ou Twitter/X lorsque vous « suivez », « aimez » ou liez votre compte au Service. Veuillez noter que certains de ces prestataires, en particulier des fournisseurs d'analytique comme Google, peuvent traiter des données provenant de résidents de l'Espace économique européen (EEE) en dehors de l'EEE.
​

**Gérer les cookies dans votre navigateur**

Vous pouvez ajuster les paramètres de votre navigateur pour gérer vos préférences de cookies, par exemple pour qu'il vous notifie lorsque vous recevez un cookie et vous donne le choix de l'accepter ou non. Si vous rejetez les cookies, vous pouvez toujours utiliser notre site, mais la fonctionnalité de certaines zones peut être limitée.
​

Voici des liens vers des informations sur la gestion de vos préférences de cookies dans les navigateurs courants :

· Cookies Google Chrome : Google Chrome Cookies

· Cookies Mozilla Firefox : Mozilla Firefox Cookies

· Cookies Internet Explorer : Internet Explorer Cookies

· Cookies Safari : Safari Cookies

· Google Analytics : Google Analytics

#### 4. PARTENAIRES DE PUBLICITÉ CIBLÉE ET PUBLICITAIRES

**Publicité ciblée**

La publicité est la principale façon dont Disqus gagne de l'argent. Les revenus publicitaires permettent à Disqus d'exploiter, de soutenir et d'améliorer le Service. Disqus utilise et partage avec des partenaires publicitaires tiers et affiliés, des identifiants de cookies, des identifiants d'appareils (y compris mobiles), des adresses e-mail hachées, des adresses IP, des informations sur les fournisseurs d'accès Internet (FAI) et navigateurs, des données démographiques ou d'intérêt, le contenu consulté et les actions entreprises sur le Service, sur les sites partenaires ou sur d'autres sites tiers. Cela inclut des informations sur les sites web que vous avez consultés et les publicités avec lesquelles vous avez interagi afin de vous fournir des publicités plus pertinentes adaptées à vos préférences et centres d'intérêt. Cela peut provenir de votre interaction avec le Service, les sites partenaires ou d'autres sites tiers. Pour une liste des partenaires publicitaires tiers avec lesquels Disqus travaille actuellement, voir [Tiens](#cookies-and-data-recipients).
​

**Marketing par e-mail**

Disqus peut également vous envoyer des newsletters et des messages marketing par e-mail si vous nous avez donné la permission ou consenti à recevoir ces e-mails, comme requis dans la juridiction où vous résidez. Les messages de marketing par email peuvent être adaptés à vos centres d'intérêt en se basant sur les informations décrites ci-dessus dans cette section. Pour plus d'informations sur la manière de se désinscrire et d'exercer vos droits à la vie privée, veuillez consulter [Tiens](#updating-your-account-settings).
​

**Partenaires publicitaires et divulgations tierces**

Nous partageons et partageons des données avec des tiers qui collectent des informations sur divers canaux, y compris en ligne et hors ligne, dans le but de vous proposer des publicités plus pertinentes à vous ou à votre entreprise. Nos partenaires utilisent ces informations pour vous reconnaître à travers différents canaux et plateformes, au fil du temps (y compris, mais sans s'y limiter, ordinateurs, appareils mobiles, télévision adressable ou autres médias), à des fins de marketing, d'analyse, d'attribution et de reporting. Bien que Disqus ne fasse pas d'inférences à partir de vos données, nos partenaires publicitaires peuvent tirer des conclusions de vos données afin de comprendre vos préférences, caractéristiques, tendances psychologiques, prédisposances, comportements, attitudes, intelligence, capacités et aptitudes.

Au cours des douze (12) derniers mois, nous avons partagé ou vendu les données suivantes :

· Identifiants ;

· Catégories d'informations personnelles listées dans le Statut californien sur les dossiers clients (Code civil de Californie § 1798.80(e) ;

· Caractéristiques de classification protégées ;

· des informations sur Internet ou d'autres réseaux électroniques ;

· Informations professionnelles ou liées à l'emploi ;

· Informations éducatives ; et,

· Informations personnelles sensibles.

#### 5. COMMENT NOUS UTILISONS LES DONNÉES PERSONNELLES ET NOTRE BASE LÉGALE POUR L'UTILISATION

Dans l'EEE, le Royaume-Uni (UK) et le Brésil, nous comptons généralement sur votre consentement pour utiliser des données personnelles. Dans certains cas, nous utilisons les données selon les besoins de nos intérêts légitimes (ou ceux d'un tiers), mais uniquement lorsque vos droits et intérêts ne sont pas affectés négativement. Enfin, nous pouvons également l'utiliser lorsque nous devons respecter une obligation légale ou réglementaire, ou pour protéger la santé, la sécurité ou les droits légaux de toute personne.
​

Nous n'utilisons pas d'informations personnelles pour soutenir uniquement des décisions automatisées qui ont des effets juridiques ou similaires importants à votre sujet, appelés « profilage » selon certaines lois.
​

Nous avons présenté ci-dessous une description des façons dont nous utilisons les données personnelles, ainsi que des bases juridiques sur lesquelles nous nous appuyons pour le faire. Nous avons également identifié quels sont nos intérêts légitimes lorsque cela est approprié.
​

Notez que nous pouvons traiter vos données personnelles pour plusieurs motifs juridiques selon l'objectif spécifique pour lequel nous utilisons vos données. Veuillez nous contacter si vous avez d'autres questions.

-   **But ou activité**
    -   Type de données personnelles
    -   Source des données personnelles
    -   Base du traitement
-   **Pour vous enregistrer comme nouvel utilisateur**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Activité sur Internet ou réseau électronique
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Exécution d'un contrat avec vous · Avec votre consentement
-   **Pour gérer notre relation avec vous, ce qui peut inclure vous informer des modifications de nos conditions d'utilisation ou de notre politique de confidentialité, vous demander de laisser un avis ou de répondre à un sondage**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Exécution d'un contrat avec vous · Nécessaire pour respecter une obligation légale · Nécessaire pour nos intérêts légitimes (pour tenir nos dossiers à jour et étudier comment les clients utilisent notre service) · Avec votre consentement
-   **Gérer et protéger notre activité et le Service (y compris le dépannage, l'analyse des données, les tests, le support, le reporting et l'hébergement des données)**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Nécessaire pour respecter une obligation légale · Nécessaire pour nos intérêts légitimes (par exemple, pour tenir nos dossiers à jour et étudier comment les clients utilisent notre service)
-   **Pour vous proposer du contenu et des publicités pertinents et mesurer ou comprendre l'efficacité de la publicité que nous vous proposons**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique · Informations professionnelles ou liées à l'emploi · Informations éducatives · \[Informations sensibles\]
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Consentement. Lorsque vous avez donné un consentement explicite à notre utilisation de données personnelles afin de vous fournir un contenu pertinent et personnalisé · Nécessaire pour nos intérêts légitimes
-   **Utiliser l'analyse de données pour améliorer notre service**
    -   · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Nécessaire pour nos intérêts légitimes (pour maintenir notre site web à jour et pertinent, pour développer notre entreprise et pour orienter notre stratégie marketing)
-   **Pour vous envoyer des newsletters et des emails promotionnels qui pourraient vous intéresser**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique · Informations professionnelles ou liées à l'emploi · Informations éducatives · \[Informations sensibles\]
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Consentement. Lorsque vous avez donné un consentement explicite à notre utilisation de données personnelles afin de vous fournir un contenu pertinent et personnalisé · Nécessaire pour nos intérêts légitimes (développer notre service et développer notre activité)
-   **Vendre ou partager des données personnelles avec des tiers à des fins marketing et publicitaires**
    -   · Identifiants · Catégories d'informations personnelles listées dans la loi californienne sur les dossiers clients · Caractéristiques de classification protégées selon la loi californienne ou fédérale · Activité sur Internet ou réseau électronique · Informations professionnelles ou liées à l'emploi · Informations éducatives · \[Informations sensibles\]
    -   · Interactions directes · Technologies automatisées · Cookies · Informations sur le fichier journal · Gifs clairs · Tiers ou sources publiques
    -   · Consentement. Lorsque vous avez donné votre consentement explicite à notre utilisation de données personnelles afin de vous fournir un contenu pertinent et personnalisé. · Intérêt légitime. En particulier en dehors de l'EEE, nous pouvons vendre ou partager des données selon les lois locales applicables, ce qui peut ne pas nécessiter votre consentement pour que nous le permettez.

#### 6. DIVULGATION DE VOS DONNÉES PERSONNELLES

Nous pouvons divulguer n'importe quelle catégorie de données personnelles que nous collectons auprès de vous et à votre sujet dans les circonstances suivantes, pour les objectifs évoqués dans le tableau ci-dessus. Au cours des 12 derniers mois, nous avons vendu et/ou partagé toutes les catégories de données personnelles mentionnées ci-dessus avec Zeta Global et nos partenaires commerciaux (comme décrit ci-dessous).
​

**Zeta Global**

Les données utilisateurs de Disqus sont partagées avec notre société affiliée Zeta Global à des fins marketing, y compris la publicité comportementale intercontexte. Les données Disqus sont fusionnées par Zeta avec celles provenant d'autres sources en ligne et hors ligne pour orienter les décisions marketing et les messages au nom des clients annonceurs de Zeta (marques et agences). Zeta peut également partager les données obtenues auprès de Disqus avec ces clients annonceurs.
​

**Partenaires d'affaires**

Nous partageons également des données personnelles avec des tiers externes à des fins similaires à celles de Zeta Global. Ces tiers incluent d'autres entreprises participant à la publicité comportementale intercontextuelle, ainsi que des courtiers en données qui peuvent vendre ou partager vos données personnelles. Nous pouvons également partager des données personnelles avec des annonceurs avec lesquels vous avez déjà partagé vos catégories d'intérêt en ligne en fonction de votre activité en ligne, à des fins marketing. Pour une liste de tiers externes, veuillez consulter [Tiens](#cookies-and-data-recipients).
​

**Fournisseurs de services**

Nous comptons sur les prestataires de services (par exemple, les entreprises qui proposent des services d'hébergement web ou d'analytique) pour le fonctionnement de notre site web et de notre Service, et ces prestataires peuvent avoir accès à vos données personnelles. Nos prestataires de services sont contractuellement interdits d'utiliser vos données personnelles à leurs propres fins et sont tenus de traiter vos données personnelles et de préserver leur confidentialité de la même manière que nous.
​

**Transactions commerciales**

Nous pourrions chercher à acquérir d'autres entreprises ou à fusionner avec elles. Si un changement concerne notre activité, la nouvelle société pourra utiliser vos données personnelles de la même manière qu'indiqué dans cet avis de confidentialité. Nous pouvons également nous engager dans un partage limité des données pour évaluer et tester nos partenaires potentiels en matière de données. Lorsque ces tests ont lieu, les données sont chiffrées ou encodées avant d'être partagées. De plus, ce partage est soumis à des dispositions contractuelles appropriées et les données sont rapidement supprimées après la fin des tests.
​

**Exigences légales**

De plus, nous pouvons accéder, préserver et divulguer vos données personnelles, si nous estimons que cela est exigé par la loi, une ordonnance judiciaire ou d'autres procédures juridiques valides. Nous pouvons également accéder, préserver et divulguer ces données personnelles si nous croyons de bonne foi que la divulgation est nécessaire pour protéger vos droits, nos droits ou ceux d'autrui, ou pour enquêter sur la fraude.
​

Nous exigeons que tous les tiers avec lesquels nous partageons des données prennent les mesures appropriées pour assurer la sécurité de vos données personnelles et les traiter conformément à toutes les lois applicables.

#### 7. CONSERVATION DES DONNÉES

Nous ne conserverons vos données personnelles que le temps nécessaire pour remplir les objectifs que nous avons collectés, y compris pour satisfaire à des exigences légales, comptables ou de rapports.
​

Pour déterminer la période de conservation appropriée des données personnelles, nous prenons en compte la quantité, la nature et la sensibilité des données personnelles, le risque potentiel de préjudice lié à une utilisation ou divulgation non autorisée de vos données personnelles, les finalités pour lesquelles nous traitons vos données personnelles et la possibilité d'atteindre ces objectifs par d'autres moyens, ainsi que les exigences légales applicables.

#### 8. SÉCURITÉ

Disqus utilise des garanties commercialement raisonnables pour préserver l'intégrité et la sécurité de toutes les informations collectées par le Service. Pour protéger votre vie privée et votre sécurité, nous prenons des mesures raisonnables (comme demander un mot de passe unique) pour vérifier votre identité avant de vous accorder l'accès à votre compte. Vous êtes responsable de maintenir le secret de vos informations uniques de mot de passe et de votre compte, ainsi que de contrôler l'accès à vos communications par email depuis Disqus. Disqus n'est pas responsable des fonctionnalités ou des mesures de sécurité de tout tiers.

#### 9. VOS DROITS

Selon les lois de votre région, vous pouvez avoir un ou plusieurs des droits suivants en vertu des lois locales. Disqus étend ces droits à toutes les personnes, quel que soit leur lieu de résidence, y compris le droit de :

1\. **Demandez une copie** des données personnelles que nous avons recueillies à votre sujet ;

2\. **Désinscription de notre vente ou partage** de données personnelles à votre sujet ;

3\. **Refus de recevoir des e-mails** du Service ;

4\. **Nous demandons que nous supprimions** les données que nous avons collectées à votre sujet ;

5\. **Demandons que nous corrigions les** données incorrectes ;

6\. **Demandez que nous limitions notre utilisation ou que nous supprimions des données personnelles sensibles** que vous auriez pu fournir précédemment dans un commentaire (en raison de limitations techniques, toutes ces demandes seront traitées comme des demandes de suppression des données).

Disqus ne vous discriminera pas pour l'exercice de vos droits à la vie privée. Avant de répondre à une demande de copie de vos informations personnelles, Disqus est tenu de vérifier raisonnablement votre identité, ce que nous faisons généralement en envoyant un lien de vérification à l'adresse e-mail associée à vos informations personnelles. Selon les lois applicables, vous pouvez faire appel à un agent autorisé pour faire une demande en votre nom, mais cet agent doit pouvoir compléter notre processus de vérification afin de démontrer qu'il a été autorisé à faire la demande.
​

Vous avez également le droit, dans de nombreux pays, de contacter l'autorité compétente en matière de protection de la vie privée ou des données si vous estimez que nous ne respectons pas les lois sur la vie privée et que nous n'avons pas pu résoudre la situation à votre satisfaction. Si vous êtes résident de l'Espace économique européen, vous pouvez consulter l'autorité locale de protection des données Tiens. Les résidents du Royaume-Uni peuvent contacter le Bureau du Commissaire à l'Information Tiens.
​

Visitez notre page Choix de confidentialité Page Choix de confidentialité Pour plus de détails sur ces droits ou pour les exercer.
​

**Demandes de droits**

Selon la loi californienne, nous sommes tenus de publier des statistiques sur le nombre de personnes ayant exercé leur droit à la vie privée au cours de l'année précédente. Les statistiques suivantes concernent les personnes ayant fait des demandes concernant Disqus durant les douze mois se terminant le 31 décembre 2025. Ces statistiques couvrent toutes les demandes reçues de particuliers à travers le monde.
​

· DEMANDES DE CONNAISSANCE / ACCÈS AUX DONNÉES : Reçues : 258 \| Réalisé : 258 \| Refusé : 0 \| Temps de réponse médian : 1,19 jour \| Temps de réponse moyen : 1,78 jour

· DEMANDES DE SUPPRESSION : Reçues : 1 175 \| Obtenu avec : 1 175 \| Refusé : 0 \| Temps de réponse médian : 0,83 jour \| Temps de réponse moyen : 1,33 jour

· DEMANDES DE DÉSINSCRIPTION DE VENTE OU DE PARTAGE : Reçues : 1 303 \| Réalisé : 1 303 \| Refusé : 0 \| Temps de réponse médian : 0,85 jour \| Temps moyen de réponse : 1,34 jour

· DEMANDES DE LIMITATION DE L'UTILISATION D'INFORMATIONS PERSONNELLES SENSIBLES : Reçu : 280 \| Réalisé : 280 \| Refusé : 0 \| Temps de réponse médian : 1,00 jour \| Temps de réponse moyen : 1,44 jour

· DEMANDES NON TRAITÉES CAR LE DEMANDEUR N'A PAS COMPLÉTÉ LA VÉRIFICATION : 595

· DEMANDES DE SUPPRESSION NON SATISFAITES, EN TOUT OU EN PARTIE : 0 — Disqus n'a refusé aucune demande de suppression en 2025.

#### 10. TRANSFERTS INTERNATIONAUX DE DONNÉES

Pour les utilisateurs basés dans l'Espace économique européen (EEE) et au Royaume-Uni, nous pouvons partager vos données personnelles au sein du groupe Disqus ou avec des tiers externes. Cela peut impliquer le transfert de vos données hors de l'EEE. Plus précisément, les données seront traitées par des équipes de support technique aux États-Unis, en Inde et aux Philippines.

Chaque fois que vos données personnelles sont traitées en dehors de l'EEE, nous les protégeons à l'aide de contrats spécifiques approuvés par la Commission européenne et/ou des dispositions contractuelles équivalentes approuvées par le Bureau du Commissaire à l'Information du Royaume-Uni, qui garantissent que les données personnelles bénéficient de la même protection qu'en Europe, quel que soit leur lieu de traitement.

#### 11. NE PAS SUIVRE / CONTRÔLE GLOBAL DE LA CONFIDENTIALITÉ (« GPC »)

Disqus reconnaît et traite les préférences des utilisateurs telles que définies dans les signaux « Ne pas suivre » basés sur le navigateur via la fonction de Contrôle Global de la Confidentialité (« GPC »).

#### 12. GÉNÉRAL

**Contact**

Si vous avez des questions concernant cette politique de confidentialité, veuillez nous envoyer un e-mail à privacy@disqus.com, ou nous contacter par courrier au 3 Park Avenue, 33e étage, New York, NY 10016.
​

**Modifications de la politique de confidentialité**

Disqus peut, à sa seule discrétion, modifier ou mettre à jour cette politique de confidentialité de temps à autre, vous devriez donc consulter cette page périodiquement. Lorsque nous modifierons la politique, nous mettrons à jour la date de « dernière modification » en haut de cette page. Le fait que vous continuiez à utiliser le Site après la publication de toute modification de cette politique signifie que vous acceptez ces modifications.
​

Click [Tiens](#terms-of-service) pour consulter les Conditions d'utilisation.

Notre politique de confidentialité est également disponible dans les langues suivantes:
[Deutsch](#disqus-datenschutzrichtlinie)

[English](#disqus-privacy-policy)
[Español](#politica-de-privacidad-de-disqus)
[Italiano](#disqus-informativa-sulla-riservatezza)
[Português](#politica-de-privacidade-do-disqus)

### Política de Privacidad de Disqus {#politica-de-privacidad-de-disqus}

Política de privacidad de Disqus

**Actualizado** el 10 de julio de 2026

Esta Política de Privacidad te explica cómo Disqus recopila, utiliza, vende, divulga y protege los datos relacionados contigo (el "Usuario") en relación con nuestro Servicio (según se define a continuación), así como tus opciones respecto a nuestra recopilación y uso de estos datos.

#### 1. INTRODUCCIÓN

**Resumen**

Disqus ofrece una plataforma pública online de comentarios y opiniones donde los usuarios inician sesión y crean perfiles para participar en conversaciones con sus compañeros y disfrutar de una experiencia interactiva en las secciones de comentarios, encuestas y otras funciones interactivas de Disqus que se ofrecen en este sitio, así como incrustadas en sitios de terceros. El uso de nuestra plataforma y software, así como la interacción con nuestras cookies o tecnologías de seguimiento similares (colectivamente el "Servicio"), ya sea en este sitio o en un sitio de terceros, está sujeto a los términos de esta Política de Privacidad. El Servicio es una plataforma pública y Disqus u otros pueden buscar, ver, usar o volver a publicar cualquiera de tus Contenidos de Usuario (según lo definido en nuestros Términos de Uso) que publiques a través del Servicio. Disqus también es una empresa de marketing y datos, y utiliza y comparte datos personales recogidos de sitios de terceros donde nuestro Servicio está habilitado con fines de marketing, incluyendo publicidad conductual cruzada en contexto. Para más información sobre nuestras actividades de marketing, consulte la Sección 4: Publicidad Segmentada y Socios Publicitarios más abajo.
​

**Aplicabilidad a sitios web y servicios de terceros**

Disqus ofrece un servicio de interacción online que otros sitios web utilizan para facilitar la discusión e interacción entre sus usuarios. Esta Política de Privacidad se aplica a los datos que Disqus recopila sobre los usuarios del Servicio y a través de cookies en sitios web habilitados por el Servicio, y no se aplica a las prácticas independientes de recogida de datos de ningún sitio web que utilice el Servicio u otro sitio web vinculado desde el Servicio. Para obtener información sobre cómo los sitios web de terceros recopilan y utilizan tu información personal, consulta las políticas de privacidad de dichos sitios.
​

**Tus derechos de privacidad**

Tienes derechos sobre tu información personal. Estos derechos se describen con mayor detalle en la Sección 9: Tus Derechos, más abajo. Puedes ejercer tus derechos de privacidad de datos en Aquí.
​

**Privacidad infantil**

El servicio no está destinado a ser utilizado por menores de 18 años. No recopilamos ni vendemos a sabiendas información personal de menores de 18 años ni permitimos que dichas personas se registren en una cuenta en el servicio. En caso de que descubramos que hemos recopilado información personal de un menor de 18 años, la eliminaremos. Si crees que podríamos haber recopilado información personal de un menor de 18 años, por favor contáctanos o presenta una solicitud de derechos de privacidad de datos en Aquí.
​

**Nuestro uso de la inteligencia artificial**

Disqus utiliza la inteligencia artificial y el aprendizaje automático de dos maneras. En primer lugar, la IA se utiliza para ayudar a moderar el contenido en la plataforma, detectando spam y contenido que viola nuestras normas comunitarias para que el Servicio pueda mantenerse seguro y funcional para los usuarios. En segundo lugar, la IA se utiliza para ayudarte a ofrecer publicidad más relevante basada en tus intereses y actividad online. En ambos casos, Disqus no utiliza IA para tomar decisiones que tengan efectos legales o de igual importancia sobre individuos. Disqus no utiliza ni vende datos personales de consumidores con el propósito de entrenar grandes modelos de lenguaje. Nuestros sistemas de IA funcionan de acuerdo con nuestros compromisos de privacidad descritos en esta política.

#### 2. LOS DATOS QUE RECOPILAMOS SOBRE TI

Datos personales, o información personal, significa cualquier información sobre una persona que pueda razonablemente vincularse, directa o indirectamente, con una persona o hogar específico.

Recopilamos, y hemos recopilado en los últimos doce (12) meses, los siguientes tipos de datos personales sobre los usuarios:

1\. **Identificadores ("Datos de Identidad")** como nombre, apellido, nombre de usuario u identificador similar, dirección de protocolo de internet (IP), identificador de cookie único, ID de dispositivo, fecha de nacimiento, dirección de correo electrónico, número de teléfono y dirección postal;

2\. **Las categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California (Código Civil de California § 1798.80(e)),** que incluye cualquier información que identifique, se relacione, describa o pueda asociarse con una persona concreta pero que no esté disponible públicamente para el público general a través de registros gubernamentales federales, estatales o locales;

3\. **Características de clasificación** protegidas bajo la ley de California o federal, como raza, género o edad;

4\. **Información sobre la actividad en Internet u otras redes electrónicas,** como historial de navegación y comentarios, comentarios y respuestas a encuestas, tus datos de acceso, tipo y versión del navegador, configuración y ubicación de zona horaria, tipos y versiones de complementos del navegador, sistema operativo y plataforma, y otra tecnología en los dispositivos que utilices para acceder al Servicio;

5\. **Información profesional o relacionada con el empleo,** en la medida en que la incluyas en tu perfil o comentarios, o que pueda inferirse a partir de las páginas que visitas;

6\. **Información educativa,** en la medida en que la incluyas en tu perfil o comentarios, o que pueda inferirse a partir de las páginas que visitas;

7\. **Información personal sensible.** No recopilamos intencionadamente datos personales sobre tu raza o etnia, creencias religiosas o filosóficas, vida sexual, orientación sexual, opiniones políticas, afiliación sindical, información sobre tu salud o datos genéticos o biométricos, ni información sobre condenas y delitos penales. Sin embargo, si haces comentarios utilizando el Servicio que incluyen dichos datos sobre ti, estos serán disponibles públicamente y pueden ser procesados por Disqus u otros. Además, recopilamos y compartimos información obtenida mediante cookies o tecnología de seguimiento similar sobre las páginas web que has visitado, lo que puede permitir que los terceros con los que compartimos tu información hagan inferencias sobre ti que puedan constituir información personal sensible.

También podemos combinar, desidentificar o agregar cualquiera de la información que recopilamos a través de nuestro Servicio para cualquiera de los fines descritos a continuación.

#### 3. ¿CÓMO SE RECOPILAN TUS DATOS PERSONALES?

Utilizamos diferentes métodos para recopilar datos de ti y sobre ti, incluyendo a través de:

**Interacciones directas**

Esto incluye los datos personales que proporcionas al crear una cuenta o dejar un comentario.
​

**Tecnologías o interacciones automatizadas**

Mientras interactúas con nuestro Servicio, podemos recopilar automáticamente datos técnicos sobre tu equipo, acciones de navegación y patrones. Específicamente, de las siguientes maneras:

**Galletas**

Una cookie es un pequeño archivo digital que se coloca en el disco duro de tu ordenador. Puedes negarte a aceptar cookies de navegador activando la configuración de tu navegador o, en algunos casos, interactuando con banners emergentes de cookies. Utilizamos cookies colocadas en sitios web de terceros en los que el Servicio está habilitado para recopilar información sobre cómo interactúas con esos sitios incluso si no dejas comentarios, no respondes a encuestas o interactúas directamente con el Servicio en esos sitios, así como información sobre los otros sitios web que visitas. Disqus utiliza cookies y permite a los socios también configurar cookies a través del Servicio para facilitar la publicidad conductual entre contextos. Esto significa en la práctica que usamos cookies para ayudar a determinar qué anuncios ves en línea, registrando tus visitas a muchos sitios web de anunciantes o marcas, y luego mostrándote anuncios de productos y servicios similares.

Disqus utiliza cookies de 'autenticación', por ejemplo, sessionid, disqusauth y disqusauths, para mantenerte conectado desde tu navegador web y personalizar tu experiencia con Disqus.
​

Disqus utiliza cookies 'únicas', por ejemplo, disqus_unique y \_jid, para asociar actividades basadas en la web con una carga de página y con un navegador web, y para entender tus intereses y el uso del producto.
​

Cuando Disqus carga anuncios, utilizamos tecnologías de servicio de anuncios de Google que pueden establecer cookies con fines de marketing personalizado, asociar anuncios con actividades posteriores y limitar la frecuencia con la que se muestran anuncios específicos.
​

**Información del archivo de registro**

Los registros del servidor recopilan datos técnicos como tu dirección IP, tipo de navegador e información sobre el número de clics y cómo interactúas con los enlaces del Servicio, sitios de socios, nombres de dominio, páginas de destino, páginas vistas y otra información similar.
​

**Píxeles y rastreadores similares**

Cuando utilizas el Servicio, empleamos píxeles, gifs claros (también conocidos como webbeacons) que se emplean para recopilar datos técnicos e información como patrones de uso en línea. También usamos gifs claros en correos electrónicos basados en HTML enviados a nuestros usuarios para rastrear qué correos se abren y qué enlaces o anuncios son clicados por los destinatarios. También usamos píxeles y gifs claros en sitios web de terceros en los que el Servicio está habilitado para recopilar información sobre cómo interactúas con esos sitios, incluso si no dejas comentarios, no respondes a encuestas o interactúas directamente con el Servicio en esos sitios.
​

**Terceros o fuentes públicas**

Obtenemos o recibimos datos personales sobre ti de proveedores de análisis como Google; Socios publicitarios [Aquí](#cookies-and-data-recipients); y intermediarios de datos externos que venden datos personales. Obtenemos o recibimos datos personales de conexiones o inicios de sesión de terceros a través de plataformas de redes sociales como Facebook Connect, Google o Twitter/X cuando "sigues", "me gusta" o vinculas tu cuenta al Servicio. Ten en cuenta que algunos de estos proveedores, en particular proveedores de análisis como Google, pueden procesar datos de residentes del Espacio Económico Europeo (EEE) fuera del EEE.
​

**Gestión de cookies en tu navegador**

Quizá puedas ajustar la configuración de tu navegador para gestionar tus preferencias de cookies, como configurar tu navegador para que te notifique cuando recibas una cookie y te dé la opción de decidir si aceptarla o no. Si rechazas las cookies, puedes seguir usando nuestro sitio, pero la funcionalidad de algunas áreas puede ser limitada.
​

A continuación se muestran enlaces a información sobre cómo gestionar tus preferencias de cookies en navegadores comunes:

· Cookies de Google Chrome: Google Chrome Cookies

· Cookies de Mozilla Firefox: Mozilla Firefox Cookies

· Cookies de Internet Explorer: Internet Explorer Cookies

· Cookies Safari: Galletas Safari

· Google Analytics: Google Analytics

#### 4. SOCIOS DE PUBLICIDAD Y PUBLICIDAD DIRIGIDA

**Publicidad dirigida**

La publicidad es la principal forma en que Disqus gana dinero. Los ingresos por publicidad permiten a Disqus operar, apoyar y mejorar el Servicio. Disqus utiliza e intercambia con socios publicitarios y afiliados terceros, identificadores de cookies, identificadores de dispositivos (incluidos móviles), direcciones de correo electrónico hasheadas, dirección IP, información del proveedor de servicios de Internet (ISP) y del navegador, datos demográficos o de interés, contenido visto y acciones realizadas en el Servicio, en sitios asociados u otros sitios de terceros. Esto incluye información sobre los sitios web que has visitado y los anuncios con los que has interactuado para ofrecerte publicidad más relevante y dirigida a tus preferencias e intereses. Esto puede derivarse de tu interacción con el Servicio, sitios asociados u otros sitios web de terceros. Para una lista de socios publicitarios externos con los que Disqus está trabajando actualmente, véase [Aquí](#cookies-and-data-recipients).
​

**Marketing por correo electrónico**

Disqus también puede enviarte boletines por correo electrónico y mensajes de marketing por correo electrónico si nos has dado permiso o has consentido recibir dichos correos, según lo requiera la jurisdicción en la que residas. Los mensajes de email marketing pueden adaptarse a tus intereses basándose en la información descrita anteriormente en esta sección. Para información sobre cómo optar por no participar y ejercer tus derechos de privacidad, consulta [Aquí](#updating-your-account-settings).
​

**Socios publicitarios y divulgaciones de terceros**

Colaboramos y compartimos datos con terceros que recopilan información a través de diversos canales, tanto presenciales como online, con el fin de ofrecer publicidad más relevante para ti o tu negocio. Nuestros socios utilizan esta información para reconocerte a través de diferentes canales y plataformas, a lo largo del tiempo (incluyendo, pero no limitado a, ordenadores, dispositivos móviles, televisión direccionable u otros medios), con fines de marketing, análisis, atribución y reportes. Aunque Disqus no hace inferencias basadas en tus datos, nuestros socios publicitarios pueden extraer inferencias de tus datos para entender tus preferencias, características, tendencias psicológicas, predisposiciones, comportamiento, actitudes, inteligencia, habilidades y aptitudes.
​

En los últimos doce (12) meses, hemos compartido o vendido los siguientes datos:

· Identificadores;

· Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California (Código Civil de California § 1798.80(e);

· Características de clasificación protegida;

· información sobre la actividad de Internet u otras redes electrónicas;

· información profesional o relacionada con el empleo;

· Información educativa; y,

· Información personal sensible.

#### 5. CÓMO UTILIZAMOS LOS DATOS PERSONALES Y NUESTRA BASE LEGAL PARA SU USO

En el EEE, Reino Unido (UK) y Brasil, normalmente dependemos de tu consentimiento para usar datos personales. En algunos casos, utilizamos los datos cuando es necesario para nuestros intereses legítimos (o los de un tercero), pero solo cuando tus derechos e intereses no se ven afectados negativamente. Por último, también podemos utilizarlo cuando necesitemos cumplir con una obligación legal o regulatoria, o para proteger la salud, seguridad o derechos legales de cualquier persona.
​

No utilizamos información personal para realizar únicamente decisiones automatizadas que generen efectos legales o de igual importancia relevante sobre ti, conocido como "perfilado" según ciertas leyes.
​

A continuación exponemos una descripción de las formas en que utilizamos los datos personales y en cuáles de las bases legales nos apoyamos para hacerlo. También hemos identificado cuáles son nuestros intereses legítimos cuando es apropiado.
​

Ten en cuenta que podemos procesar tus datos personales por más de un motivo legal dependiendo del propósito específico para el que utilizamos tus datos. Por favor, contáctanos si tienes alguna pregunta adicional.

-   **Propósito o actividad**
    -   Tipo de datos personales
    -   Fuente de los datos personales
    -   Base de procesamiento
-   **Para registrarte como nuevo usuario**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Actividad en Internet o en redes electrónicas
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Cumplimiento de un contrato contigo · Con tu consentimiento
-   **Para gestionar nuestra relación contigo, lo que puede incluir notificarte sobre cambios en nuestros términos o política de privacidad, pedirte que dejes una reseña o respondas a una encuesta**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Cumplimiento de un contrato contigo · Necesario para cumplir con una obligación legal · Necesario para nuestros intereses legítimos (mantener actualizados nuestros registros y estudiar cómo los clientes utilizan nuestro servicio) · Con tu consentimiento
-   **Administrar y proteger nuestro negocio y el Servicio (incluyendo resolución de problemas, análisis de datos, pruebas, soporte, informes y alojamiento de datos)**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Necesario para cumplir con una obligación legal · Necesario para nuestros intereses legítimos (por ejemplo, mantener actualizados nuestros registros y estudiar cómo los clientes utilizan nuestro servicio)
-   **Entregarte contenido y anuncios relevantes y medir o entender la efectividad de la publicidad que te ofrecemos**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas · Información profesional o relacionada con el empleo · Información educativa · \[Información sensible\]
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Consenso. Cuando has dado tu consentimiento explícito para nuestro uso de datos personales para ofrecerte contenido relevante y personalizado · Necesario para nuestros intereses legítimos
-   **Utilizar el análisis de datos para mejorar nuestro servicio**
    -   · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Necesario para nuestros intereses legítimos (mantener nuestra web actualizada y relevante, desarrollar nuestro negocio y informar nuestra estrategia de marketing)
-   **Para enviarte boletines y correos promocionales que puedan interesarte**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas · Información profesional o relacionada con el empleo · Información educativa · \[Información sensible\]
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Consenso. Cuando has dado tu consentimiento explícito para nuestro uso de datos personales para ofrecerte contenido relevante y personalizado · Necesario para nuestros intereses legítimos (para desarrollar nuestro servicio y hacer crecer nuestro negocio)
-   **Vender o compartir datos personales con terceros con fines de marketing y publicidad**
    -   · Identificadores · Categorías de Información Personal listadas en el Estatuto de Registros de Clientes de California · Características de clasificación protegidas bajo la ley de California o federal · Actividad en Internet o en redes electrónicas · Información profesional o relacionada con el empleo · Información educativa · \[Información sensible\]
    -   · Interacciones directas · Tecnologías automatizadas · Galletas · Información del archivo de registro · Gifs claros · Terceros o fuentes públicas
    -   · Consenso. Cuando has dado tu consentimiento explícito para nuestro uso de datos personales para ofrecerte contenido relevante y personalizado. · Interés legítimo. En particular, fuera del EEE, podemos vender o compartir datos según lo permitan las leyes locales aplicables, lo que puede no requerir tu consentimiento para ello.

#### 6. DIVULGACIÓN DE TUS DATOS PERSONALES

Podemos divulgar cualquiera de las categorías de datos personales que recopilemos de ti y sobre ti en las siguientes circunstancias, para los fines establecidos en la tabla anterior. En los últimos 12 meses, hemos vendido y/o compartido todas las categorías de datos personales mencionadas anteriormente con Zeta Global y nuestros socios comerciales (como se describe a continuación).
​

**Zeta Global**

Los datos de los usuarios de Disqus se comparten con nuestra empresa afiliada Zeta Global con fines de marketing, incluyendo publicidad conductual cruzada en contexto. Los datos de Disqus son fusionados por Zeta con datos de otras fuentes online y offline para informar decisiones y mensajes de marketing en nombre de los clientes anunciantes de Zeta (marcas y agencias). Zeta también puede compartir datos obtenidos de Disqus con estos clientes anunciantes.
​

**Socios comerciales**

También compartimos datos personales con terceros externos con fines similares a los de Zeta Global. Estos terceros incluyen otras empresas que participan en publicidad conductual intercontextual, y intermediarios de datos que pueden vender o compartir tus datos personales. También podemos compartir datos personales con anunciantes con los que ya hayas compartido tus categorías de interés basada en tu actividad online, con fines de marketing. Para una lista de terceros externos, consulte [Aquí](#cookies-and-data-recipients).
​

**Proveedores de servicios**

Dependemos de proveedores de servicios (por ejemplo, empresas que ofrecen alojamiento web o servicios de análisis) para el funcionamiento de nuestro sitio web y servicio, y estos proveedores pueden tener acceso a tus datos personales. Nuestros proveedores de servicios tienen prohibido contractualmente utilizar tus datos personales para sus propios fines, y están obligados a tratar tus datos personales y mantener su confidencialidad de la misma manera que nosotros.
​

**Transacciones comerciales**

Podemos buscar adquirir otros negocios o fusionarnos con ellos. Si ocurre un cambio en nuestro negocio, la nueva empresa puede utilizar tus datos personales de la misma manera que se indica en este aviso de privacidad. También podemos participar en un intercambio limitado de datos para evaluar y probar a nuestros posibles socios de datos. Cuando se realizan tales pruebas, los datos se cifran o codifican antes de compartirse. Además, dicho intercambio está sujeto a las disposiciones contractuales apropiadas y los datos se eliminan rápidamente una vez finalizadas las pruebas.
​

**Requisitos legales**

Además, podemos acceder, preservar y divulgar tus datos personales si consideramos que la ley, una orden judicial u otros procesos legales válidos lo exigen. También podemos acceder, preservar y divulgar dichos datos personales si creemos de buena fe que la divulgación es necesaria para proteger tus derechos, nuestros o los de otros, o para investigar fraudes.

Exigimos a todos los terceros con los que compartimos datos que tomen las medidas adecuadas para garantizar la seguridad de sus datos personales y tratarlos conforme a todas las leyes aplicables.

#### 7. RETENCIÓN DE DATOS

Solo conservaremos tus datos personales durante el tiempo necesario para cumplir con los fines para los que los recopilamos, incluyendo el cumplimiento de cualquier requisito legal, contable o de informe.
​

Para determinar el periodo adecuado de conservación de los datos personales, consideramos la cantidad, naturaleza y sensibilidad de los mismos, el riesgo potencial de daño por el uso o divulgación no autorizados de sus datos personales, los fines para los cuales tratamos sus datos personales y si podemos lograr esos fines por otros medios, así como los requisitos legales aplicables.

#### 8. SEGURIDAD

Disqus utiliza salvaguardas comercialmente razonables para preservar la integridad y seguridad de toda la información recopilada a través del Servicio. Para proteger tu privacidad y seguridad, tomamos medidas razonables (como solicitar una contraseña única) para verificar tu identidad antes de concederte acceso a tu cuenta. Eres responsable de mantener el secreto de tu contraseña y información de cuenta, así como de controlar el acceso a tus comunicaciones por correo electrónico desde Disqus. Disqus no se hace responsable de la funcionalidad ni de las medidas de seguridad de ningún tercero.

#### 9. TUS DERECHOS

Dependiendo de las leyes donde vivas, puedes tener uno o más de los siguientes derechos según las leyes locales. Disqus extiende estos derechos a todas las personas, independientemente de dónde vivan, incluyendo el derecho a:

1\. **Solicita una copia** de los datos personales que hemos recopilado sobre ti;

2\. **Optar por no participar en nuestra venta o compartir** datos personales sobre ti;

3\. **Exclusión de recibir correos electrónicos** del Servicio;

4\. **Solicitamos que** eliminemos los datos que hemos recopilado sobre usted;

5\. **Solicita que corregamos** datos incorrectos;

6\. **Solicita que limitemos nuestro uso o eliminemos datos personales sensibles** que hayas podido haber proporcionado previamente en un comentario (debido a limitaciones técnicas, todas estas solicitudes se tratarán como solicitudes para eliminar datos).

Disqus no te discriminará por ejercer tus derechos de privacidad. Antes de cumplir una solicitud de copia de tu información personal, Disqus está obligado a verificar razonablemente tu identidad, lo cual generalmente hacemos enviando un enlace de verificación a la dirección de correo electrónico asociada a tu información personal. Según las leyes aplicables, puedes utilizar un agente autorizado para hacer una solicitud en tu nombre, pero ese agente debe poder completar nuestro proceso de verificación para demostrar que ha sido autorizado para hacer la solicitud.
​

También tiene derecho en muchos países a contactar con la autoridad correspondiente de privacidad o protección de datos si considera que no estamos cumpliendo con las leyes de privacidad y no hemos podido resolver la situación a su satisfacción. Si resides en el Espacio Económico Europeo, puedes consultar la autoridad local de protección de datos Aquí. Los residentes del Reino Unido pueden contactar con la Oficina del Comisionado de Información Aquí.
​

Visita nuestra página de Opciones de Privacidad Página de Opciones de Privacidad Para más detalles sobre estos derechos o para ejercerlos.
​

**Solicitudes de derechos**

Según la ley de California, estamos obligados a publicar estadísticas sobre cuántas personas ejercieron sus derechos de privacidad en el año anterior. Las siguientes estadísticas corresponden a las personas que realizaron solicitudes relacionadas con Disqus durante los doce meses que terminaron el 31 de diciembre de 2025. Estas estadísticas cubren todas las solicitudes recibidas de personas de todo el mundo.
​

· SOLICITUDES PARA SABER / ACCEDER A DATOS: Recibido: 258 \| Cumplido: 258 \| Denegado: 0 \| Tiempo de respuesta mediano: 1,19 días \| Tiempo medio de respuesta: 1,78 días

· SOLICITUDES DE ELIMINACIÓN: Recibido: 1.175 \| Cumplido: 1.175 \| Denegado: 0 \| Tiempo medio de respuesta: 0,83 días \| Tiempo medio de respuesta: 1,33 días

· SOLICITUDES PARA NO PARTICIPAR EN LA VENTA O COMPARTIR: Recibido: 1.303 \| Cumplido: 1.303 \| Denegado: 0 \| Tiempo de respuesta mediano: 0,85 días \| Tiempo medio de respuesta: 1,34 días

· SOLICITUDES PARA LIMITAR EL USO DE INFORMACIÓN PERSONAL SENSIBLE: Recibido: 280 \| Cumplido: 280 \| Denegado: 0 \| Tiempo de respuesta mediano: 1,00 día \| Tiempo medio de respuesta: 1,44 días

· SOLICITUDES NO ATENDIDAS PORQUE EL SOLICITANTE NO COMPLETÓ LA VERIFICACIÓN: 595

· SOLICITUDES DE ELIMINACIÓN NO CUMPLIDAS TOTAL O PARCIALMENTE: 0 — Disqus no denegó ninguna solicitud de eliminación en 2025.

#### 10. TRANSFERENCIAS INTERNACIONALES DE DATOS

Para usuarios con base en el Espacio Económico Europeo (EEE) y el Reino Unido, podemos compartir sus datos personales dentro del Grupo Disqus o con terceros externos. Esto puede implicar transferir tus datos fuera del EEE. Específicamente, los datos serán procesados por equipos de soporte técnico en Estados Unidos, India y Filipinas.
​

Siempre que sus datos personales se procesan fuera del EEE, los protegemos mediante contratos específicos aprobados por la Comisión Europea y/o las disposiciones contractuales equivalentes aprobadas por la Oficina del Comisionado de Información del Reino Unido, que garantizan que los datos personales reciban la misma protección que en Europa, independientemente de dónde se procesen.

#### 11. NO RASTREAR / CONTROL GLOBAL DE PRIVACIDAD ("GPC")

Disqus reconoce y procesa las preferencias del usuario tal como se establecen en las señales de "No rastrear" basadas en el navegador mediante la función de Control Global de Privacidad ("GPC").

#### 12. GENERAL

**Contacto**

Si tiene alguna pregunta sobre esta Política de Privacidad, por favor envíenos un correo electrónico a privacy@disqus.com o contáctenos por correo en 3 Park Avenue, 33rd Floor, Nueva York, NY 10016.
​

**Cambios en la Política de Privacidad**

Disqus puede, a su entera discreción, modificar o actualizar esta Política de Privacidad de vez en cuando, por lo que deberías revisar esta página periódicamente. Cuando cambiemos la política, actualizaremos la fecha de 'última modificación' en la parte superior de esta página. El uso continuado del Sitio tras la publicación de cualquier cambio en esta política significa que aceptas dichos cambios.
​

Clic [Aquí](#terms-of-service) para consultar los Términos de Servicio.

Nuestra Política de privacidad también está disponible en los siguientes idiomas:

[English](#disqus-privacy-policy)
[Deutsch](#disqus-datenschutzrichtlinie)
[Français](#politique-de-confidentialite-de-disqus)
[Italiano](#disqus-informativa-sulla-riservatezza)
[Português](#politica-de-privacidade-do-disqus)

### Política de Privacidade do Disqus {#politica-de-privacidade-do-disqus}

Política de Privacidade do Disqus

**Atualizado** a 10 de julho de 2026

Esta Política de Privacidade indica-lhe como a Disqus recolhe, utiliza, vende, divulga e protege dados relacionados consigo (o "Utilizador") em ligação com o nosso Serviço (conforme definido abaixo), bem como as suas escolhas relativamente à recolha e utilização destes dados.

#### 1. INTRODUÇÃO

**Visão geral**

A Disqus oferece uma plataforma online de comentários públicos e partilha de opiniões onde os utilizadores iniciam sessão e criam perfis para participar em conversas com colegas e desfrutar de uma experiência interativa nas secções de comentários, sondagens e outras funcionalidades interativas disponibilizadas neste site, bem como incorporadas em sites de terceiros. A utilização da nossa plataforma e software, bem como a interação com os nossos cookies ou tecnologias de rastreamento semelhantes (coletivamente o "Serviço"), quer neste site quer num site de terceiros, está sujeita aos termos desta Política de Privacidade. O Serviço é uma plataforma pública e a Disqus ou outros podem pesquisar, ver, usar ou republicar qualquer um dos seus Conteúdos de Utilizador (conforme definido nos nossos Termos de Uso) que publique através do Serviço. A Disqus é também uma empresa de marketing e dados, utilizando e partilhando dados pessoais recolhidos de sites de terceiros onde o nosso Serviço está ativado para fins de marketing, incluindo publicidade comportamental cruzada em contexto. Para mais informações sobre as nossas atividades de marketing, consulte a Secção 4: Publicidade Direcionada e Parceiros de Publicidade abaixo.
​

**Aplicabilidade a websites e serviços de terceiros**

A Disqus oferece um serviço de envolvimento online que outros sites utilizam para permitir a discussão e a interatividade entre os seus utilizadores. Esta Política de Privacidade aplica-se aos dados que a Disqus recolhe sobre os Utilizadores do Serviço e através de cookies em sites habilitados pelo Serviço, e não se aplica às práticas independentes de recolha de dados de qualquer site que utilize o Serviço ou outro site ligado ao Serviço. Para informações sobre como websites de terceiros recolhem e utilizam as suas informações pessoais, consulte as políticas de privacidade desses sites.
​

**Os Seus Direitos de Privacidade**

Tem direitos sobre as suas informações pessoais. Estes direitos são descritos com mais detalhe na Secção 9: Os Seus Direitos, abaixo. Pode exercer os seus direitos de privacidade de dados em aqui.
​

**Privacidade das Crianças**

O Serviço não se destina a ser utilizado por crianças com menos de 18 anos. Não recolhemos nem vendemos conscientemente informações pessoais de crianças com menos de 18 anos, nem permitimos conscientemente que essas pessoas se registem numa conta no serviço. Caso saibamos que recolhemos informações pessoais de uma criança com menos de 18 anos, iremos apagá-las. Se acredita que possamos ter recolhido informações pessoais de uma criança com menos de 18 anos, por favor contacte-nos ou submeta um pedido de direitos de privacidade de dados em aqui.
​

**A Nossa Utilização da Inteligência Artificial**

A Disqus utiliza inteligência artificial e aprendizagem automática de duas formas. Primeiro, a IA é usada para ajudar a moderar o conteúdo na plataforma — detetando spam e conteúdos que violem as nossas diretrizes comunitárias, para que o Serviço possa ser mantido seguro e funcional para os utilizadores. Em segundo lugar, a IA é usada para ajudar a fornecer publicidade mais relevante para si, com base nos seus interesses e atividade online. Em ambos os casos, a Disqus não utiliza IA para tomar decisões que tenham efeitos legais ou igualmente significativos sobre indivíduos. A Disqus não utiliza nem vende dados pessoais dos consumidores para fins de treino de grandes modelos de linguagem. Os nossos sistemas de IA funcionam de acordo com os nossos compromissos de privacidade, conforme descrito nesta política.

#### 2. OS DADOS QUE RECOLHEMOS SOBRE SI

Dados pessoais, ou informação pessoal, significa qualquer informação sobre uma pessoa que possa razoavelmente ser ligada, direta ou indiretamente, a uma pessoa ou agregado familiar específico.
​

Recolhemos, e recolhemos nos últimos doze (12) meses, os seguintes tipos de dados pessoais sobre os Utilizadores:

1\. **Identificadores ("Dados de Identidade")** como primeiro nome, apelido, nome de utilizador ou identificador semelhante, endereço de protocolo de internet (IP), ID de Cookie único, ID de Dispositivo, data de nascimento, endereço de email, número de telefone e endereço postal;

2\. **As categorias de Informação Pessoal listadas no Estatuto de Registos de Clientes da Califórnia (Código Civil da Califórnia § 1798.80(e)),** que inclui qualquer informação que identifique, se relacione, descreva, possa ser associada a um indivíduo em particular, mas que não esteja disponível publicamente ao público em geral através de registos governamentais federais, estaduais ou locais;

3\. **Características de classificação protegidas** pela lei da Califórnia ou federal, como raça, género ou idade;

4\. **Informação sobre atividade na Internet ou outras redes eletrónicas,** como histórico de navegação e comentários, feedback e respostas a inquéritos, os seus dados de login, tipo e versão do navegador, definição e localização do fuso horário, tipos e versões dos plug-ins do navegador, sistema operativo e plataforma e outras tecnologias nos dispositivos que utiliza para aceder ao Serviço;

5\. **Informação profissional ou relacionada com o emprego,** na medida em que a inclua no seu perfil ou comentários, ou que possa ser inferida a partir das páginas que consulta;

6\. **Informação educativa,** na medida em que a inclua no seu perfil ou comentários, ou que possa ser inferida a partir das páginas que consulta;

7\. **Informação Pessoal Sensível.** Não recolhemos intencionalmente quaisquer dados pessoais sobre a sua raça ou etnia, crenças religiosas ou filosóficas, vida sexual, orientação sexual, opiniões políticas, filiação sindical, informações sobre a sua saúde ou dados genéticos ou biométricos, ou informações sobre condenações e crimes criminais. No entanto, se fizer comentários utilizando o Serviço que incluam tais dados sobre si, estes ficarão disponíveis publicamente e poderão ser processados pela Disqus ou por outros. Além disso, recolhemos e partilhamos informações recolhidas através de cookies ou tecnologia de rastreamento semelhante sobre as páginas web que visualizou, o que pode permitir que os terceiros com quem partilhamos a sua informação façam inferências sobre si que possam constituir informações pessoais sensíveis.

Também podemos combinar, desidentificar ou agregar qualquer uma das informações que recolhemos através do nosso Serviço para qualquer um dos fins descritos abaixo.

#### 3. COMO SÃO RECOLHIDOS OS SEUS DADOS PESSOAIS?

Utilizamos diferentes métodos para recolher dados sobre si, incluindo através de:

**Interações diretas**

Isto inclui dados pessoais que fornece ao criar uma conta ou deixar um comentário.
​

**Tecnologias ou interações automatizadas**

À medida que interage com o nosso Serviço, podemos recolher automaticamente Dados Técnicos sobre o seu equipamento, ações de navegação e padrões. Especificamente, das seguintes formas:
​

**Cookies**

Um cookie é um pequeno ficheiro digital colocado no disco rígido do seu computador. Pode recusar aceitar cookies do navegador ativando as definições do seu navegador ou, em alguns casos, interagindo com banners pop-up de cookies. Utilizamos cookies colocados em sites de terceiros onde o Serviço está habilitado para recolher informações sobre como interage com esses sites, mesmo que não deixe comentários, responda a sondagens ou interaja diretamente com o Serviço nesses sites, bem como informações sobre os outros sites que visita. O Disqus utiliza cookies e permite que os parceiros também definam cookies através do Serviço para facilitar a publicidade comportamental entre contextos. Isto significa, na prática, que usamos cookies para ajudar a determinar que anúncios vê online, registando as suas visitas a muitos sites de anunciantes/marcas e depois mostrando-lhe anúncios de produtos e serviços semelhantes.
​

O Disqus utiliza cookies de 'autenticação', por exemplo, sessionid, disqusauth e disqusauths, para o manter ligado a partir do seu navegador e personalizar a sua experiência com o Disqus.
​

A Disqus utiliza cookies 'únicos', por exemplo, disqus_unique e \_jid, para associar atividades baseadas na web a uma carga de página e a um navegador web, e compreender os seus interesses e o uso do produto.
​

Quando a Disqus carrega anúncios, utilizamos tecnologias de inserção de anúncios da Google que podem definir cookies para fins de marketing personalizado, associar anúncios a atividades posteriores e limitar a frequência com que é mostrado anúncios específicos.
​

**Informação do Ficheiro de Registo**

Os registos do servidor recolhem dados técnicos como o seu endereço IP, tipo de navegador e informações sobre o número de cliques e como interage com links no Serviço, sites parceiros, nomes de domínio, páginas de destino, páginas visualizadas e outras informações semelhantes.
​

**Pixels e Rastreadores Semelhantes**

Quando utiliza o Serviço, utilizamos pixels, gifs claros (também conhecidos como web beacons) que são usados para recolher Dados Técnicos e informações, como padrões de utilização online. Também usamos gifs claros em emails baseados em HTML enviados aos nossos utilizadores para acompanhar que emails são abertos e que links ou anúncios são clicados pelos destinatários. Também usamos pixels e gifs claros em sites de terceiros onde o Serviço está ativado para recolher informações sobre como interage com esses sites, mesmo que não deixe comentários, não responda a sondagens ou interaja diretamente com o Serviço nesses sites.
​

**Terceiros ou fontes públicas**

Obtemos ou recebemos dados pessoais sobre si de fornecedores de análise como a Google; Parceiros de publicidade [aqui](#cookies-and-data-recipients); e corretores de dados terceiros que vendem dados pessoais. Obtemos ou recebemos dados pessoais de ligações de terceiros ou logins através de plataformas de redes sociais como Facebook Connect, Google ou Twitter/X quando "segue", "gosta" ou liga a sua conta ao Serviço. Por favor, note que alguns destes fornecedores, em particular os de análise como a Google, podem processar dados de residentes do Espaço Económico Europeu (EEE) fora do EEE.
​

**Gerir Cookies no Seu Navegador**

Pode ser possível ajustar as definições do seu navegador para gerir as preferências de cookies, como definir o navegador para o notificar quando receber um cookie e dar-lhe a escolha de decidir se o aceita ou não. Se rejeitar cookies, pode continuar a usar o nosso site, mas a funcionalidade de algumas áreas pode ser limitada.
​

Abaixo estão ligações para informações sobre como gerir as suas preferências de cookies em navegadores comuns:

· Cookies do Google Chrome: Google Chrome Cookies

· Cookies Mozilla Firefox: Mozilla Firefox Cookies

· Cookies do Internet Explorer: Internet Explorer Cookies

· Bolachas Safari: Safari Cookies

· Google Analytics: Google Analytics

#### 4. PUBLICIDADE DIRECIONADA E PARCEIROS DE PUBLICIDADE

**Publicidade Direcionada**

A publicidade é a principal forma pela qual a Disqus ganha dinheiro. A receita publicitária permite que a Disqus opere, apoie e melhore o Serviço. A Disqus utiliza e partilha com parceiros e afiliados de publicidade terceiros, IDs de cookies, IDs de dispositivos (incluindo móveis), endereços de email hashados, endereços IP, informações do Provedor de Serviços de Internet (ISP) e do navegador, dados demográficos ou de interesse, conteúdos visualizados e ações tomadas no Serviço, em sites parceiros ou noutros sites de terceiros. Isto inclui informações sobre os sites que visitou e os anúncios com que interagiu, de modo a fornecer publicidade mais relevante, direcionada às suas preferências e interesses. Isto pode derivar da sua interação com o Serviço, sites parceiros ou outros sites de terceiros. Para uma lista de parceiros publicitários terceiros com quem a Disqus está atualmente a trabalhar, consulte [aqui](#cookies-and-data-recipients).
​

**Email Marketing**

A Disqus também pode enviar-lhe newsletters e mensagens de marketing por email se nos tiver dado permissão ou consentido em receber esses emails, conforme exigido na jurisdição onde reside. As mensagens de email marketing podem ser adaptadas aos seus interesses com base nas informações descritas acima nesta secção. Para informações sobre como optar por sair e exercer os seus direitos de privacidade, consulte [aqui](#updating-your-account-settings).
​

**Parceiros de Publicidade e Divulgações de Terceiros**

Fazemos parceria e partilhamos dados com terceiros que recolhem informações através de vários canais, incluindo offline e online, com o objetivo de disponibilizar publicidade mais relevante para si ou para o seu negócio. Os nossos parceiros utilizam esta informação para o reconhecer através de diferentes canais e plataformas, ao longo do tempo (incluindo, mas não se limitando a, computadores, dispositivos móveis, televisão endereçável ou outros meios), para fins de marketing, análise, atribuição e relatórios. Embora a Disqus não faça inferências com base nos seus dados, os nossos parceiros publicitários podem tirar inferências dos seus dados para compreender as suas preferências, características, tendências psicológicas, predisposições, comportamento, atitudes, inteligência, capacidades e aptidões.

Nos últimos doze (12) meses, partilhámos ou vendemos os seguintes dados:

· Identificadores;

· Categorias de Informação Pessoal listadas no Estatuto de Registos de Clientes da Califórnia (Código Civil da Califórnia § 1798.80(e);

· Características de classificação protegida;

· Informação sobre a atividade da Internet ou de outras redes eletrónicas;

· Informação profissional ou relacionada com o emprego;

· Informação educativa; e,

· Informação Pessoal Sensível.

#### 5. COMO UTILIZAMOS OS DADOS PESSOAIS E A NOSSA BASE LEGAL PARA O USO

No EEE, Reino Unido (UK) e Brasil, normalmente dependemos do seu consentimento para utilizar dados pessoais. Em alguns casos, usamos os dados conforme necessário para os nossos interesses legítimos (ou de terceiros), mas apenas quando os seus direitos e interesses não são negativamente afetados. Por fim, também podemos utilizá-lo quando precisamos de cumprir uma obrigação legal ou regulatória, ou para proteger a saúde, segurança ou direitos legais de qualquer pessoa.
​

Não utilizamos informações pessoais para apoiar decisões exclusivamente automatizadas que produzam efeitos legais ou igualmente significativos sobre si, conhecidos como "perfilamento" segundo certas leis.
​

Abaixo apresentamos uma descrição das formas como utilizamos os dados pessoais e quais as bases legais em que nos baseamos para tal. Também identificámos quais são os nossos interesses legítimos quando apropriado.

Note que podemos tratar os seus dados pessoais para mais do que um motivo legal, dependendo do propósito específico para o qual estamos a utilizar os seus dados. Por favor, contacte-nos se tiver mais alguma questão.

-   **Propósito ou Atividade**
    -   Tipo de Dados Pessoais
    -   Fonte dos Dados Pessoais
    -   Base do Processamento
-   **Para o registar como novo utilizador**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Atividade na Internet ou em Rede Eletrónica
    -   · Interações Diretas · Tecnologias Automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Execução de um contrato consigo · Com o seu consentimento
-   **Para gerir a nossa relação consigo, que pode incluir notificá-lo sobre alterações aos nossos termos ou política de privacidade, pedir-lhe para deixar uma avaliação ou responder a um inquérito**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Execução de um contrato consigo · Necessário para cumprir uma obrigação legal · Necessário para os nossos interesses legítimos (para manter os nossos registos atualizados e estudar como os clientes utilizam o nosso Serviço) · Com o seu consentimento
-   **Administrar e proteger o nosso negócio e o Serviço (incluindo resolução de problemas, análise de dados, testes, suporte, relatórios e alojamento de dados)**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Necessário para cumprir uma obrigação legal · Necessário para os nossos interesses legítimos (por exemplo, manter os nossos registos atualizados e estudar como os clientes utilizam o nosso Serviço)
-   **Para lhe entregar conteúdos e anúncios relevantes e medir ou compreender a eficácia da publicidade que lhe servimos**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica · Informação Profissional ou Relacionada com o Emprego · Informação Educativa · \[Informação Sensível\]
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Consentimento. Quando deu consentimento explícito para o uso de dados pessoais para lhe fornecer conteúdo relevante e personalizado · Necessário para os nossos interesses legítimos
-   **Utilizar a análise de dados para melhorar o nosso Serviço**
    -   · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Necessário para os nossos interesses legítimos (manter o nosso site atualizado e relevante, desenvolver o nosso negócio e informar a nossa estratégia de marketing)
-   **Para lhe enviar newsletters e emails promocionais que possam interessar-lhe**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica · Informação Profissional ou Relacionada com o Emprego · Informação Educativa · \[Informação Sensível\]
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Consentimento. Quando deu consentimento explícito para o uso de dados pessoais para lhe fornecer conteúdo relevante e personalizado · Necessário para os nossos interesses legítimos (para desenvolver o nosso Serviço e fazer crescer o nosso negócio)
-   **Vender ou partilhar dados pessoais com terceiros para fins de marketing e publicidade**
    -   · Identificadores · Categorias de Informação Pessoal listadas no Estatuto dos Registos de Clientes da Califórnia · Características de Classificação Protegidas ao abrigo da lei da Califórnia ou federal · Atividade na Internet ou em Rede Eletrónica · Informação Profissional ou Relacionada com o Emprego · Informação Educativa · \[Informação Sensível\]
    -   · Interações Diretas · Tecnologias automatizadas · Cookies · Informação do Ficheiro de Registo · Gifs Claros · Terceiros ou fontes públicas
    -   · Consentimento. Quando deu consentimento explícito para o uso de dados pessoais para lhe fornecer conteúdo relevante e personalizado. · Interesse legítimo. Em particular, fora do EEE, podemos vender ou partilhar dados conforme permitido pelas leis locais aplicáveis, o que pode não exigir o seu consentimento para tal.

#### 6. DIVULGAÇÃO DOS SEUS DADOS PESSOAIS

Podemos divulgar qualquer uma das categorias de dados pessoais que recolhemos de si e sobre si nas seguintes circunstâncias, para os fins definidos na tabela acima. Nos últimos 12 meses, vendemos e/ou partilhámos todas as categorias de dados pessoais acima referidas com a Zeta Global e os nossos Parceiros de Negócio (conforme descrito abaixo).
​

**Zeta Global**

Os dados dos utilizadores Disqus são partilhados com a nossa empresa afiliada Zeta Global para fins de marketing, incluindo publicidade comportamental cruzada em contexto. Os dados da Disqus são fundidos pela Zeta com dados de outras fontes online e offline para informar decisões e mensagens de marketing em nome dos clientes anunciantes da Zeta (marcas e agências). A Zeta pode também partilhar dados obtidos da Disqus com estes clientes anunciantes.
​

**Parceiros de Negócios**

Também partilhamos dados pessoais com terceiros externos para fins semelhantes aos da Zeta Global. Estes terceiros incluem outras empresas que participam em publicidade comportamental entre contextos e corretores de dados que podem vender ou partilhar os seus dados pessoais. Também podemos partilhar dados pessoais com anunciantes com quem já partilhou as suas categorias de interesse em dados com base na sua atividade online, para fins de marketing. Para uma lista de terceiros externos, consulte [aqui](#cookies-and-data-recipients).
​

**Prestadores de Serviços**

Dependemos de prestadores de serviços (por exemplo, empresas que fornecem alojamento web ou serviços de análise) para a operação do nosso site e Serviço, e estes prestadores de serviços podem ter acesso aos seus dados pessoais. Os nossos prestadores de serviços estão contratualmente proibidos de usar os seus dados pessoais para fins próprios e são obrigados a tratar os seus dados pessoais e manter a sua confidencialidade da mesma forma que nós.
​

**Transações Comerciais**

Podemos tentar adquirir outros negócios ou fundir-nos com eles. Se ocorrer uma alteração no nosso negócio, a nova empresa poderá utilizar os seus dados pessoais da mesma forma que está indicado neste aviso de privacidade. Também podemos participar numa partilha limitada de dados para avaliar e testar os nossos potenciais parceiros de dados. Quando tais testes ocorrem, os dados são encriptados ou codificados antes de serem partilhados. Além disso, tal partilha está sujeita às disposições contratuais adequadas e os dados são prontamente eliminados após a conclusão dos testes.
​

**Requisitos legais**

Além disso, podemos aceder, preservar e divulgar os seus dados pessoais, se considerarmos que tal é exigido por lei, ordem judicial ou outros processos legais válidos. Também podemos aceder, preservar e divulgar tais dados pessoais se acreditarmos de boa-fé que a divulgação é necessária para proteger os seus direitos, propriedade ou segurança ou de outros, ou para investigar fraudes.
​

Exigimos que todos os terceiros com quem partilhamos dados tomem as medidas adequadas para garantir a segurança dos seus dados pessoais e tratem-nos de acordo com todas as leis aplicáveis.

#### 7. RETENÇÃO DE DADOS

Só reteremos os seus dados pessoais pelo tempo necessário para cumprir os propósitos para os quais os recolhemos, incluindo o cumprimento de quaisquer requisitos legais, contabilísticos ou de reporte.
​

Para determinar o período adequado de retenção dos dados pessoais, consideramos a quantidade, natureza e sensibilidade dos dados pessoais, o risco potencial de dano devido ao uso ou divulgação não autorizada dos seus dados pessoais, os fins para os quais tratamos os seus dados pessoais e se podemos alcançar esses fins por outros meios, bem como os requisitos legais aplicáveis.

#### 8. SEGURANÇA

A Disqus utiliza salvaguardas comercialmente razoáveis para preservar a integridade e segurança de toda a informação recolhida através do Serviço. Para proteger a sua privacidade e segurança, tomamos medidas razoáveis (como solicitar uma palavra-passe única) para verificar a sua identidade antes de lhe conceder acesso à sua conta. É responsável por manter o segredo da sua palavra-passe única e informações da conta, bem como por controlar o acesso às suas comunicações por email a partir da Disqus. A Disqus não é responsável pela funcionalidade ou pelas medidas de segurança de terceiros.

#### 9. OS SEUS DIREITOS

Dependendo das leis onde vive, pode ter um ou mais dos seguintes direitos ao abrigo das leis locais. Disqus estende estes direitos a todos os indivíduos, independentemente de onde vivam, incluindo o direito a:

1\. **Solicite uma cópia** dos dados pessoais que recolhemos sobre si;

2\. **Optar por não participar na nossa venda ou partilha** de dados pessoais sobre si;

3\. **Optar por não receber emails** do Serviço;

4\. **Solicite que eliminemos** os dados que recolhemos sobre si;

5\. **Solicita que corrigamos** dados incorretos;

6\. **Solicite que limitemos o nosso uso ou eliminemos dados pessoais sensíveis** que possa ter fornecido anteriormente num comentário (devido a limitações técnicas, todos estes pedidos serão tratados como pedidos para eliminar dados).

A Disqus não irá discriminá-lo por exercer os seus direitos de privacidade. Antes de cumprir um pedido de cópia das suas informações pessoais, a Disqus é obrigada a verificar razoavelmente a sua identidade, o que geralmente fazemos enviando um link de verificação para o endereço de email associado às suas informações pessoais. De acordo com as leis aplicáveis, pode usar um agente autorizado para fazer um pedido em seu nome, mas esse agente deve ser capaz de completar o nosso processo de verificação para demonstrar que foi autorizado a fazer o pedido.
​

Também tem o direito, em muitos países, de contactar a autoridade relevante de privacidade ou proteção de dados se considerar que não estamos a cumprir as leis de privacidade e não conseguimos resolver a situação de forma satisfatória. Se for residente do Espaço Económico Europeu, pode consultar a autoridade local de proteção de dados aqui. Os residentes do Reino Unido podem contactar o Gabinete do Comissário de Informação aqui.
​

Visite a nossa página de Escolhas de Privacidade Página de Escolhas de Privacidade Para mais detalhes sobre estes direitos ou para os exercer.
​

**Pedidos de Direitos**

Ao abrigo da lei da Califórnia, somos obrigados a publicar estatísticas sobre quantas pessoas exerceram os seus direitos de privacidade no ano anterior. As estatísticas seguintes referem-se a pessoas que fizeram pedidos relacionados com a Disqus durante os doze meses que terminaram a 31 de dezembro de 2025. Estas estatísticas abrangem todos os pedidos recebidos de indivíduos em todo o mundo.
​

· PEDIDOS PARA SABER / ACEDER A DADOS: Recebidos: 258 \| Cumprido: 258 \| Negado: 0 \| Tempo mediano de resposta: 1,19 dias \| Tempo médio de resposta: 1,78 dias

· PEDIDOS DE ELIMINAÇÃO: Recebido: 1.175 \| Cumprido: 1.175 \| Negado: 0 \| Tempo mediano de resposta: 0,83 dias \| Tempo médio de resposta: 1,33 dias

· PEDIDOS PARA OPTAR POR NÃO VENDER OU PARTILHAR: Recebido: 1.303 \| Cumprido: 1.303 \| Negado: 0 \| Tempo mediano de resposta: 0,85 dias \| Tempo médio de resposta: 1,34 dias

· PEDIDOS PARA LIMITAR O USO DE INFORMAÇÕES PESSOAIS SENSÍVEIS: Recebido: 280 \| Cumprido: 280 \| Negado: 0 \| Tempo mediano de resposta: 1,00 dia \| Tempo médio de resposta: 1,44 dias

· PEDIDOS NÃO PROCESSADOS PORQUE O REQUERENTE NÃO COMPLETOU A VERIFICAÇÃO: 595

· PEDIDOS DE ELIMINAÇÃO NÃO CUMPRIDOS TOTAL OU PARCIALMENTE: 0 — A Disqus não negou quaisquer pedidos de eliminação em 2025.

#### 10. TRANSFERÊNCIAS INTERNACIONAIS DE DADOS

Para utilizadores sediados no Espaço Económico Europeu (EEE) e no Reino Unido, podemos partilhar os seus dados pessoais dentro do Grupo Disqus ou com terceiros externos. Isto pode envolver a transferência dos seus dados para fora do EEE. Especificamente, os dados serão processados por equipas de apoio técnico nos Estados Unidos, Índia e Filipinas.
​

Sempre que os seus dados pessoais são processados fora do EEE, protegemo-los através de contratos específicos aprovados pela Comissão Europeia e/ou das disposições contratuais equivalentes aprovadas pelo Gabinete do Comissário de Informação do Reino Unido, que garantem que os dados pessoais recebem a mesma proteção que têm na Europa, independentemente do local onde são processados.

#### 11. NÃO RASTREAR / CONTROLO GLOBAL DE PRIVACIDADE ("GPC")

O Disqus reconhece e processa as preferências do utilizador conforme definido nos sinais baseados no navegador "Não Rastrear" através da funcionalidade Global de Controlo de Privacidade ("GPC").

#### 12. GERAL

**Contacto**

Se tiver alguma questão sobre esta Política de Privacidade, por favor envie-nos um email para privacy@disqus.com ou contacte-nos por correio para 3 Park Avenue, 33rd Floor, Nova Iorque, NY 10016.
​

**Alterações à Política de Privacidade**

A Disqus pode, a seu critério exclusivo, modificar ou atualizar esta Política de Privacidade periodicamente, pelo que deve consultar esta página periodicamente. Quando alterarmos a política, iremos atualizar a data de 'última modificação' no topo desta página. A sua utilização contínua do Site após a publicação de quaisquer alterações a esta política significa que aceita essas alterações.
​

Clique [aqui](#terms-of-service) para consultar os Termos de Serviço.

Nossa Política de Privacidade também está disponível nos seguintes idiomas:
[Deutsch](#disqus-datenschutzrichtlinie)

[English](#disqus-privacy-policy)
[Español](#politica-de-privacidad-de-disqus)
[Français](#politique-de-confidentialite-de-disqus)
[Italiano](#disqus-informativa-sulla-riservatezza)
​

### Privacy FAQ {#privacy-faq}

**How do I delete my Disqus Account?**
​
You may delete your Disqus account by following the instructions found at this link: Delete My Disqus Account.

Please note that when you delete your Disqus account, your comments will no longer visible to the public, but the operator / publisher of the website where you left a comment will still be able to view your previous comments.

**How can I unsubscribe from Disqus emails?**

You can unsubscribe from the emails Disqus sends you by clicking the unsubscribe link found in any email you receive from us. You can also unsubscribe or manage your emails subscription by clicking this link: Manage My Disqus Email Settings.

Please note that unsubscribing from Disqus emails will unsubscribe you from Disqus notification emails and Disqus digest emails. You will still receive transactional emails necessary for the Disqus service to function; these include, but may not be limited to, emails confirming you created a new account, emails helping you reset your password etc.

**How do I reset my password?**

If you want to reset your password, please click here: Disqus Password Reset.

**What is Disqus doing to prepare for the General Data Protection Regulation (“GDPR”)?**

Disqus has been fully GDPR compliant since May 25, 2018, the ‘Effective Date’ of the GDPR. As part of our GDPR compliance program we have implemented new procedures to obtain your consent for the collection of your personal data both for processing by Disqus and, where applicable, third parties.

If you are a publisher using the Disqus SaaS commenting platform, we require that you take all measures to be compliant with the GDPR as of the Effective Date. This means that you will need to implement a means to obtain consent from EU citizens for the collection of personal data on your website.

For more information about the Disqus platform, please visit our Knowledge Base page by clicking here.

If you are a publisher looking for support resources, please click here.

### Spam {#spam}

What constitutes spam is constantly evolving. Spam can be generally described as unsolicited, repeated actions that negatively impact other users. The following types of content are common characteristics that may be viewed as spam and warrant removal:

-   Comments or discussions posted in large quantities to promote or sell a product/service. If you’re just posting links to a site of a business you operate or adding a link as a signature at the end of your comments, they may be flagged as spam.

-   The exact same comment posted repeatedly to disrupt a thread.

-   Following users multiple times even after they’ve removed you as a follower.

-   Following a large number of users for the purposes of disruption or self-promotion.

-   Posting off-topic discussions or comments not relevant to a community.

-   Profile display name, bio, location, or link that promotes or sell a product/service.

A good practice before posting is to review the community guidelines for the site, if provided. Moderators are the best people to inquire about a potential topic and whether or not it is considered spam. Also, take a moment to review recent comments in the community to see the posting requirements in action.

### Targeted harassment or encouraging others to do so {#targeted-harassment-or-encouraging-others-to-do-so}

Hate speech and other forms of targeted and systematic harassment of people have no place on Disqus, nor do we tolerate communities dedicated to fostering harassing behavior.
​
Factors that we examine when determining if conduct is considered to be targeted harassment include:

-   if the primary purpose of the reported account is to bully or attack someone so that they no longer feel safe or welcome in the community.

-   if the reported user is inciting others to harass another user.

-   if the reported user is posting harassing content directed at a user using multiple accounts.

You may find that some content on Disqus to be offensive, disrespectful, or that you disagree with. Disqus does not moderate or remove potentially offensive content unless there’s been a violation of the Basic Rules or Terms of Service.
​
For more information on how we enforce against abusive accounts that violate the Basic Rules and reporting abuse to Disqus, read our Abusive Behavior Policy.

### Terms of Service {#terms-of-service}

***If you create an account with Disqus, you agree to the User Terms of Service. If you are using Disqus comments on your website you are a “Publisher” and you also agree to the Publisher Terms of Service**which follow the User Terms of Service below.***

#### DISQUS USER TERMS OF SERVICE.

Disqus, Inc. (“Disqus”, “we”, “us” or “our”) offers an online public comment sharing platform where you may login and create profiles to participate in conversations with peers and enjoy an interactive experience. These Terms of Service (the “Terms”) govern your use of and access to our comment sharing platform, software and website (collectively the “Service”) by using the Service you understand and agree to be bound by these Terms.

THESE TERMS CONTAIN A MANDATORY ARBITRATION OF DISPUTES PROVISION THAT REQUIRES THE USE OF ARBITRATION ON AN INDIVIDUAL BASIS TO RESOLVE DISPUTES, RATHER THAN JURY TRIALS OR CLASS ACTIONS, AND ALSO LIMITS THE REMEDIES AVAILABLE TO YOU IN THE EVENT OF A DISPUTE.

**Use of the Service**.

You may only access and use the Service if you agree to be bound by these Terms, are over the age of 18, and are not a person barred from receiving or using the Services under the laws of the applicable jurisdiction. If you are accepting these Terms and using the Service on behalf of a company, organization, government or other legal entity, you represent and warrant that you are authorized to do so. In the event you breach these Terms, or violate the Basic Rules of Disqus, Disqus may, in our sole discretion, revoke your rights to use the Service and terminate your account.

**License to Use the Service.**

Disqus grants you a non-exclusive, limited, non-transferable, revocable license to access and use the Service in accordance with the Terms and in the manner contemplated hereunder. Disqus reserves all rights not expressly granted herein in and to the Service and the Disqus Content (as defined below). Disqus reserves the right to revoke your license to use the Service at any time and for any reason.

Disqus reserves the right to access, read, preserve, and disclose any information as we reasonably believe is necessary to (i) satisfy applicable law; (ii) enforce these Terms, including investigation of potential violations hereof; (iii) detect, prevent, or otherwise address fraud, security or other technical issues; (iv) respond to user support requests; (v) protect the rights, property or safety of Disqus; (vi) or as otherwise set forth in the Disqus Privacy Policy.

**Modifying or Discontinuing the Service.**

We are constantly changing and improving Service. We may, without prior notice to you, add or remove functionalities or features, and we may suspend or stop the Service altogether.

**Disqus Account**.

By creating a Disqus account, you agree to these Terms. When creating your account, you must provide accurate and complete information. You are solely responsible for the activity that occurs on your account, and you must keep your account password secure. We encourage you to use “strong” passwords (passwords that use a combination of upper and lowercase letters, numbers and symbols) with your account. You may never use another user’s account without permission. You must notify Disqus immediately of any breach of security or unauthorized use of your account. Disqus will not be liable for any losses caused by any unauthorized use of your account. You may control your User profile and how you interact with the Service by changing the settings in your profile settings.

**Privacy.**

The Disqus Privacy Policy describes how we use and process the information you provide to us when you use the Service. You understand that by using the Services you consent to the collection, use and disclosure of your information as set forth in our Privacy Policy.

**Content on the Services**.

You are responsible for your use of the Services and for any content you submit, post, display or otherwise make available on or through the Service (“User Content”), including that such User Content complies with applicable laws, rules, and regulations. You should only provide Content that you are comfortable sharing with others.

Disqus takes no responsibility and assumes no liability for any User Content that you or any other User or third-party posts or sends over the Service. You shall be solely responsible for your User Content and the consequences of posting or publishing it, and you agree that we are only acting as a passive conduit for your online distribution and publication of your User Content. You understand and agree that you may be exposed to User Content that is inaccurate, objectionable, inappropriate for children, or otherwise unsuited to your purpose, and you agree that Disqus shall not be liable for any damages you allege to incur as a result of User Content.

Any use of or reliance on User Content or materials posted via the Services or obtained by you through the Services is at your own risk. We do not endorse, support, represent or guarantee the completeness, truthfulness, accuracy, or reliability of any User Content or communications posted via the Services or endorse any opinions expressed via the Services. You understand that by using the Services, you may be exposed to User Content that might be offensive, harmful, inaccurate, inappropriate for children or otherwise inappropriate, or in some cases, postings that have been mislabeled or are otherwise deceptive. All User Content is the sole responsibility of the person who originated such User Content. We may not monitor or control the User Content posted via the Services and, we cannot take responsibility for such User Content. You agree that Disqus shall not be liable for any damages you incur as a result of User Content.

Disqus respects the intellectual property rights of others and expects users of the Service to do the same. We reserve the right to remove User Content alleged to be infringing without prior notice, at our sole discretion and without liability to you. We will respond to notices of alleged copyright infringement that comply with applicable law and are properly provided to us as described below.

By using the Service you represent and warrant that your User Content does not violate any applicable law or infringe any third party proprietary rights, including but not limited to, any Intellectual Property Rights.


If you believe that your copyrighted work has been copied in a way that constitutes copyright infringement under the DMCA and is accessible via the Service, please notify Disqus’ copyright agent at the contact information below. For your complaint to be valid under the DMCA, you must provide the following information in writing:

-   An electronic or physical signature of a person authorized to act on behalf of the copyright owner;

-   Identification of the copyrighted work that you claim has been infringed;

-   Identification of the material that is claimed to be infringing and where it is located on the Service;

-   Information reasonably sufficient to permit Disqus to contact you, such as your address, telephone number, and, e-mail address;

-   A statement that you have a good faith belief that use of the material in the manner complained of is not authorized by the copyright owner, its agent, or law; and

-   A statement, made under penalty of perjury, that the above information is accurate, and that you are the copyright owner or are authorized to act on behalf of the owner.

DMCA Agent Contact Information:

Disqus, Inc.

Attn: DMCA Notice Disqus, Inc.

3 Park Ave, 33rd floor
New York, NY 10016

Email: dmca@disqus.com

Please note that this procedure is exclusively for notifying Disqus and its affiliates that your copyrighted material has been infringed. In accordance with the DMCA and other applicable law, Disqus has adopted a policy of terminating, in appropriate circumstances, Users who are deemed to be repeat infringers. Disqus may also at its sole discretion limit access to the Service and/or terminate the accounts of any Users who infringe any intellectual property rights of others, whether or not there is any repeat infringement.

UNDER FEDERAL LAW, IF YOU KNOWINGLY MISREPRESENT THAT ONLINE MATERIAL IS INFRINGING, YOU MAY BE SUBJECT TO CRIMINAL PROSECUTION FOR PERJURY AND CIVIL PENALTIES, INCLUDING MONETARY DAMAGES, COURT COSTS, AND ATTORNEYS’ FEES.

**Rights Regarding User Content**.

You retain your rights to any User Content (“User Content”). By submitting, posting or displaying any Content on the Service, you expressly grant, and you represent and warrant that you have all rights necessary to grant, Disqus a worldwide, royalty-free, non-exclusive, sublicensable, transferable, perpetual and irrevocable license to use, copy, reproduce, process, adapt, modify, publish, transmit, display, distribute, and make derivative works of such User Content in any and all media, technology or distribution methods (now known or later developed). This license authorizes Disqus to make your User Content available, to the rest of the world and to let others do the same. You agree that this license also includes the right for Disqus to provide, promote, and improve the Services and to make User Content submitted to or through the Services available to other companies, organizations or individuals for the syndication, broadcast, distribution, promotion, publication, or otherwise of such User Content on other media and services. Such use by Disqus or other companies, organizations or individuals may be made with no compensation paid to you with respect to your content.

**Disqus Content**

Disqus’ name, logo, designs, trademarks, trade dress, service marks, copyrights, patents or other intellectual property rights in Disqus’ software, images, text, graphics, illustrations, logos, APIs etc. (the “Disqus Content”) is the exclusive property of Disqus or its licensors. Except as explicitly provided herein, nothing in these Terms shall be deemed to create a license in or to Disqus Content, and you agree not to sell, license, rent, modify, distribute, copy, reproduce, transmit, publicly display, publicly perform, publish, adapt, edit or create derivative works from any Disqus Content. Use of the Disqus Content for any purpose not expressly permitted by these Terms is strictly prohibited.

**Feeds and API**

Disqus provides access to portions of its Service via RSS feeds and an API. For the purposes of these Terms, such access constitutes use of the Service. Disqus asks that you use these features respectfully, and as may be outlined in any documentation that we provide. You may not use these or any other features of the Service itself to allow the display of any portion of the Disqus database or reproduce, duplicate or copy any or all of the Disqus Service. Disqus reserves the right to change these features at any time and to disable access to the feeds and the API at any time for any reason or no reason.

**Service Rules**

Please review the Disqus Service Rules below, in consideration of the license to use the Services you agree to comply with the Service Rules which are part of these Terms and outline what is prohibited on the services. Please also note, Disqus comments often appear in websites and online communities not owned by Disqus, these websites and online communities may have their own rules about content and comments on their site, please respect the rules of the communities in which you are using Disqus to comment.

Bullying; Harassment; Hate Speech. We do not allow bullying or hate speech on the Disqus platform. Hate speech attacks people based on “protected characteristics” which include race, ethnicity, sexual orientation, religious affiliation, sex, gender, gender identity or serious disability or disease. Bullying targets individuals with the intention of degrading or shaming them. Bullying is especially harmful to minors because they may be more vulnerable. Disqus prohibits bullying and hate speech and requires our users to respect each other and comment with the respect and sensitivity of others in mind.

Trademark Rights and Rights of Publicity; Impersonation. Users are required to respect the intellectual property rights of others, and are prohibited from posting content that violates someone else’s copyright, trademark, or right of publicity. Additionally, users are prohibited from impersonating others in a manner that does or is intended to mislead or deceive others. Accounts portraying another person in a confusing or deceptive manner may be banned at Disqus’ discretion.

Safety; Self-Harm. Users are prohibited from promoting or encouraging suicide or self-harm. When we receive reports that a person is threatening suicide or self-harm, we may take a number of steps to assist them, such as reaching out to that person and providing resources such as contact information for our mental health partners.

Violence and Criminal Acts. Users are prohibited from promoting or publicizing violent crime, theft, or fraud. We also prohibit users from making credible threats of violence, serious physical harm, or death. This includes, but is not limited to, promoting, publicizing or threatening terrorist activity, organized hate crime, mass or serial murder, human trafficking, organized violence.

Child sexual exploitation. Disqus prohibits content that sexually exploits or endangers children. If we become aware of apparent child exploitation, we will report it in compliance with applicable law.

Inappropriate Content. Graphic media, including explicit violence, gore, and pornographic content are not allowed.

Deceitful data collection; Malware Collecting or harvesting any personally identifiable information, including account names, from the Service; attempting to interfere with, to compromise the system integrity or security or to decipher any transmissions to or from the servers running the Service; (v) taking any action that imposes, or may impose at our sole discretion an unreasonable or disproportionately large load on our infrastructure; (vi) uploading data, viruses, worms, or other software agents through the Service accessing any content on the Service through any technology or means other than those provided or authorized by the Service; or (xiii) bypassing the measures we may use to prevent or restrict access to the Service, including without limitation features that prevent or restrict use or copying of any content or enforce limitations on use of the Service or the content therein.

Spam. Users are prohibited from posting or sending Spam through the service. What constitutes Spam is constantly evolving. Generally, Spam means repeated actions that negatively impact others, such as repeatedly posting a comment with the intent to post a thread etc.

The list of rules above is contently evolving. Disqus may update and revise these rules at any time, please review [Disqus Basic Rules](#basic-rules-for-disqus) for more information.

**Disclaimers and Limitation of Liability**

THE SERVICE IS PROVIDED ON AN “AS IS” AND “AS AVAILABLE” BASIS. YOUR ACCESS TO AND USE OF THE SERVICE IS AT YOUR OWN RISK. WITHOUT LIMITING THE FOREGOING, DISQUS, ITS PARENTS, AFFILIATES, RELATED COMPANIES, OFFICERS, DIRECTORS, EMPLOYEES, AGENTS, REPRESENTATIVES, PARTNERS, AND LICENSORS (THE “DISQUS ENTITIES”) DISCLAIM, TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, ALL WARRANTIES OF ANY KIND, WHETHER EXPRESS OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT. THE DISQUS ENTITIES DO NOT WARRANT OR REPRESENT AND DISCLAIM, TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, ALL LIABILITY FOR: (I) THE COMPLETENESS; ACCURACY, RELIABILITY OR CORRECTNESS OF THE SERVICES OR ANY CONTENT; (II) ANY HARM TO YOUR COMPUTER SYSTEM, LOSS OF DATA, OR OTHER HARM THAT RESULTS FROM YOUR ACCESS TO OR USE OF THE SERVICE OR CONTENT; (III) THE DELETION OF OR FAILURE TO STORE OR TRANSMIT ANY CONTENT AND OTHER COMMUNICATIONS MAINTAINED BY THE SERVICE; (IV) THAT THE SERVICE WILL BE AVAILABLE AT ANY PARTICULAR TIME OR LOCATION, UNINTERRUPTED, SECURE, OR ERROR FREE. ANY CONTENT DOWNLOADED OR OTHERWISE OBTAINED THROUGH THE USE OF THE SERVICE IS DOWNLOADED AT YOUR OWN RISK AND YOU WILL BE SOLELY RESPONSIBLE FOR ANY DAMAGE TO YOUR COMPUTER SYSTEM OR LOSS OF DATA THAT RESULTS FROM SUCH DOWNLOAD OR YOUR USE OF THE SERVICE.

TO THE MAXIMUM EXTENT PERMITTED BY LAW, IN NO EVENT SHALL THE DISQUS ENTITIES, BE LIABLE FOR ANY DIRECT, INDIRECT, PUNITIVE, INCIDENTAL, SPECIAL, CONSEQUENTIAL OR EXEMPLARY DAMAGES, INCLUDING WITHOUT LIMITATION DAMAGES FOR LOSS OF PROFITS, GOODWILL, USE, DATA OR OTHER INTANGIBLE LOSSES, THAT RESULT FROM THE USE OF, OR INABILITY TO USE, THIS SERVICE. UNDER NO CIRCUMSTANCES WILL THE DISQUS ENTITIES BE RESPONSIBLE FOR ANY DAMAGE, LOSS OR INJURY RESULTING FROM HACKING, TAMPERING OR OTHER UNAUTHORIZED ACCESS OR USE OF THE SERVICE OR YOUR ACCOUNT OR THE INFORMATION CONTAINED THEREIN. IN NO EVENT SHALL DISQUS’ CUMULATIVE LIABILITY EXCEED \$1,000 TO YOU AT ANY TIME.

THE SERVICE IS CONTROLLED AND OPERATED FROM ITS FACILITIES IN THE U.S.A. THE DISQUS ENTITIES MAKE NO REPRESENTATIONS THAT THE SERVICE IS APPROPRIATE OR AVAILABLE FOR US IN OTHER LOCATIONS. THOSE WHO ACCESS OR USE THE SERVICE FROM OTHER JURISDICTIONS DO SO AT THEIR OWN VOLITION AND ARE ENTIRELY RESPONSIBLE FOR ANY LIABILITY INCURRED BY DOING SO.

**Disputes, Choice of Law and Jurisdiction**. These Terms will be governed by and construed in accordance with the laws of the State of California, without giving effect to any principles of conflicts of laws. You agree to resolve any claim, dispute, or controversy (“Claims”) arising out of or relating to these Terms or your use of the Service by binding arbitration by the American Arbitration Association (“AAA”) in Santa Clara County, California under the commercial rules then in effect for the AAA, Nothing in this Section shall be deemed as preventing Disqus from seeking injunctive or other equitable relief from the courts as necessary to protect any of Disqus’ proprietary interests.

ALL CLAIMS MUST BE BROUGHT IN THE PARTIES’ INDIVIDUAL CAPACITY, AND NOT AS A PLAINTIFF OR CLASS MEMBER IN ANY PURPORTED CLASS OR REPRESENTATIVE PROCEEDING. YOU AGREE THAT, BY ENTERING INTO THESE TERMS, YOU ARE WAIVING THE RIGHT TO A TRIAL BY JURY OR TO PARTICIPATE IN A CLASS ACTION.

**Any dispute resolution proceedings relating to these Terms or the Site will be conducted only on an individual basis and not as a class, consolidated, joined or representative action and the parties expressly waive all rights to commence or participate in any class, consolidated or representative action/proceeding. You agree that Disqus’ agreement to arbitrate claims constitutes consideration for such waiver.U.S. Government Entities.**

If you are a federal, state, or local government entity in the United States using the Services in your official capacity and legally unable to accept the controlling law, jurisdiction or venue clauses above, then those clauses do not apply to you. For such U.S. federal government entities, these Terms and any action related thereto will be governed by the laws of the United States of America (without reference to conflict of laws) and, in the absence of federal law and to the extent permitted under federal law, the laws of the State of California (excluding choice of law).

**Indemnity**

You agree to defend, indemnify and hold harmless Disqus and its subsidiaries, agents, licensors, managers, and other affiliated companies, and their employees, contractors, agents, officers and directors, from and against any and all claims, damages, obligations, losses, liabilities, costs or debt, and expenses (including but not limited to attorney’s fees) relating to your use of Service or actions taken through the Service, your User Content or any other data or content transmitted or received by you; or your violation of applicable law, third party proprietary rights or these Terms.

**Severability.**

In the event that any provision of these Terms is held to be invalid or unenforceable, then that provision will be limited or eliminated to the minimum extent necessary, and the remaining provisions of these Terms will remain in full force and effect. Disqus’ failure to enforce any right or provision of these Terms will not be deemed a waiver of such right or provision.

Contact Disqus:

3 Park Ave 33rd floor, New York, NY 10016

Click here to view the Privacy Policy.

#### PUBLISHER TERMS OF SERVICE AGREEMENT

**This Publisher Terms of Service Agreement (the, "Agreement") is entered into by and between Disqus, Inc. (“Licensor”) and the publisher ("Publisher") as of the date of executing this Agreement electronically through the Licensor’s website (Effective Date"). Therefore, in consideration of the mutual covenants of the parties and other valuable considerations, the sufficiency and receipt of which is hereby acknowledged, the parties agree as follows:1. Access and Use.**

1.1 *Access*. Licensor hereby grants Publisher a non-exclusive, non-transferable right to access and use Licensor’s software application, application program interface (API), website, and software as a service, (the “Service”) during the Term (as defined below). Publisher may integrate the Service on any web sites owned, operated or controlled by Publisher as set forth in the Service Order, each an “Applicable Site”. Publisher may add Applicable Sites not set forth in the Service Order upon execution of an additional Service Order which shall be governed by this Agreement. Publisher shall not in any way deliver, transfer, or otherwise provide access to or make available the Service to any third parties except as specifically permitted by this Agreement. Publisher is solely responsible for the activity that occurs on Publisher’s account, and is required to keep its account password secure. In the event of any breach of security or unauthorized use of Publisher’s account, Publisher shall notify Licensor immediately. Licensor will not be liable for any losses caused by any unauthorized use of Publisher’s account.

1.2 *Use*. Publisher shall use the Service in accordance with the terms of this Agreement and the Licensor’s privacy policy. Publisher shall be solely responsible for maintaining its own equipment and establishing its own connection via the Internet to the Service. In no event shall Publisher, or any third party, use the Licensor’s APIs to “harvest” or read in bulk the contents of the data files used in the Service, expose or otherwise make available the Licensor’s APIs, including pass-through of the APIs to third parties, nor repackage the APIs to make available their functionality to third parties. Publisher shall not take any action to interfere with the Service or any other user's use of the Service, Licensor’s host or network, including, without limitation, via means of overloading, “flooding”, “mailbombing” or “crashing” the Service.

1.3 *Updates*. The parties agree that Licensor may make updates, modifications or improvements (collectively, “Updates”) to the Service from time to time in its sole discretion.

1.4 *License to Use Service.* Disqus reserves the right to revoke your license to use the Service at any time and for any reason. Disqus may also modify or discontinue the Services or any of its features at any time in our sole discretion without any responsibility or liability to you.

**2. Payments and Fees.** Publisher shall pay Licensor all fees set forth on the Service Order, including any sales, excise, service, use or other taxes now or hereafter imposed upon or required to be collected by Licensor by any authority in connection with this Agreement, excluding taxes based upon Licensor's net income (collectively, the “Fees”). Publisher is solely responsible to ensure that (a) every payment is received by Disqus on time, and (b) all payment information is accurate and up to date. Disqus is not required to inform Publisher about late payments.

2.1 *Paid Subscription.* In the event Publisher elects a paid subscription ("Paid Subscription") for the Service, the Fees for the Service shall be billed in advance monthly and shall be due thirty (30) days from the date of invoice.The first invoicing will occur immediately after execution of this Agreement. Any additional customization or setup fees for additional integration work or work required to add Applicable Sites shall be set forth on a subsequent Service Order Form which shall be governed by the terms of this Agreement. Publisher shall be responsible for interest on all Fees overdue by more than thirty (30) days from the date on the invoice at a rate of the lesser of one and one-half percent (1.5%) per month or the maximum rate allowable by applicable law. Such interest will accrue on a daily basis and be compounded on a monthly basis. Publisher will also be responsible for payment of all reasonable expenses (including reasonable attorneys’ fees and costs) incurred by Licensor in collecting any overdue amounts. Disqus reserves the right to move Publisher from a Paid Subscription to the Ads Version as defined in Section 2.2 if (a) Publisher does not fulfill the payment obligations as set forth in this Agreement and the Service Order and/or (b) Publisher exceeds the eligibility requirements for a given plan under the current pricing plan as set forth in Publisher’s account. For Publisher to receive the Revenue Share as defined in Section 2.2 Publisher will have to comply with the requirements set forth in Section 2.2. Disqus is not required to inform Publisher about these, and any other changes made to Publisher’s account. It is Publishers obligation to verify Publisher’s account settings and to cancel the Ads Version according to this Agreement.

2.2 *Advertising; Revenue Share.* If Publisher has selected a plan that is supported by advertising (“Ads Version”), Publisher agrees that Licensor may include advertisements and/or content provided by Licensor and/or a third party (collectively “Ads”) as part of the Service. Disqus, in its sole discretion, determines whether the Publisher’s Applicable Site(s) are eligible to receive payments for running advertisements ("Revenue Share"). Publisher agrees to comply with any specifications that may be required by Licensor from time to time to enable proper delivery, display, tracking and/or reporting of Ads. As a prerequisite to earning Revenue Share, Publisher must adhere to [Disqus’ Ads.txt policy](#ads-txt-faq), and Publisher shall be required to submit valid payment information and relevant tax forms via Licensor’s publisher dashboard. Licensor shall have no obligation to pay Publisher in the event Licensor has not received payment from its advertisers.Publisher acknowledges and accepts the risk that third parties may generate impressions, clicks or other actions by fraudulent or improper means (“Fraudulent Activity”). Licensor shall have no responsibility or liability to Publisher, and shall have no obligation to pay Publisher, in connection with any Fraudulent Activity. Licensor shall pay Publisher the Revenue Share due to Publisher ninety (90) days from the end of each calendar month that Ads are running on the Applicable Site(s). Payment will be distributed through Tipalti, their Payee Agreement may be found here. Licensor shall not distribute Revenue Share to Publisher if the amount due to Publisher is less than US\$100. Publisher shall be required to claim Revenue Share from Licensor within three (3) months of the date Revenue Share was distributed to Publisher. In the event Publisher does not claim Revenue Share within such time period, Licensor shall have the right to reclaim such Revenue Share. Licensor reserves the right, in its sole discretion, not to run Ads on the Applicable Site(s) for any reason, or no reason, including, but not limited to, quality of the content or content requirements from Licensor’s advertisers. Publishers not eligible for advertising must elect a paid subscription ("Paid Subscription") for the Service, or else service to Applicable Site(s) may be terminated by the Licensor.

**3. Reporting and Audit Rights**. In the event that Publisher has a Paid Subscription to use the Service, the amount of such Paid Subscription is determined based on the Applicable Site(s) page views per month (the “Monthly License Fee”).Publisher shall be required to track and maintain accurate records of the number of average monthly page views per each Applicable Site (“Page Views”) and shall provide such records of Page Views to Licensor after the first 60 days of the Agreement, and thereafter, 15 business days prior to the end of each twelve (12) month period. Licensor shall use such records to prepare the invoice for the following twelve (12) months’ Monthly License Fee in accordance with the fee tiers set forth in the Service Order. Licensor shall have the right, during normal business hours, upon at least five (5) days’ advance written notice to Publisher and no more than twice annually, to audit, examine, inspect, review and make copies or take extracts from, all books and records of Publisher relating to the tracking and reporting of Page Views. If such audit reveals an under-reporting of page views by an amount which would put Publisher in a higher fee tier, than Publisher shall promptly (a) pay to Licensor the difference between the amount paid and the fee tier in which the Publisher should have been; and (b) reimburse Licensor for all reasonable costs incurred by Licensor in performing such audit (including reasonable attorneys’ fees, expenses, and costs).

**4. Data Ownership and Privacy**.

4.1 *Data Ownership.* Licensor shall own all rights, title and interest in and to the comments, content, data and information that is displayed, uploaded, exchanged, transmitted or collected through the Service as provided to the Publisher (the “Disqus Personal Data”). Licensor hereby grants Publisher a limited, non-exclusive and revocable license to use the Disqus Personal Data for comment moderation and analytics purposes only (the “Permitted Purpose”).

4.2 *Data Processing.* For the purposes of this clause, the terms "controller", "data subjects", "personal data", "processor", "processing", and “supervisory authority” shall have the meaning given to them by the European Regulation 2016/679 (“GDPR”). Licensor and Publisher shall be the co-controller of the Disqus Personal Data, and both parties shall process Disqus Personal Data only in accordance with the Permitted Purpose. If Publisher is required to process Disqus Personal Data for any other purpose by a law to which Publisher is subject, (i) Publisher shall inform Licensor of this requirement before the processing, unless that law prohibits this on grounds of public interest, (ii) ensure that its personnel and subcontractors who have access to the Disqus Personal Data have committed themselves to confidentiality and are aware of and comply with Publisher's duties and their personal duties and obligations under this Agreement (iii) implement appropriate technical and organizational security measures to ensure a level of security appropriate to the risks that are presented by the processing of Disqus Personal Data. In case of a personal data breach which affects Disqus Personal Data, Publisher will notify Licensor without undue delay after becoming aware of it, (iv) taking into account the nature of the processing, assist Licensor by appropriate technical and organizational measures insofar as it is possible to fulfill Licensor's obligations to respond to requests from data subjects exercising their rights; (v) taking into account the nature of the processing and the information available to Publisher, assist Licensor, at Licensor's cost, to ensure compliance with the obligations under applicable privacy law with respect to security, breach notifications, impact assessments and consultations with supervisory authorities or regulators; (vi) upon termination of this Agreement or upon Licensor's request, destroy or return all Disqus Personal Data to Licensor (unless a law requires storage of the Disqus Personal Data), and (vii) make available to Licensor all information reasonably necessary to demonstrate compliance with the obligations laid down in this section and allow for and contribute to audits, including inspections, conducted by Licensor or an auditor mandated by Licensor. Licensor acknowledges and agrees that Publisher may retain its affiliates and other third parties as sub-processors (all together "Sub-Processors") in connection with the provision of the Services having imposed on such Sub-Processors the same data protection obligations as are imposed on Publisher under this Agreement. Publisher will be liable to Licensor for the performance of the Sub-Processors' obligations. Publisher will inform Licensor in advance of any changes concerning the addition or replacement of third party processors.

4.3 *Cookies*. Licensor shall be permitted to drop or recognize a cookie on the visitors to the Applicable Sites for the purpose of collecting Disqus Personal Data relating to the visitor’s activity and interaction with the Service, or content on the Applicable Sites, and information about the visitor’s device ID, browser type, environmental or location information, or other similar information, as set forth in the Disqus privacy policy (“Disqus Cookie Data”). To the extent that Cookie Tracking is turned on, and subject to its compliance with applicable Privacy Laws (as defined below), Disqus will also cause third-party cookies to be served. Publishers may choose to turn off Cookie Tracking at any time, however, Publisher shall not be eligible to for Ad Revenue unless Cookie Tracking is turned on. Publisher further agrees that, to the extent Cookie Tracking is turned on, and to the extent required by Privacy Laws, the Applicable Sites contain a mechanism to obtain the user’s consent for the collection of the Disqus Cookie Data for GDPR or other applicable legal purposes and a “Do Not Sell” button for California Consumer Privacy Act of 2018 (“CCPA”) purposes.

4.4 *Compliance with Privacy Laws.* Both Licensor and Publisher shall comply fully with all applicable laws, rules, regulations, and government orders relating to data protection and data privacy, including, but not limited to, the GDPR, the CCPA (collectively “Privacy Laws”), and will only collect, use and disclose Disqus Personal Data collected through the Service and the Applicable Site(s) as set forth in this Agreement and in compliance with applicable Privacy Laws. Publisher will ensure that each of its Applicable Sites contains, a privacy policy that complies with all Privacy Laws and specifically (i) discloses the usage of third-party technology; and to the extent Cookie Tracking is turned on, the data collection and usage by Disqus; and (ii) contains a conspicuous live hyperlink to give users the ability to opt out of interest-based advertising through the Service. Publisher and Licensor agree to comply with the obligations set out in the Standard Contractual Clauses, which are incorporated herein by reference. “Standard Contractual Clauses” means the applicable module(s) of the European Commission’s standard contractual clauses for the transfer of personal data to third countries pursuant to Regulation (EU) 2016/679 of the European Parliament and of the Council, as set out in the Annex to Commission Implementing Decision (EU) 2021/914 (“Standard Contractual Clauses”). The Controller-to-Controller Standard Contractual Clauses shall apply in all cases where Disqus Personal Data that relates to residents of a Restricted Country (as defined below) is processed by Licensor. In particular, and without limiting the above obligations: (i) Publisher and Licensor agree that their respective obligations under the Standard Contractual Clauses shall be governed by the law(s) of the Member State(s) (or Switzerland or the United Kingdom) in which users are established; and (ii) the details of the appendices applicable to the Standard Contractual Clauses are set out in **Exhibit B** to the data processing agreement, which is incorporated herein by reference. “Restricted Country” means a member state of the European Economic Area, Argentina, Brazil, China, Costa Rica, Ghana, Hong Kong, Israel, Malaysia, Mexico, Morocco, Russia, Singapore, Switzerland, Tunisia, Turkey, the United Kingdom, or Uruguay.

**5. Intellectual Property.** Notwithstanding anything to the contrary in this agreement, all intellectual property rights (a) owned or licensed by a party before the date of this agreement and (b) created, developed or licensed by that party after the date of this Agreement independently of this Agreement shall continue to vest in that party or its licensors. Publisher acknowledges that all intellectual property rights in the Service (including any improvements, enhancements and modifications thereto), are Licensor’s Confidential Information (as defined below) and any other software, data, or information provided or made available to Publisher under this Agreement (together the “Licensor’s Intellectual Property”) shall belong to Licensor and Publisher shall have no rights in or to Licensor’s Intellectual Property other than the right to use it in accordance with the terms of this Agreement. Unless otherwise agreed to in writing, Publisher shall not remove or obscure any copyright, trademark or patent notice that appears on the Service.

**6. Confidential Information**

6.1 *Confidential Information.* In connection with this Agreement, each party may disclose, or may learn of or have access to, certain confidential proprietary information owned by the other party (“Confidential Information”). Confidential Information means any non-public data or information, oral or written, that relates to a party, or any of its business activities, technology, developments, inventions, processes, trade secrets, know how, source code, plans, financial information, Publisher and supplier lists, forecasts, and projections. Notwithstanding the foregoing, Confidential Information is deemed not to include information that: (i) is publicly available or in the public domain at the time disclosed; (ii) is or becomes publicly available or enters the public domain through no fault of the receiving party; (iii) is rightfully communicated to the receiving party by persons not bound by confidentiality obligations with respect thereto; (iv) is already in the receiving party's possession free of any confidentiality obligations with respect thereto; (v) can be documented as independently developed by a party without use of any Confidential Information of the other party; or (vi) is approved for release or disclosure by the disclosing party without restriction. Each party shall use reasonable measures to maintain the Confidential Information of the other party in confidence and shall not disclose, publish or copy any part of such Confidential Information, to any third party.Each party shall only use the Confidential Information of the other party for the purpose of this Agreement and shall limit disclosures to any employees on a strict need-to-know basis.Notwithstanding the foregoing, a party may disclose Confidential Information of the other party pursuant to the order or requirement of a court, administrative agency, or other governmental body, provided that such party gives reasonable prior notice (if permissible) to the other party to contest such order or requirement.Upon request, each party shall return to the other party, or certify the destruction of, all Confidential Information of the other party.

**7. Representations and Warranties.**

7.1 *Mutual Representations.* Each party represents and warrants to the other party that: (i) it has the full corporate right, power and authority to enter into this Agreement and to perform the acts required of it hereunder; (ii) the execution of this Agreement and the performance of its obligations hereunder, do not and will not violate any agreement to which it is a party or by which it is bound; and (iii) when executed and delivered, this Agreement will constitute the legal, valid and binding obligation of such party, enforceable against it in accordance with its terms.

7.2. *Licensor Representations.* Licensor makes the following ongoing representations and warranties: (i) that Licensor's software is not contaminated by harmful code (e.g., self-propagating program instructions commonly called viruses or worms); and (ii) that if Licensor's software contains any third party software, Licensor has all rights necessary to license such software.

7.3 *Publisher Representations.* Publisher represents and warrants to Licensor that: (i) it owns, operates, or controls all Applicable Sites; (ii) the Applicable Sites do not contain materials that infringe or violate any third party proprietary rights including, but not limited to, third party intellectual property rights, or materials that violate any applicable laws, rules, or regulations and Privacy Laws; and (iii) the Applicable Sites do not contain any harmful or disabling software code, including without limitation any virus, time-bomb or trojan horse.

7.4 *Disclaimer of Warranties.* except for the express warranties provided for herein, the service, and any support services are provided to Publisher “as is” and Licensor expressly disclaims all warranties, express, implied or statutory, including but not limited to the implied warranties of merchantability, fitness for a particular purpose, and noninfringement, and any warranties arising out of course of dealing, usage, or trade. Licensor does not warrant that the service or any updates will meet Publisher's specific requirements or that the operation of the service or updates will be completely error-free or uninterrupted. Licensor shall not be liable to Publisher for any inoperability of the service or for any loss of information or other injury, damage or disruption of any kind.

**8. Limitation of Liability.** IN NO EVENT WILL EITHER PARTY BE LIABLE TO THE OTHER FOR ANY SPECIAL, INDIRECT, INCIDENTAL OR CONSEQUENTIAL DAMAGES (INCLUDING WITHOUT LIMITATION LOSS OF USE, DATA, BUSINESS OR PROFITS OR COSTS OF COVER) ARISING OUT OF OR IN CONNECTION WITH THIS AGREEMENT OR THE USE OR PERFORMANCE OF THE SERVICE AND/OR UPDATE(S), WHETHER SUCH LIABILITY ARISES FROM ANY CLAIM BASED UPON CONTRACT, WARRANTY, TORT (INCLUDING NEGLIGENCE), PRODUCT LIABILITY OR OTHERWISE, AND WHETHER OR NOT LICENSOR HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH LOSS OR DAMAGE. IN NO EVENT SHALL LICENSOR’S CUMULATIVE LIABILITY TO THE OTHER EXCEED THE FEES PAID TO LICENSOR BY PUBLISHER DURING TWELVE (12) MONTHS PRECEDING THE INCIDENT GIVING RISE TO SUCH LIABILITY.

**9.Indemnification**.

9.1 *Licensor.* Licensor shall indemnify, defend and hold harmless Publisher and its affiliates, and their respective shareholders, officers, directors, employees, agents, successors and assigns from and against any and all third party claims for losses, liabilities, costs, expenses (including amounts paid in settlement and reasonable attorneys’ fees and expenses), penalties, judgments and damages (“Losses”) resulting from any claim by a third party that the Services or infringe or violate the intellectual property rights of any third party, provided, in each case, that Licensee is promptly notified in writing of the claim; (ii) Licensor has sole control of the defense and any negotiations for the settlement of such claim; and (iii) the indemnified party provides to Licensor, at Licensor’s expense, with all reasonable assistance, information, and authority necessary to perform the above.Should the Services in Licensor's opinion, be likely to become, the subject of a claim of infringement, Licensor may, at its option and expense, either procure for Publisher the right to continue using the Services or replace or modify the Services or Work Product in order to make them non-infringing.

9.2 *Publisher.* Publisher agrees to indemnify, defend and hold harmless Licensor, its affiliates and their respective officers, directors, and employees from and against any and all Losses to the extent that such is based upon any third party claim in connection with (i) Publisher’s breach of any of its representations or warranties made hereunder; (ii) Publisher’s violation of any applicable laws, rules or regulations, including, but not limited to, any data protection and data privacy laws and regulations and industry association guidelines; or (iii) Publisher’s violation of any third party intellectual property right.

**10. Term and Termination**

10.1 *Term.* This Agreement shall commence on the Effective Date and shall continue for an initial term of twelve (12) months following the Effective Date (the “Initial Term”). After the expiration of the Initial Term, this Agreement shall automatically renew for additional twelve (12) month periods unless either party gives not less than ninety (90) days’ prior written notice of its intention not to renew (the initial term and any Renewal Term collectively referred to as the “Term”).

10.2 *Termination.* This Agreement shall terminate: (i) by a party thirty (30) business days after the other party’s receipt of written notice that such party is in material breach of any of the terms or conditions set forth in this Agreement, unless such party cures such breach within said thirty (30) business days period or (ii) upon written notice if the other party becomes insolvent, makes a general assignment for the benefit of creditors, files a voluntary petition of bankruptcy, suffers or permits the appointment of a receiver for its business or assets, becomes subject to any proceedings under any bankruptcy or insolvency law, whether domestic or foreign, or has wound up or liquidated its business voluntarily or otherwise, and same has not been discharged or terminated within ninety (90) days. Notwithstanding the foregoing, Licensor may immediately and without prior notice terminate or suspend Publisher’s access to the Service in the event Licensor reasonably believes that continued Publisher access or storage may harm the Service, expose Licensor to liability or is necessary to comply with applicable law.

10.3 *Obligations Upon Termination.* Upon the effective date of expiration or termination of this Agreement for any reason, whether by Publisher or Licensor, Publisher’s right to use the Service shall immediately cease. It is Publisher’s sole responsibility to download Disqus Personal Data; Licensor has no obligation to make any data available to the Publisher following the date of termination. Publisher can request a copy of Disqus Personal Data from Licensor only for additional cost determined by Licensor. Licensor has the right to deny such request at its sole discretion. Promptly upon expiration or termination of this Agreement for any reason, Publisher shall pay any unpaid and outstanding Fees due to Licensor that have accrued as of the date of expiration or termination and Publisher shall return to Licensor, or certify the destruction of, all copies of the Licensor’s Confidential Information.

**11. General Provisions**

11.1 *Severability and Waiver.* If any provision of this Agreement is held to be void, invalid or inoperative, the remaining provisions of this Agreement shall continue in effect and the invalid portion of any provision shall be deemed modified to the least degree necessary to remedy such invalidity while retaining the original intent of the parties.The failure of either party to partially or fully exercise any rights or the waiver of either party of any breach shall not prevent a subsequent exercise of such right or be deemed a waiver of any subsequent breach of the same or any other term of this Agreement.

11.2 *Independent Contractors.* Each party to this Agreement is an independent contractor in relation to the other party with respect to all matters arising under this Agreement. Nothing herein shall be deemed to establish a partnership, joint venture, association or employment relationship between the parties.Publisher may not assign any of its rights or obligations under this Agreement to any other entity without the prior written consent of Licensor, which shall not be unreasonably withheld.

11.3 *Assignment.* Neither party may, or shall have the power to, assign this Agreement without the prior written consent of the other; provided, however, that either party may assign its rights and obligations under this Agreement without the approval of the other party to any subsidiary or Affiliate or successor in connection with a merger, consolidation, sale of all of the equity interests of the party, or a sale of all or substantially all of the assets of the party to which this Agreement relates; provided, that in no event shall such assignment relieve such party of its obligations under this Agreement. Subject to the foregoing, this Agreement shall be binding on the parties hereto and their respective successors and assigns.

11.4 *Entire Agreement.* This Agreement, including any exhibits and schedules attached hereto, constitutes the entire agreement between the parties on this subject matter and supersedes all prior negotiations, understandings and agreements between the parties concerning this subject matter. Neither Party will be bound by, and each party specifically objects to, any term, condition, or other provision which is different from or in addition to the provisions of this Agreement (whether or not it would materially alter this agreement).No amendment or modification of this Agreement shall be made except by a writing signed by both parties.

11.5 *Survival.* The provisions of this Agreement, which by their nature are intended to survive after termination or expiration of this Agreement shall so survive the expiration or termination of this Agreement regardless of the reason or reasons therefore.

11.6 *Freedom of Action.* Either party is free to enter into similar agreements with others and may design, develop, manufacture, acquire or market competitive products or services. Either party may assign and re-assign its employees in any way it may choose and neither party is restricted in any way from hiring or soliciting employees of the other.

11.7 *Counterparts Acceptable.* This Agreement may be executed in any number of counterparts, each of which shall be an original and all of which together shall constitute one and the same document.

11.8 *Publicity.* Licensor shall be entitled, without prior consultation with or approval of the Publisher, to make press releases or other public disclosures with respect to this transaction. Publisher grants Licensor a non-exclusive license during the Term to use its name and trademarks in marketing materials, website or customer lists; provided, that Publisher has the right to notify Licensor in writing if it does not agree to any of the foregoing uses of its name and trademarks.

11.9 *Force Majeure.* Except for payment obligations, neither party shall be in breach of this Agreement or responsible for damages caused by delay or failure to perform, in full or in part, its obligations hereunder, provided that there is due diligence in attempted performance under the circumstances and that such delay or failure is due to fire, earthquake, unusually severe weather, strikes, government sanctioned embargo, flood, act of God, act of war or terrorism, act of any public authority or sovereign government, civil disorder, delay or destruction caused by public carrier, or any other circumstance substantially beyond the control of the party to be charged.

11.10 *Governing Law; Jurisdiction.* The validity, interpretation, performance and enforcement of this Agreement shall be governed by the laws of the State of California and each party irrevocably submits to exclusive jurisdiction and venue in the courts located in Santa Clara County, California. The United Nations Convention on contracts for the International Sales of Goods shall not apply. The remedies under this Agreement shall be cumulative and not alternative and the election of one remedy for a breach shall not preclude pursuit of other remedies unless expressly provided otherwise in this Agreement. Licensor shall be entitled to collect its reasonable attorney’s fees, costs and expenses in any action brought to seek amounts past due or to otherwise enforce rights hereunder.

11.11 *Notice.* All notices and other communications hereunder shall be in writing and shall be deemed to have been duly given when delivered in person (including by overnight courier) or three days after being mailed by registered or certified mail (postage prepaid, return receipt requested) or sent by email, and on the date the notice is sent when sent by verified facsimile or email, in each case to the respective Parties at the address first set forth hereto.

### Trademark Policy {#trademark-policy}

Using a company or business name, logo, or other trademark-protected materials in a manner that may mislead or confuse others with regard to its brand or business affiliation may be considered a trademark policy violation.

#### How does Disqus respond to reported trademark policy violations?

When we receive reports of trademark policy violations from holders of federal or international trademark registrations, we review the account and may take the following actions:

-   When there is a clear intent to mislead others through the unauthorized use of a trademark, Disqus will suspend the account and notify the account holder.

-   When we determine that an account appears to be confusing users, but is not purposefully passing itself off as the trademarked good or service, we give the account holder an opportunity to clear up any potential confusion. We may also release a username for the trademark holder's active use.

-   We are responsive to reports about confusing or misleading Promoted Discovery copy or information. When we receive valid reports, we may give the advertiser an opportunity to clear up any potential confusion. We may also remove specific items from Promoted Discovery, or remove the account from our advertising platform.

#### What is not a trademark policy violation?

Using another's trademark in a way that has nothing to do with the product or service for which the trademark was granted is not a violation of Disqus' trademark policy.

-   Disqus usernames are provided on a first-come, first-served basis and may not be reserved. For information on why you may not be able to select a certain username, please see What is the difference between my Username and my Display Name?

#### Guidelines for fan accounts

Disqus users are allowed to create fan accounts. Disqus provides a platform for its users to share and receive a wide range of ideas and content, and we greatly value and respect our users' expression. Because of these principles, we do not actively monitor users' content and will not edit or remove user content, except in cases of violations of our Terms of Service.

An account's profile information should make it clear that the account is not actually the company or business entity that is the subject of the fan account. Here are some suggestions for distinguishing your account:

-   **Username**: The username should not be the trademarked name of the subject of the fan account.

-   **Name**: The profile name should not be the trademarked name of the company or include the trademarked name in a misleading manner.

-   **Bio**: The bio should include a statement to distinguish it from the real company, such as “Unofficial Account," "Fan Account," or "Not affiliated with..."

-   **Profile photo, header photo, or background image**: The account should not use another’s trademark, logo or other copyright-protected image without express permission.

-   **Communication with other users**: The account should not, through private or public communication with other users, try to deceive or mislead others about its identity.

Users may also choose to use different language to indicate that an account is not associated with the actual brand/company/product so long as it is clear and not confusing to others, and does not mislead or deceive.

If an account is reported to be confusing, we may request that the account holder make further changes to bring the account in compliance with these best practices.

#### How can I make my own account's brand or business affiliation clear?

We strongly recommend that you use all of Disqus' account settings (account name, location, web site, and bio) to make your account's affiliation clear.

-   Please see the profile and avatar sections of our Updating your Disqus Settings documentation for instructions on customizing your account. In particular, we recommend clearly stating your location, including your website if you have one, and clearly describing your brand or business in the bio, if applicable.

#### What information is required when reporting trademark policy violations?

In order to investigate trademark policy violations, please provide all of the following information:

-   Username of the reported account (e.g., cocacola or //disqus.com/cocacola):

-   Your company name:

-   Your company Disqus account (if there is one):

-   Company website:

-   Your trademarked word, symbol, etc. (e.g., Coca Cola):

-   Trademark registration number:

-   Trademark registration office (e.g., USPTO):

Note: A federal or international trademark registration number is required. If the name you are reporting is not a registered mark (e.g., a government agency or non-profit organization), please let us know:

-   Your first and last name:

-   Title:

-   Address:

-   Phone:

-   Fax:

-   Email (must be from company domain):

-   Description of confusion (e.g., passing off as your company, including specific descriptions of content or behavior):

-   Requested Action (e.g., removal of violating account or transfer of trademarked username to an existing company account):

#### How do I report a trademark policy violation?

You do not need a Disqus account to submit a trademark report. Holders of registered trademarks can report possible violations to Disqus through our Support Form.

Please submit trademark-related requests from your company email address and follow the format above to help expedite our response. Also, be sure to clearly describe to us why the account or comments posted by it may cause confusion with your mark.

## Troubleshooting {#cat-troubleshooting}

### Adding Disqus to static Wordpress Pages {#adding-disqus-to-static-wordpress-pages}

When you integrate Disqus into your Wordpress site, you may notice that Disqus only appears on the Wordpress "Posts", such as your blog posts, and does not appear on standalone Pages, such as your site's homepage.
​
This is because Disqus is designed to replace the native Wordpress comments form, and will only appear where the Wordpress comments form is present. For any posts where you'd like to hide or prevent comments, this will allow you to disable comments within Wordpress for that post, and Disqus will be hidden. However, as Wordpress comments are only enabled for Posts, this will also prevent Disqus from appearing on any Pages by default.
​
To enable comments on static WordPress Pages, you'll want to add the comments_template() code below to your page template wherever you want the comment form to appear:

0
​

Additional information from the WordPress dev team may be found here: 0

### Blogger Troubleshooting {#blogger-troubleshooting}

Most often, the best default position for the Disqus widget is in the bottom slot of the bottom-most right column.

#### Comment counts are missing

If comment counts were appearing on your blog but have disappeared, most likely your Blogger Layout > Blog Posts widget has become corrupted. To fix this, try the following solutions, in order:

#### Solution 1: Reset the corrupted Blog posts widget

This can be fixed by following Blogger's How to reset corrupted Blogger Blog Posts template guide.

#### Solution 2: Revert all widget templates to default

If the above guide doesn't work (e.g., if it errors out or if the comment counts still don't appear) you'll need to revert your widget templates to their default status. To do so:

1.  Go to Blogger > Template > Edit HTML tab.

2.  Click the "Revert widget templates to default" link.

3.  Click OK in the confirmation message.

#### Synced Comments Showing Site Owner Name Instead of Commenter Name

If your blog's settings allow only certain people to comment within Blogger, all comments synced from Disqus will be shown as authored by the blog owner instead of the commenter.
To prevent this, go to **Blogger → Settings → Comments** and set the **Who Can Comment?** option to Anyone.

#### The Disqus gadget installer isn't working for my blog

If the gadget installer isn't working for your site, you have the option of manually installing Disqus. If you're using a standard Blogger template, see Manually adding a Disqus gadget to Blogger

#### Note that Blogger's new Dynamic Views templates don't support gadgets or custom HTML, so these templates can't use Disqus on them.

#### I've installed Disqus but Google+ comments are still showing

When Google+ comments are enabled, Disqus cannot load. You will need to disable G+ comments before Disqus will appear in Blogger – from your Blogger dashboard, click the "Google+" tab and unselect the "Use Google+ Comments" option.

### I'm receiving the message "We were unable to load Disqus." {#i-m-receiving-the-message-we-were-unable-to-load-disqus}

***Check our**status page**for updates on system-wide issues.***

When Disqus is loaded on a page for the first time, our servers check the URL on which Disqus is loaded (and the 0 configuration variable if it is set) to make sure it's valid and meets certain criteria. When this fails, we show the message *“We were unable to load Disqus. ...”*

There are several reasons why you may be prompted with this message:

Trusted domains are sites on which you want your forum shortname to load, and are set within your admin panel. If Disqus isn't loading on your site, make sure your site's domain is in the trusted domains list and is in the proper format.

CORRECT formats:

0
0

INCORRECT formats:

0
0

#### Your shortname is missing or incorrect

Disqus won't load if you haven't yet registered a forum shortname or if the shortname you have entered is incorrect (See What's a shortname?). We suggest double-checking your shortname in the Disqus embed code, which should appear in place of EXAMPLE in this line: 0 of your code, or the settings page for your respective platform.

#### Disqus is being loaded on a different domain than you registered

By default Disqus is only allowed to load on the domain specified when you originally registered. This is enforced via the Trusted Domains setting. You can add, remove, or change trusted domains at the Disqus admin > Settings > Advanced page.

#### Recent webhost or domain name change

Allow 48 hours for your new DNS settings to propagate when switching hosts or domains. Even though your site's content is visible, our servers can't connect to your pages in the meantime.

#### Timeout

Our servers must reach your site within ten seconds.

#### HTTP status error codes, e.g., 404 Not Found

HTTP status error codes like **404 Not found** and **503 Service unavailable** can be returned to our servers by your pages even when your page's contents are visible.

Try contacting your host and let them know the status code your pages are returning. You can check header status codes for any page at the WebConfs Header Status Code Checker Tool.

#### Your page URL or title contains non-ASCII characters

For a URL to verify, it cannot contain non-ASCII "special" characters (e.g., ñ, å, š) which are usually exclusive to one or more languages. Full list of supported ASCII characters.

To fix this, set the 0 or 1 configuration variable (whichever is appropriate) and make sure to convert non-ASCII characters to ASCII characters in the variable (e.g., å to a, š to s). This will allow your site's visitors to see the proper non-ASCII version and our system can load a thread for the page.

#### Incorrectly-formatted JavaScript configuration variables

-   0 cannot be longer than 200 characters.

-   0 cannot contain spaces.

-   0 must use an absolute URL; relative URLs won’t work. E.g.,
    Good — absolute URL: 0
    Bad — relative URL: 0

#### Further assistance

If you cannot find your answer, you can contact our support team.

### Installation Troubleshooting {#installation-troubleshooting}

If Disqus isn't loading on your site after you've installed it, there are a number of things you can check to make sure everything is set up correctly. This guide is intended to help you walk through common pitfalls when installing Disqus.

#### Did you follow the official install instructions?

Disqus is easy to install and we provide the most up-to-date installation instructions and plugins on this website. The quickest way to check over your installation is to review the installation instructions and make sure you've followed the instructions correctly.

Also, check out the Quickstart Guide for an overview.

#### Learn more about how Disqus works

For a slightly more technical description of Disqus, read How does Disqus work?.

#### Common troubleshooting

#### Are you using the right account information?

The Disqus JavaScript is specific to your account on Disqus, which is called your forum. Make sure you haven't mistyped your forum shortname.

#### Are you using the trusted domains feature in your settings?

Trusted domains are sites which you want your forum shortname to load on. If Disqus isn't loading on your site, make sure your site's domain is in the trusted domains list, and is in the proper format.

CORRECT formats:

0
0

INCORRECT formats:

0
0

#### **Are your thread URLs complete?**

When inserting the thread URLs in your code, you'll want to ensure that the domain is present, and that you're not using a relative URL. When only the tail end of a URL is present, Disqus will not be able to load.
​
For example, having the following value inserted as the thread URL will prevent Disqus from loading correctly:
0'

To correct this, ensure that the complete URL value is inserted, like below:

'0'

If your protocol is relative, you can leave off 0, but the 1 section of the URL must be present for Disqus to load.

#### Unable to complete installation?

Contact Disqus Support

### Introducing the Discussions Editor and FAQ {#introducing-the-discussions-editor-and-faq}

The Discussions Editor allows you to 1-click update attributes of any discussion on your site. For example, you might update the title or link associated with a discussion after updating it on your site itself.

Simply click into any cell, enter the desired new information, and hit enter or click out of the cell. The attribute will be automatically updated and you'll see a success or error message. These attributes can also be updated using the threads/update API call.

Note:

-   Discussions cannot currently be merged via this interface. For example, updating one discussion's link to be the same as another's will result in no change. To merge discussions, see our URL mapper documentation.

-   Updating an author currently requires entering that Disqus user's username, not full name. For more in the difference between the two see [What is a username?](#what-is-the-difference-between-my-username-and-my-display-na). Thus, only change the author if you're confident of that person's username.
    ​

#### What attributes can be updated?

Currently the following discussion attributes are included for updating:

-   **Title**: appears publicly. See "Where does this information appear publicly?" below.

-   **Link**: appears publicly. See "Where does this information appear publicly?" below.

-   **Author**: does not appear publicly. Update this when you want a non-moderator on your Disqus forum to be considered the author for the discussion. This allows them to moderate comments on that discussion both via Disqus notification emails and inline in the commenting embed.

-   **Category**: does not appear publicly. Categories are primarily used with our API for results filtering; categories are not used for moderation. Learn more about categories.

-   **Open/closed status**: open or close a discussion to new comments.

Discussion creation date is also included in its own column as a point of reference.
​

#### Where does this information appear publicly?

Attributes like discussion titles and links show up in numerous places around Disqus, including:

-   Notification emails

-   [Recommendations](#recommendations)

-   Disqus Home

And more places as we add new features and products...

### Troubleshooting 101 {#troubleshooting-101}

-   What is a forum

-   Can’t access the moderation panel

-   How do I check which forum is installed on my site?

#### Installation Troubleshooting:

-   Common installation problems

-   Getting the "We are unable to load Disqus" error message

-   Wordpress troubleshooting and FAQ

-   Blogger troubleshooting

-   Tumblr troubleshooting

-   Why isn’t the comment box loading?

-   Why are comments being posted to Blogger instead of Disqus?

#### Thread Troubleshooting:

-   Why are wrong URLs detected for my discussions?

-   Why are the same comments showing up on multiple pages?

-   Comments missing or are threads splitting?

-   Troubleshooting imports

-   Why are comments posted to other sites showing up in my admin?

-   Why are comments visible in the Disqus admin but not on my site?

#### Misc Troubleshooting:

-   What does the error “Disqus seems to be taking longer than usual to load” indicate?

-   Why is Disqus loading slowly when I do site performance tests?

-   Troubleshooting Disqus in Internet Explorer 8/9/10

-   Why isn’t Google indexing my comments?

-   Troubleshooting common error messages

### Troubleshooting Common Error Messages {#troubleshooting-common-error-messages}

This error occurs when your browser isn't sending a valid HTTP referer. The possible reasons you may not be sending referers are if:

-   Referers have been manually disabled in your browser

-   You're connecting through a proxy server

-   You have a browser plugin/extension which disables referrers

#### Disqus seems to be taking longer than usual

This error message occurs when Disqus has taken more than 10 seconds to load. This can be caused by several different factors:

-   slow browser plugins

-   your internet connection speed

-   additional scripts loading on the page

-   speed of the server hosting the website in question

#### You are not allowed to access this page. Please make sure you're logged into the correct account

This error message indicates that you are not logged in to the moderator account for the current forum you are trying to access. This may happen when you click through to 0 from WordPress. Try the following troubleshooting steps:

-   Log out of Disqus and Log in to the moderator account for that forum

-   Reset your Disqus password for any possible email addresses for you used to create the Disqus forum.

### Troubleshooting Disqus in Internet Explorer 8/9/10 {#troubleshooting-disqus-in-internet-explorer-8-9-10}

Disqus officially supports Internet Explorer version 10, though it may also work in Internet Explorer versions 8 and 9. If Disqus is loading inconsistently or not at all for you in these browsers, there are several factors that may cause this:

#### Compatibility View

When Internet Explorer's Compatibility View is enabled, you may see the following message in place of the normal Disqus commenting area:

*"Your browser is not currently supported in Disqus. Please use a modern browser."*

This is because Internet Explorer is loading the page using an older, "compatible" version of Internet Explorer's rendering engine.

To fix this, **turn off Compatibility View**. See Microsoft's Compatibility View help page for steps on how to do this.

Web developers also have the option of setting their own document mode using the following:

    meta http-equiv="X-UA-Compatible" content="IE=8"

#### Quirks Mode

Internet Explorer will render in Quirks Mode when the <!DOCTYPE> is either missing or doesn't comply with standards. This may result in strange Disqus loading behavior; for example, Disqus may load wider than the page, forcing a horizontal scrollbar.

To check if Internet Explorer is using Quirks Mode, open the development tools (press F12) > check the Document Mode. If it says Quirks, you can try changing this to IE8 or IE9 Standards to view the page normally.

#### For Publishers

Pages being loaded in Quirks Mode can be fixed by using a standard <!DOCTYPE>. The most common causes include:

#### No doctype present

A doctype is required to comply with rendering standards. Not including one will cause Internet Explorer to render in Quirks Mode. You can find a list of standard doctypes here.

#### Whitespace or other characters before the doctype

No characters (including whitespace) can precede the doctype. Make sure it's the very first thing listed on the page.

#### Missing public identifier

The public identifier is the URL that comes after the doctype, and is required to comply with standards.

This doctype is **incorrect** because it's missing the identifier URL:

    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

This would be the **correct** version with the identifier URL:

    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "0">

#### IE7 Emulation

Some websites force Internet Explorer to load emulating an older version of Internet Explorer. You can check by opening the development tools (Press F12) and checking the Document Mode. If it says **Internet Explorer 7 Standards**, you can switch this to **Internet Explorer 8/9 Standards** to load Disqus properly.

#### For Publishers

The most common reason the page would load in IE7 standards is because of the *EmulateIE7* meta tags. Look in your page source and find an HTML tag similar to this:

    <meta content='IE=EmulateIE7' http-equiv='X-UA-Compatible '/>

You can either remove this or change it to use the latest version of IE available:

    <meta http-equiv="X-UA-Compatible" content="IE=9; IE=8; IE=7; IE=EDGE; chrome=1" />

For instructions using Blogger, see: Updating a Blogger template to support all versions of Internet Explorer

### Troubleshooting Imports {#troubleshooting-imports}

Additional information on imports for your account can be found on the Import and Exports page.

Note that imports don't start instantly but are placed into a queue. Imports queued normally complete within 24 hours.

If you're importing your comments into a different domain name, you'll need to migrate the thread URLs of the comments imported.

#### Missing required post key

These post keys are required in order to generate a thread within Disqus. Replace these missing post keys in order to resolve this error.

#### Skipped (check that the link exists)

When you perform the import process, our servers try to load your threads. If your website returns an error (e.g. 404 or 500), our servers will be unable to find your threads. This can happen on websites that use complex redirects, or caching plugins (e.g. the WP Super Cache or W3 Total Cache plugins for WordPress).

#### General Troubleshooting

#### For errors without a description

Run your import through W3C's XML validator. Correct any additional errors and requeue your import from within the Disqus admin.

#### For Self-Hosted WordPress Installations

1.  We recommend disabling any caching plugins installed on your site e.g. WP Super Cache or W3 Total Cache.

2.  Wait an hour for your cache to refresh.

3.  Manually export your blog at your WP Dashboard > Tools > Export.

4.  Manually import your blog at our Import/Export page.

5.  Re-enable the plugins you disabled one by one until you find the culprit.

6.  Shoot us an email with your debug information and the plugin that caused the exporting trouble at help+wp@disqus.com.

*Note: If these steps don't work, try disabling all WordPress plugins during Step 1 in addition to just the caching plugins.*

**FYI:** When importing from Wordpress into Disqus, source URLs need to be available and on a static domain.

### Tumblr Troubleshooting {#tumblr-troubleshooting}

-   First you'll want to verify that you've entered the correct shortname within the Appearance settings. Note that a space or uppercase letters will conflict with Disqus functioning properly.

-   If your theme doesn't have a shortname setting or has been coded incorrectly, you'll want to using our Manual Installation Instructions.

-   If you're already using the manual installation instructions, please be sure to remove the "<" and ">" from either side of your shortname.

#### Customization

These customizations require an advanced understanding of HTML and may differ among Tumblr themes.

#### Hide Disqus on static pages

    {block:Posts}{block:Permalink}{block:Date}
          {block:IfDisqusShortname}
            ... Disqus embed code ...
          {/block:IfDisqusShortname}
    {/block:Date}{/block:Permalink}{/block:Posts}

#### Disqus not appearing on your mobile device?

Disqus doesn't load in Tumblr's optimized mobile layout. You can disable the mobile layout by unchecking Customize > Advanced > "Use optimized layout on mobile devices".

### Use Configuration Variables to Avoid Split Threads and Missing Comments {#use-configuration-variables-to-avoid-split-threads-and-missi}

When 0 or 1 are not defined, the Disqus embed will use the window URL as the main identifier when creating a thread. In other words, each unique URL Disqus loads on will result in a new unique thread. This works well for some sites, however, this method of creating threads can lead to duplicate “split threads” for the same page of content, especially when your site accepts many different URLs for the same page of content.

For example, one post on your blog may be accessed via 0 and 1. See Google’s Use Canonical URLs for other common reasons why a site may have many URLs for the same content.

When your site creates many “duplicate” Disqus threads, this can lead to the following unwanted results:

-   Comments in your Disqus forum may not display on the expected thread on your site, resulting in “missing comments”.

-   Engage email notifications may contain your site’s non original (a.k.a “noncanonical”) URLs that get seen by commenters and moderators.

-   Users may comment on a split thread resulting in multiple discussions that later need to be merged using our migration tools.

-   Your thread URL data that’s used on Disqus.com and in other Disqus features may be inaccurate.

#### Define url and identifier to Avoid Thread Splitting

We recommend that you define the 0 and 1 configuration variables within the 2 section of your embed code. See the installation instructions for the complete embed code.
​
Example:

0
0
0
0

These variables should be defined by using your CMS’ or platform’s dynamic values, if available. For example, the Disqus WordPress plugin defines these variables by default using the following PHP values:

0
0

See our full documentation on Configuration Variables for more information.

-   Why are comments visible in the Disqus admin but not on my site?

-   Why are the wrong URLs detected for my discussions?

-   How can I update discussion URLs?

-   How do I load the same thread of comments on multiple pages?

-   Migration Tools

### Why are comments being posted to Blogger instead of Disqus? {#why-are-comments-being-posted-to-blogger-instead-of-disqus}

Here's what's happening: Blogger mobile templates don’t currently support widgets, so the Disqus widget doesn’t replace Blogger’s default comments system when your site is viewed via a mobile device. Essentially, when people view your site using a mobile device their comments bypass Disqus and are posted straight to Blogger .

Until Blogger mobile templates do support widgets, you can fix this by either:

1.  Disabling Blogger mobile templates; or

2.  Disabling commenting in Blogger — this will not disable Disqus comments, it will only disable commenting on your mobile template; or

3.  Manually enable Disqus in Blogger mobile templates (see below).

#### How to manually enable Disqus in Blogger mobile templates

Note: This requires manually editing your Blogger template and HTML. This is meant for advanced Blogger template developers only. We also recommend backing up your template first.

1.  In Blogger visit your site's settings > Template.

2.  Choose "Edit HTML".

3.  Choose "Proceed".

4.  Enable the "Expand Widget Templates" checkbox.

5.  Search for "disqus".

6.  Add mobile='yes' to the following:

7.  Choose "Save template".

The result should look like:

    <b:widget
        id='HTML3'
        locked='false'
        mobile='yes'
        title='Sample site'
        type='HTML'
    >

### Why are comments posted to other sites showing up in my Disqus admin? {#why-are-comments-posted-to-other-sites-showing-up-in-my-disq}

The most common reason for comments from other sites showing up in your Disqus admin is because you're using a Disqus shortname owned by another site. This can be resolved by verifying the shortname installed on your site.

#### Verifying the shortname on your site

See the "How do I check which forum is installed on my site?" section at What is a forum?

#### Another site is using your Disqus shortname

Alternatively, if you've verified you're using a Disqus shortname you own, it's possible another site is using your Disqus shortname. You can prevent sites from being able to post comments to your Disqus shortname by adding your site's domain to your Trusted Domains list, available in the Disqus Admin > Setup > Advanced

### Why are comments visible in the Disqus admin but not on my site? {#why-are-comments-visible-in-the-disqus-admin-but-not-on-my-s}

Whenever comments appear in the Disqus admin but not on the relevant page on your site, this means there are actually two unique discussion threads in Disqus. One thread (with 0 comments) appears on your site, the other thread (with all missing comments) appears in the admin.

Most often this results after changing either your website's URL configuration or the Disqus configuration on your site. For example, this can happen after:

-   **Moving domains**, e.g., 0 to 1

-   **Changing permalink structures**, e.g., 0 to 1

-   **Changing permalink styles** such as camel-style slugs to all-lowercase slugs, e.g., 0 to 1

To fix this issue, merge the affected threads using one of our migration tools.

### Why are the same comments showing up on multiple pages? {#why-are-the-same-comments-showing-up-on-multiple-pages}

When the same comment thread appears on multiple pages, this is normally due to multiple thread identifiers being assigned to the same thread whenever the thread was initially created.

#### Resolving identifier conflict

Identifier conflict can be resolved by verifying that a unique URL is being set and then assigning a unique identifier. This will create a new thread on pages which were previously displaying the same thread. Note that thread identifiers and URLs are set using our embed Javascript Variables.

You can also check if identifiers are being assigned to the same thread using our API console.

#### Prevention

To prevent this from happening, you'll need to ensure that both the **identifier** and **thread** are unique when creating new threads.

The first time Disqus is loaded on a page, our system creates a new thread ID and associates that thread ID with the page. Unfortunately, if multiple pages are loaded from the same URL that means each page will be associated with the same thread ID.
​
This is known as **identifier conflict** because, even though each page will have a unique Disqus identifier, the pages share the same thread ID. Identifier conflict causes each of these pages to show the same comments.

#### Example situations known to cause identifier conflict:

-   **Setting the same0or1variable**

-   **Previewing pages** before publishing: Disqus will associate that preview URL with the thread ID.

-   **Copying pages**: If Disqus is loaded on the newly copied page before changing the JavaScript configuration variables.

If you're experiencing identifier conflict on your site, you'll need to replace the disqus_identifiers on the pages where this is occurring with ones that are unique. We'd also recommend double-checking your publishing process since this commonly occurs when multiple identifiers are assigned to post preview URLs. For this reason, Disqus is disabled within the WordPress preview mode.

-   What is a Disqus identifier?

### Why are the wrong URLs detected for my discussions? {#why-are-the-wrong-urls-detected-for-my-discussions}

The first time Disqus is loaded on a page, our system creates a new discussion thread and associates that discussion with the URL of the page used during that first load.

In more advanced implementations this may not be desirable; we've provided below a few example situations and solutions.

#### Previewing pages before publishing

In this scenario, Disqus will associate that preview URL with the discussion thread. For example, say a publishing environment uses 0 for its public-facing site and 1 for previewing pages before publishing. If Disqus is loaded in the preview environment, discussion threads will be associated with the 2 domain.

To fix already-affected discussions, see How can I update discussion URLs?

To prevent this from happening further, either:

1.  Prevent Disqus from loading in the preview environment; or,

2.  Set the disqus_url variable upon thread creation.

If your CMS is Hubspot, use the following HubL to prevent Disqus from loading on preview pages.

    {% unless request.path_and_query is string_containing "?hs_preview=" %}

        {% unless request.full_url is string_containing "preview.hs-sites.com" %}

         <div id="disqus_thread"></div>

        {% endunless %}

    {% endunless %}

#### Setting the wrong disqus_url variable

To fix already-affected discussions, see How can I update discussion URLs?

Thread URLs cannot be updated by passing \0 after a thread has been created. \1 can only be set once, upon thread creation.

### Why isn't the comment box loading? {#why-isn-t-the-comment-box-loading}

While Disqus works most browsers, not all browsers and versions are officially supported and this can affect the Disqus experience. Here's a resource to see which browsers Disqus works with.
​

If your browser is supported, you should try the following solutions in order:

#### Solution 1

Clear both your browser's cache and cookies - How do I clear my cache and cookies?

#### Solution 2

Temporarily disable all plugins, extensions, and add-ons in your browser.

Tip: If you use Firefox or Chrome, you can use Firefox’s Safe Mode or Chrome’s Incognito mode instead.

**See our list of knownBrowser plugin/extension conflicts.**

#### Still not working?

Contact the moderator(s) of the site in question.

#### "Disqus could not be loaded because it is not being loaded from a trusted domain."

Contact the moderator(s) of the site in question.

#### "We were unable to load Disqus. For more information please see our documentation on identifier and urls"

Contact the moderator(s) of the site in question.

For more information, see our JavaScript configuration variables documentation.

#### Broken images or links

Contact us with the following:

-   Name of your ISP (Internet Service Provider), e.g., Comcast

-   Screenshot of what you’re seeing - How do I take a screenshot?

-   Link to specific page where you’re seeing this

### WordPress Troubleshooting and FAQ {#wordpress-troubleshooting-and-faq}

Keep in mind that the Disqus WordPress (WP) plugin is designed to replace your WP comment system wherever it is referenced within your templates. Any actions such as closing comments or removing references to the comment template within WP could result in Disqus not displaying properly.

#### Troubleshooting steps

First, you'd want to ensure that comments are enabled on your posts.

1.  Check that "Allow comments" is checked on all posts you expect Disqus to display. You can verify this by going to 0 on your either of your posts.

2.  See if your Wordpress discussions are not closed, as that will also prevent Disqus from loading. You can do this by going to the discussions tab in your WP admin.

Most WordPress irregularities are caused by theme or plugin conflicts. To test, we'd recommend the following:

1.  Reverting to a default WordPress theme

2.  Disabling all plugins

#### Common conflicting themes/plugins:

-   Thesis 2.0 (theme) - Resolution for this conflict

-   Twenty Thirteen (theme) - Resolution for this conflict

-   MailChimp's Social Plugin

-   Facebook Connect

-   Cloudflare RocketLoader

-   GeoDirectory (plugin) - Resolution for this conflict

#### Possible conflicting themes/plugins:

-   HTTP / HTTPS Remover - (To resolve, use Make Paths Relative instead)

-   Spam Free Wordpress plugin

-   AWeber

-   Social

-   Root Relative URLs

-   Google Friend Connect Plugin

-   WPCache

-   iThemes Security Plugin

-   Headway (theme) - Resolution for this conflict

-   Hatch (theme) - Have a fix for this? Let us know!

-   Hybrid 1.2 (theme) - Have a fix for this? Let us know!

Additional conflicts could occur with any plugins that shorten links or optimize/accelerate pages, as these can edit or confine the scripts needed for Disqus to run.
​

If none of the above conflicts are causing the issue, it may be a conflict in your browser, please see Browser plugin/extension conflicts for more information.

#### If you've encountered an error message

Visit the Error Message FAQ to see if the error is listed.

#### If Syncing is not working in version 3.x

*Take a look at the Troubleshooting section in our Syncing with Wordpress article.*

#### WordPress F.A.Q.

#### Can Disqus be used on WordPress.com?

Using Disqus on a 0 site requires a WordPress.com Business plan which supports third-party plugins like Disqus. By contrast, a site set up using Wordpress.org will be able to use the Disqus plugin without the WordPress.com Business plan.

#### Is Disqus free to use on my site?

Yes! Disqus is free to use. We also provide subscription plans for larger, commercial sites that want access to more powerful moderation and audience tools and customization.

#### How do I customize the look-and-feel of Disqus?

Disqus automatically checks your site's font and background color and adapts to either a light or dark color scheme, along with a serif or sans-serif font. If these are detected incorrectly, you can override them in your Settings.

#### Will I lose comments if I deactivate Disqus?

The Disqus for WordPress plugin supports syncing comments back to your WordPress database. These comments will be available in WordPress should Disqus be deactivated or removed from your site. You can also export your comments from Disqus at any time.

#### Can I import my existing WordPress comments into Disqus?

Yes, you can import your existing WordPress comments into Disqus during installation.

#### How do I set up Single Sign-On (SSO)?

SSO allows users in your database to comment without requiring them to register with Disqus. Access to SSO is currently available as an add-on for users with a Business subscription. Also check out our guide for setting up SSO on WordPress.

#### How can I install an older version of the plugin?

To install an older version (before the major 3.0+ update), go to 0 to download version 2.6. For more information on how to install, see WordPress' Managing Plugins. Use this version only if you have trouble with the current plugin.

## User Profile {#cat-user-profile}

### Improving Your User Profile {#improving-your-user-profile}

Your Disqus profile is the hub of all of your comment history. From here, you can see comments, frequented communities, discussions, recommends, and followers of yours or a user who's profile you may be viewing. This profile makes it easy to manage your Disqus-self, in addition to aiding commenters chatting with one another by seeing where someone has commented recently.

#### Filling out Your Profile

Completing your Disqus profile is not as time consuming as it would appear. Ensuring that you fill out your bio, username, display name, and avatar can pay off in the short time it takes to enter this information. This allows your community, or other communities you comment in, get to know a little bit about you and increases the chance that users would be interested in what you say and do.
​
**Bonus Tip:** Other commenters are more likely to interact with a user that has a custom username, display name, and avatar, than they are with a user without any of that information filled out.

#### Where can I find my total number of positive votes?

The upvote count is displayed for each user directly in the profile bio.

Total number of upvotes and comments can also be seen by hovering your mouse over an avatar next to any comment.

#### Why would I want to follow another user?

Many comments posted through Disqus are hilarious, insightful, or just fascinating to read. Following other users keeps you up-to-date with the latest news that's relevant to you. After you choose to follow someone, his or her activity will show up in your Home feeds and Daily Digest Email.
​
Following users is a great way to discover new communities that you never knew existed. Everyone has different interests so following some of your favorite commenters can be a great way to find new and interesting communities.
​

#### Where can I find interesting people to follow?

To get you started, head over to Home and follow some users that are already using Disqus. You can follow other Disqus users the way you normally would from the discussion or on someone's profile, but we've also added some new ways to follow people.
​
If someone starts following you, or they reply to one of your comments, the you can follow them from within your notification feed by clicking the follow button next to their name.
​

#### Can I hide my profile activity from being shown?

Yes! You can make your profile activity private which will also disable following for your account. Your avatar and display name will still be public. While you can make your profile activity private, your comments will still be public in the original discussion thread, and may show up in search engines, for example.

### Log into Disqus with a Social Media Account {#log-into-disqus-with-a-social-media-account}

Our login process ensures that all social login users (Google, Facebook, and Twitter) set an email and password for their account. Although having a Disqus account is required to comment using Disqus (unless guest commenting is enabled on the site where you want to comment), users can still login with the following social logins: Google, Facebook, and Twitter. Users will only need to authenticate with Disqus one time.

#### Login with a social media service:

From the Disqus embed, click the icon for your preferred social media network or click its name in the dropdown. Authenticate with the social media network using your credentials for that site.

#### You'll now be asked to create a new Disqus account or link an existing Disqus account.

If you're having trouble logging into an existing Disqus account, here are some [Login Troubleshooting](#logging-into-disqus) steps to try.

#### What are the Benefits?

-   Login with your preferred method yet manage a single account.

-   You’ll be sure to have a password linked with your account, making your account more secure.

-   Ensure that an email address is linked with your account. This keeps you connected via notifications and the Disqus community through Disqus Digests and other important announcements.

*To adjust notifications, go to* *0, then click the “Notifications” tab.*

**Where did my comments go?!** Accounts created via a Social Media log in will exist separately from a full account registered with Disqus by email address. If you intended to log into another account you already had registered, try logging in on 0 with your email address.

-   Connect Facebook, Twitter, and Google

### Logging into Disqus {#logging-into-disqus}

Registered commenters can log into their account several different ways using their registered **username** or **email address** and password. Instructions for creating a new commenter account may be found [here](#registering-a-commenter-account).

To login through the embed, you may click the login dropdown in the top right corner, or click one of the login icons below the postbox.

#### Logging in through the 0 homepage

To login from the Disqus homepage, click the Login option in the top right corner of the page, and then select Publishers or Commenters. The Publishers option will bring you directly into your admin panel to begin moderating comments, while the Commenters option will bring you to your Disqus Home feed.

#### **Social Login on Disqus**

Disqus also offers several social login options (Facebook, X (Twitter), Google, Microsoft, and Apple), which allow users to create or authenticate with a Disqus account connected to these services.

Once you've authenticated your social login with a Disqus account, you'll be able to use that social login to log into Disqus with one click in the future. More information on these profiles may be found [here](#log-into-disqus-with-a-social-media-account).

#### Logging in through the specific site

Some sites may use what is called Single Sign-On (SSO) which allows you to comment on Disqus through their own account system. Sites using SSO can often be identified by having the site name appear in the login dropdown options, or a custom login button as seen below.
​
Some sites using SSO will only allow you to login through their site portal. If you do not see any login dropdown present in the Disqus embed, you'll want to log in directly with the site, and they will in turn log you into Disqus.

If you're having trouble logging in with SSO, we'd recommend contacting the moderator of the site in question. Because SSO sites have a completely different account system from Disqus, the site moderator will be able to better assist you as they have access to the user database you're trying to sign into.
​

#### Troubleshooting

-   If you're seeing the error **"The e-mail address you specified is already in use"** when trying to login, make sure you're using the login buttons or the login dropdown to login. The “Name” field with “Sign up for Disqus” above it is for new account registration only.

-   If you're having trouble logging in and/or staying logged in, ensure that you have third-party cookies enabled as outlined in our Use of cookies document.

    -   If you’re getting the message “**You must authenticate the user or provide author_name and author_email**”, this means that you have a bad cookie stored for Disqus, where it appears that you are logged in, but you are not. This can occur if you have Disqus open in multiple tabs, and get logged out in one of the other tabs. To fix this issue, clear your cookies in your browser, or search your browser cookies, and delete just the "0" cookies. After refreshing the page, you should be able to log in and post normally.

-   If you’re getting an error message saying “**We couldn’t log you in. Please check what you’ve entered**”, please ensure that you are entering the correct email and password values. If you’re not positive what your password was, you can request a password reset email at 0

    -   If the Forgot password form states that **we do not have an account registered with the email address you entered**, your account must be registered with another address. Please contact our team at disqus.com/support, or send a DM to our twitter account for additional help. You may also reach out to our team if your account is registered with an old email address you no longer have access to.

    -   If your email isn't recognized, it could indicate that you've only commented as a guest (includes Twitter, Facebook, and Google) or that you need to register a commenter account.

    -   If you don’t receive a password reset email within 24 hours of requesting one, please check your Spam filter, add \*@disqus.com and \*@disqus.net to your email’s whitelist and request a new password reset email.

        -   Each new password reset email requested will invalidate the tokens in the previously requested resets, so we don’t recommend making multiple requests for the same email within a few hours.

### Making Your Activity Private {#making-your-activity-private}

The ability to make your activity private allows you to prevent others from seeing the activity within your profile and keeps your activity from displaying within Digest emails.
​
​

Upon enabling your private activity, your previous and current activity will only be viewable by yourself. The following of your account will be disabled and those already following you will no longer be able to see any of your commenting activity, which would normally show up in your profile, digests and the My Disqus tab.

-   Your comment history and activity will only be viewable by you.

-   Existing users who are currently following your account will no longer be able see your activity via your profile or daily digest emails.

-   Disqus users will no longer be able to follow your account.

-   Existing followers will still show in your stats, but they will not be able to view your activity.

#### How to enable Private Activity

To make your profile private, log into your account, and visit your account settings. Please note that when logged in, your profile will still show your comments. To check if the profile has been correctly set as private, try logging out and visiting your profile page.

#### FAQ

#### Will people still be able to see my comments on communities where I regularly participate?

Yes. Anyone who visits a site where you have left a comment will still be able to see your comment.

#### Will new users be able to follow my account after it has become private?

Once your comment activity is marked as private, users will not be able to follow your account from that point on.

#### Why does my profile still show that there are people following me after I’ve made my activity private?

We do not erase your followers once you’ve made your activity private. People following your account will no longer be able to view your comment activity, but you will still be able to view who has followed your account.

#### Will people still be able to see my profile?

Please note that although your profile activity will be private, your profile can still be viewed which will continue to include your avatar, bio, website and social links (if you have connected your account to Facebook, Twitter or Google).

### Registering a commenter account {#registering-a-commenter-account}

Registering for a commenter account can be beneficial for several reasons. Some of which include tracking comments through a dashboard, managing account notifications, and following other user accounts.

Register for a commenter account using one of the following methods:

#### Our sign-up page:

#### Through the embed:

After registering through the embed, you'll receive a confirmation email with a link. Click the link to confirm that an account for your email address wasn't created by mistake.

If you delete the confirmation email by accident, you can still log into your account using your **email address** after requesting a password recovery email.

**The registration page says my username is already in use**

Usernames must be unique and can only be registered by one commenter at a time. If you are seeing this message, it means your desired username is already taken, and you should try registering a different name.

Keep in mind that the same [display name](#what-is-the-difference-between-my-username-and-my-display-na) can be used by more than one account.

Note that usernames are only for login and moderation purposes, and you have the option to display a different name with your comments after you have finished registering. More information on the difference between usernames and display names is noted at the top of this article.

### Site-Specific Profiles {#site-specific-profiles}

A site-specific profile is a Disqus profile that can only be used on the site where it was created.

#### How is a site-specific profile created?

A site-specific profile is created by logging into a site where single sign-on (SSO) has been enabled by the site owner.

There are a few ways you can login to create a site-specific profile: by clicking the site name in the login dropdown, by clicking the site-specific login button above social login methods, or through the site itself.

Curious about SSO? Find out more here: Single Sign-On

#### Where can I use a site-specific profile?

A site-specific profile can only be used on the site where it was created.

For instance, creating an account at 0 will allow you to comment on that site without signing up for your own Disqus account, but you won't be able to use that profile to comment on any other sites that use Disqus.

#### What can I do with a site-specific profile?

With a site-specific profile you can comment, reply, vote, share, hide media, flag comments, and recommend discussions. You'll also be able to receive web notifications for replies, upvotes, and when people follow you via the sidebar.

With a site-specific profile you also have the option to manually subscribe to a discussion and receive email notifications for any new comments that are posted to that thread. The email used for notifications will be the one you used to sign up for the site.

*Unsubscribe from email notifications as if you were a guest. More information here:* *Subscribe/Unsubscribe from Notifications.*

Although site-specific profiles have many of the features of a Disqus account, there are few things they can't do:

-   **Access Disqus Home** If you're logged into a Disqus account, you'll be able to access Home, but none of the content or settings will be associated with your site-specific profile.

-   **Adjust your profile or account settings** If you need to change your profile or accounts settings, contact the site where you made your site-specific profile.

-   **Follow other users** Although it's possible to follow other users on the site where you created your profile, notifications of their activity won't be shown in the sidebar.

#### I have a site-specific profile and I want to use Home, what can I do?

Home isn't available to site-specific profiles. If you'd like to use Home and gain access to all the features that a Disqus account has to offer, register a Disqus account by clicking here.

You'll still be able to comment on the site where you had your site-specific profile with your Disqus account, but the comment histories will be separate.

#### How can I tell if I'm using a site-specific profile?

-   Make sure you're logged out of Disqus at 0

-   Open a new browser tab and go to a site where you comment

If you're logged into Disqus on that site, the Disqus profile on that site is site-specific.

### Updating your Account Settings {#updating-your-account-settings}

In your Account Settings, you can update or change any of the following:

-   Profile (Name, location, bio, avatar, and more)

-   Avatar

-   Account - Username, email, and password

-   Connect Facebook, Twitter, Google

-   Delete account or manage data settings

-   Email notifications

-   Troubleshooting

-   Web notifications

-   Apps

-   Moderation

Your settings are just a click away whether you're using Disqus on Home or through the embed out on the web.

#### On Disqus Home

Click the gear icon, then choose "Edit Profile" or "Settings".
​

#### Through the Embed

Click your name (next to the notifications bubble) to open the drop down menu, then click "Edit Settings".

**Tip:** When logged into your Disqus account, your Settings may be accessed directly at 0

#### Profile

Once you've accessed your Disqus settings and navigated to the Profile tab, you'll then be able to edit your Avatar, Name, Website, Location, Bio, and Privacy Settings.
​

Note that the Name in your profile refers to your display name, which is different than your Username.

Your Name appears in your profile and is shown with your activity in Disqus. Your username appears in the URL of your profile and is used for logging in. More information on the differences between your Display Name and your Username can be found [here](#what-is-the-difference-between-my-username-and-my-display-na).

#### Updating your Avatar

You can manually update your avatar by uploading a picture from your computer, or by connecting to Disqus through Facebook or Twitter. All of these options can be found in the Profile settings shown above.

Image upload requirements:

-   File size limit: 1 megabyte.

-   Recommended dimensions: 128x128 pixels.

-   Transparent avatars are not supported. We're working on that.

#### Account

In your Account Settings, you can update your email address, username, and password.

-   Username - used in your unique profile address and for logging in, different than your Name

-   Email - used for email notifications and for verifying your account.

-   Password - used to log in to your registered Disqus account

#### Connect Facebook, Twitter, and Google

Connected accounts are links to your social networks such as Facebook, Twitter and Google. By linking to a service, you are able to log into Disqus with your social account.

**As of April 2015, your social account link will no longer appear publicly in your Disqus profile.**

On this page, you can also manage your Data sharing settings or permanently delete your Disqus account.

**Important: We completely remove all of your data when you delete your account. While we do retain your comments for context purposes, your comments are completely anonymized to remove all traces of your personal information. Deleting your account also removes all Disqus sites you own and their respective comments.**

#### Email notifications

Email notifications are sent based on your preferences, allowing you to receive periodic digests or marketing materials, or be notified by email when you receive replies to your comments or new badges on your account. You can modify what types of emails are sent to you, or completely unsubscribe in your Email Notification settings.

Not receiving your notifications or having trouble unsubscribing? Please see Troubleshooting Email Notifications.

#### Web Notifications

Web notifications for things like replies, upvotes, new followers, and invites to discussions show up in your Inbox. In the Web Notifications settings, you can control which ones you want to see.

#### Apps

If you have granted any 3rd party apps access to your Disqus account, they will appear in your Apps settings and can be revoked if needed.

#### Moderation

If you are a moderator on a Disqus forum, you can configure your email notifications for each forum in your Moderation settings.

#### Enabling Notifications for Flagged Comments

If you'd like moderators on your site to receive flagged comment notifications in addition to their other moderator notifications, enable "Email moderators when a post is flagged" in your forum settings at Settings > General > Community Rules.

### Use of cookies {#use-of-cookies}

To post comments and have your Disqus login information remembered across pages, cookies – specifically, first party (and in some browsers, third-party) cookies – must be enabled.

#### Third party cookies

Disqus doesn't require third party cookies to be enabled for basic functionality (like posting), but they are required to keep you logged in between pages in some browsers. If you're using Safari or a new version of Firefox, you won't need to enable third-party cookies to stay logged in because those browsers consider our cookies to be "first-party" because they're used in a trusted way.

If you're using a browser other than Safari or Firefox and you'd rather disable third-party cookies for all sites, but you'd still like Disqus to remember that you're logged in, simply add an exception for disqus.com and its subdomains to your cookie settings. Instructions on how to do that are noted underneath each browser below.

#### Single Sign-on

Disqus doesn't store an authentication cookie when signing in with Single Sign-on (website-specific profiles). The site is required to pass the user authentication each time the user loads the comments.

#### Cookies we set

These are the cookies we may set for someone visiting a site with Disqus embedded on it. This list is for sites complying with the EU cookie law.

-   \_\_qca (Domain: .disqus.com)

-   mc (Domain: .quantserve.com)

#### Google Analytics

-   UID (Domain: .scorecardresearch.com)

-   UIDR (Domain: .scorecardresearch.com)

#### Internal cookies

-   disqus_unique (Domain: .disqus.com)
     *Internal statistics, used for anonymous visitors (Sigma)*

-   testCookie (Domain: mediacdn.disqus.com)
     *Used to check whether the browser accepts 3rd-party cookies.*

#### How to enable cookies

If you can't comment, or aren't able to stay logged in, check your cookie settings using the instructions for your browser below.

#### Internet Explorer

1.  Click **Tools** in the upper right hand area of your browser

2.  Find and click **Internet Options**

3.  Go to the Privacy tab and click **Advanced**

4.  Check **Override automatic cookie handling** (if it isn't already)

5.  Make sure First-party Cookies are set to **Accept**

6.  Set Third-party cookies to **Accept**

#### If third party cookies are blocked:

1.  Under the **Privacy** tab click the **Sites** button

2.  Type **disqus.com** and click **Allow**

#### Chrome

1.  Click the Chrome Menu Icon -> **Settings**

2.  Near the bottom of the page, click **Show advanced settings**

3.  In the "Privacy" section, click **Content settings**

4.  In the "Cookies" section, choose **Allow local data to be set (recommended)**

5.  Make sure "Block all third-party cookies..." and "Clear cookies..." are **unchecked**

#### If third party cookies are blocked:

1.  Under **Content settings** click **Manage Exceptions...**

2.  Type **\[\*.\]disqus.com** as a hostname pattern and set to **Allow**

#### Firefox

1.  Windows: Go to **Tools -> Options**
    Mac: Go to **Firefox -> Preferences**

2.  Click the **Privacy** tab

3.  Under History, select **Use custom settings for history**

4.  Check both **Accept cookies from sites** and **Accept third-party cookies**

#### If third party cookies are blocked:

1.  Under the **Privacy** tab click **Exceptions**

2.  Type **disqus.com** and then click **Allow**

#### Safari

1.  Windows: Go to **Edit -> Preferences**
    Mac: Go to **Safari -> Preferences**

2.  Click the **Security** tab

3.  Under Block Cookies, select **From third parties or advertisers** or **Never**

### User Profiles 101 {#user-profiles-101}

Disqus commenter profiles make it even easier to get to know and follow the people participating in your favorite online communities.

We’ve made it easier to follow your favorite community members by including a nice large follow button as well as including two new tabs that show a user’s list of followers, and a list of people that the user is following.

The new profiles were designed with mobile in mind so they will automatically resize to fit all different screen sizes – this makes the whole experience seamless no matter what device you’re using.

#### Update Your Own Profile

You can make your new profile shine by uploading your own custom avatar as well as adding a short bio about yourself. To experience the new Disqus profiles, click on any commenter profile photo your find in your favorite Disqus community.

#### The latest profile features at a glance:

-   ***Profiles open in the sidebar next to the comments so you never lose your reading place again***

-   ***Ability to see the recent activity, comments, following lists, and favorite discussions of any user***

-   ***Report a user that's breaking the Disqus Terms of Service (spam, inappropriate profile, etc.)***

-   ***Responsive styles automatically resize the profile to fit different window heights***

-   ***Automatic Translations: The profile modal uses the same language as the embed from which it is launched***

#### Additional Resources

#### Your homepage on Disqus

#### Updating your Disqus profile settings

#### Making your activity private

#### Site specific profiles

### What is the difference between my Username and my Display Name? {#what-is-the-difference-between-my-username-and-my-display-na}

On Disqus, the Username and Display name are separate, and serve distinct purposes for a user's account. The Display Name is what shows up next to a user’s comments, and does not need to be unique. The Username is a separate account identifier, and indicates the direct URL which can be used to visit a user’s profile.

#### What are each of the names used for?  **Display name**

-   Simple handle to easily identify the user in thread conversations

-   Does not need to be unique, so allows the user to be known by whatever name they prefer

-   Can be edited in the "Name” field in the Profile tab in Profile Settings

####  **Username**

-   Identifies the unique URL at which a particular account can be viewed (for this reason, usernames must be unique)

-   Can be used in place of the account’s email address to log into an account

-   Can be edited in the "Username” field in the Account tab in Profile Settings

#### Impersonation

As Display names are not required to be unique, having the same Display name as another user is not in violation of Disqus’ terms of service. This alone is not sufficient evidence to support a case of Impersonation. If a user has the same Display name as another user, a duplicated avatar or corresponding post history that clearly demonstrates an intent to impersonate must also be present for action to be taken.

### Your Homepage on Disqus {#your-homepage-on-disqus}

People are talking passionately all across the network of Disqus-powered sites. There's never been one place to find new discussions to join about the topics you care most about. Your homepage on Disqus is where you can find everything you might care about across the Disqus network in one place.
​
Home provides a truly integrated experience for people and sites you follow. See whenever a new article is posted on a site that you follow, every time a friend replies or upvotes your comment, and whenever someone follows you. Also, you can now reply to other users and vote on their comments from within your Disqus Home feed.
​
With Home on Disqus, you'll be able to find new content, and check out the articles and updates from people you already follow, all in one place.
​
**Overview**
Here's a snapshot of what you can find in in Home:

**Notifications:** View all of your notifications in one place with the ability to reply and follow from your notifications feed. Get notified of replies, upvotes, discussions you've been invited to, and follows in the Disqus embed or in Home.
​
**Recommended:** Community recommendations are a central feature of Disqus Home. When you ♥ (“Recommend”) a discussion, it shows up in feeds of people following you. If you follow communities and people, you’ll see what discussions are getting recommended by others.

**User profiles:** A revamped user profile shows off what Disqus users are interested in and talking about. Also, if you're in a conversation with another user, you can now reply to them from their profile.
​
**Where are my notifications?**
When you get a new notification, your Inbox icon in the main navigation will change color and display the number of unread notifications that are waiting for you. When you click on the icon, you'll be taken to your Inbox.
​
In the Replies tab you'll find replies to your comments. Upvotes, invites to discussions, and follow notifications will appear in the Most Recent tab.

​
**How do i follow a community or site?**
Follow a community or a site by going to their page and clicking on the follow button in the top banner. Once you start following the site, new content will start flowing into your home feed.

**How do I see content from people I follow now?**
Content from people that you follow will now show up in your Recent Comments feed in reverse chronological order, as well as in your Recommended feed.
​
**How can I find new people and communities to follow?**
New to Disqus?
If you haven't started following any communities or other Disqus users, your feed will be empty. Populate it by following some of the interesting communities and commenters.

Communities can be followed from their [Site Profile](#site-profiles).

**How do I follow someone?**
You can follow other Disqus users the way you normally would from the discussion or on someone's profile, but we've also added some new ways to follow people.
​
If someone starts following you, or they reply to one of your comments, the you can follow them from within your notification feed by clicking the follow button next to their name.

**How do I reply or like a comment from my notification screen?**
If you get a notification that someone has replied to one of your comments, there's no longer a need to go check out the discussion somewhere else. You can read, upvote and reply from your Inbox.
​
Upvote a comment the way you normally would:

Replying to someone is also familiar:

**How do I update who I am following?**
You can manage communities and people that you follow by clicking into your profile and clicking the "Following" tab. From here, you can unfollow communities or people. You can also remove people from Followers list so they no longer see your activity.

**How can you help give feedback to improve the product?**
Join the discussion here on Discuss Disqus.

## What is Disqus? {#cat-what-is-disqus}

### Common Questions About Disqus {#common-questions-about-disqus}

Still have a few questions before getting started with Disqus? We’ve put together a list of some of the more common questions new users have. If you don’t see the answers you’re looking for here, or on our website, please visit our Knowledge Base where you can search our documentation, or contact us with questions via the link at the bottom of any help article.

Disqus is free to use with ads for most sites. In addition, we offer several paid packages including Plus, Pro, and a Business tier. Visit the [Pricing and Plans](#comments-pricing-and-plans) page for more detailed information.

If you are a small, nonprofit site, you may be eligible for a free subscription to our Comments Plus plan. You can learn more by emailing publisher-success@disqus.com.

#### Do you offer Single Sign-On (SSO)?

Single Sign-On (SSO) is currently available as an add-on for users with a Disqus Business subscription. If you would like to subscribe to Business, you can contact us to be set up with the plan at Admin > Settings > Subscription and Billing.

#### We’re a huge site. Do we need special support?

No — Disqus can uniquely say that we have the proven scale for any site and any amount of peak traffic.
​
The size of the Disqus network has never been bigger. We’re growing fast and just hit the 1 billion monthly unique user milestone. As of May 2013, Disqus handles more than 7 billion monthly pageviews. Put another way, a site with 100 million monthly pageviews represents a little over 1% of total Disqus traffic.
​
With millions of sites on our network, Disqus is constantly monitoring the health of our network and publicly documents this at 0.

#### How do I get help?

Our support team has meticulously documented everything about Disqus, from installation to the Disqus API.

If you can’t find what you’re looking for in our Knowledge Base, feel free to contact us directly.

#### Does Disqus support translations?

Disqus supports dozens of languages. If we don’t yet offer your language, you can help us get there by providing your own translations! Our help doc provides everything you need to know about language options with Disqus.

#### How is spam filtered?

Disqus has an incredibly effective built-in spam filter which is executed right after a comment is posted. If you want to be extra cautious, you can help protect your site from spam with a few extra steps.

#### What effect does Disqus have on my site’s performance?

Disqus loads asynchronously (after the rest of your site has finished loading), so it won’t affect the rest of your page performance. For more detail, please see our help doc about page load performance.

#### Can Disqus help with SEO?

While better SEO is never the goal of an active commenting community, it is a nice side effect. Disqus has demonstrated that people spend more time on pages where Disqus is installed. This translates into more page views and more comments, which keep pages fresh and give search engines more data to crawl.

Disqus has worked closely with search engines, including Google, to ensure Disqus is crawlable, though ultimately, indexing is out of Disqus’ hands. That being said, there is always the option to sync comments locally to be rendered in the HTML of the page. Please see our help doc on indexing and syncronization for more information.

#### With which certifications does Disqus comply?

From government agencies to NGOs, the Disqus network is filled with sites that have particular needs. Our Support team is happy to answer any specific questions, but some of the more common standards that Disqus complies with are as follow:

-   Section 508 (screen reader accessible)

-   US Government approved social application

-   TRUSTe/EU Safe Harbor Compliant

-   SSAE 16 (formerly SAS 70) certified data center

Other related questions can often be addressed by our Terms & Policies documentation.

#### What about data ownership — do I own the comments posted through Disqus?

You own your data, period. Further, Disqus makes it easy both to import and export data.

If you’re just getting started, you can bring all of your old comments into Disqus through an easy import process.

Of course, you can always export your existing Disqus comments if you decide at any time to leave Disqus, or need a backup copy of your comments.

#### This FAQ is awesome and all my questions are answered! How do I get started?

Great news! Glad to have you on board with the web’s community of communities.

Our [Publisher Quick Start Guide](#publisher-quick-start-guide) should get you up-and-running in no time.

### Disqus Glossary {#disqus-glossary}

Disqus has its own vocabulary that we use to describe our different tools and features. Here's a list of a few of our need-to-know terms and definitions.
​

**API** - The API enables developers to communicate with Disqus data from within their own applications. Our documentation provides further explanation on how to use our API, specifically. You can find more information on APIs, in general, on Wikipedia and Quora.
​
**Ban User** - The "Ban User" setting lets you block a user, IP address, or email address to prevent certain people from posting comments on your site.
​
**Comment Count** - The number of comments per post. You can add a comment count link to display the number of comments below the title of each post. Get the how-to.
​
**Community Guidelines** - Rules of engagement for commenters on your site. Guidelines can cover topics like privacy, etiquette, expectations, and moderation settings. Learn more.
​
**Configuration Variables** - These are parameters for Disqus's behaviors and settings. Configuration variables must be defined on each page that Disqus is loaded on, so be sure to include configuration variables in your dynamic templates that render pages.
​
**Display name** - A full name is the name carried across the Disqus network on your profile and is the name displayed with your comments. Your full name does not have to be unique and can contain spaces.
​
Full names are optional, though highly recommended. For added security, we recommend choosing a full name different from your username.
​
**Embed** - The discussion thread powered by Disqus—this is the comments section that Disqus adds to your site.
​
**Engagement** - Engagement indicates how active your commenters and readers are. Engagement is measured by number of comments and votes.
​
**Forum** - A forum is your website community on the Disqus network. When you register your website on Disqus, you are creating a forum with a unique shortname. Your shortname is different than your username.
​
Every website using Disqus has a unique forum which is moderated by their respective administrators. A forum consists of the comments and comment threads posted by other users. Users, a.k.a. community profiles, are not unique to forums since people can belong to any number of communities on Disqus.
​
**Import & Export Tools** - Let you upload comments from another system into Disqus or download your Disqus comments onto your computer.
​
**IP Address** - A unique identifier for each computer connected to the network. You can ban an IP address to ensure that no commenters using that IP can post on your site.
​
**Migration Tools** - Let you update or move discussion threads on your site to a new thread. Migration tools are useful when you update your domain name, change your blogging system, or want to merge discussion threads.
​
**Moderation Panel** - Site owners moderate the comments posted to their site (approve, mark as spam, delete) from the moderation panel. A forums' moderation panel can be accessed with the following link: 0
​
**Moderator** - A moderator is responsible for managing a site's community. Moderators delete and approve comments, mark spam, block or unblock commenters, and handle disputes between commenters. There are several different moderator types, including: Site Founder (primary moderator who can edit settings or comments), Site Admin (can edit settings or comments), and Site Moderator (can edit comments).
​
**Pre-moderation** - Turn on pre-moderation controls to require moderator approval for all comments.
​
**Shortname** - A unique identifier for your site that appears in your account URL. Access your site's Disqus account by visiting yoursitesshortname.disqus.com/admin
​
**Thread** - The string of comments that readers post on your site. Disqus creates comment threads for your site so that your readers can have discussions about your content.
​
**Trust User** - Mark a user, IP address, or email address as "trusted" so those users can bypass certain moderation filters (such as spam).
​
**Trusted Domains** - Domains set by websites to specify which domains are allowed to create and load new threads with the Disqus javascript embed. We recommend that you add a trusted domain to ensure that your comments thread is hosted exclusively on your site.
​
**Username** - The name you use to login to Disqus. A username must be unique, and cannot be in use by more than one commenter at a time. A username cannot contain any spaces or special characters.
​
Usernames are used mostly for two purposes: logging into Disqus, and moderation (for site owners).
​
**Word Filters** - Use word filters to create a list of restricted words that automatically get queued for your review. Comments containing restricted words will not appear in the discussion until they've been approved by a moderator.
​
​

#### Ads Terms

**Ad Revenue** - The amount of money you earn from Disqus ads.
​
**Below-the-fold** - The area of a webpage that is only visible after a reader scrolls down the page.
​
**Impressions** - An impression is counted every time a reader views a page on your site.
​
**RPMv** - Revenue per a thousand viewable impressions. This is the revenue you'll receive every one thousand times your readers scroll down to view the ads in your Disqus forum.
​
**Sponsored Links** - Ads for popular articles from around the web.
​
**Sponsored Story** - The default ad type. These are cost-per-impression ads informed by reader engagement.
​
**Viewable Impressions** - How often a reader scrolls down to the Disqus forum below-the-fold to see Disqus ads. Your Ads earnings are based in part on Viewable Impressions.
​
**Viewability Percent** - How often a reader scrolls down to view Disqus ads as a percentage of total pageviews. In other words, the percentage of viewable impressions per total impressions on your page.

### How does Disqus work? {#how-does-disqus-work}

Disqus is a networked community platform for your website. To learn more about Disqus, read [What is Disqus](#what-is-disqus) or visit the the Disqus website.

[Get Started as a Commenter](#commenting-101)

[Get Started as a Publisher](#publisher-quick-start-guide)

This page intends to provide a deeper understanding of how Disqus works, both from a conceptual and a technical level. This page is written for our advanced users who are curious about the inner-workings of Disqus — it's absolutely not required reading in order to use the service.

At its core, Disqus is a third party system that provides commenting and other community features. The Disqus service also acts as an intrinsic network that connects each of these enabled websites together. Disqus uses JavaScript to embed, or display, the system onto the page.

We use some terminology that may be unfamiliar or used differently from other services.

**Comments** in the Disqus backend are called **posts** (and will be described as such in our API documentation). Because Disqus is used with blogs and other content management systems, using "posts" leads to ambiguity, so they are called comments in the frontend or when describing to end-users.

**Threads** contain comments. A thread is associated with a page which has Disqus embedded. For example, a page located at 0 will have one unique thread associated with that page. This thread will contain all of the comments on the page, as well as everything else that is relevant to that instance of Disqus (such as likes, participating users, and other metadata). Threads are uniquely identified by either a page-provided identifier or a URL.

A **forum** is a website's account on Disqus. Note that this is not the same as the user account which registered the website. A forum indicates the website's community on Disqus and is identified by the forum's shortname. Take this example if your website is located at 0. Your website's name may be called *My Example Website*. Your forum shortname on Disqus may be *myexamplewebsite*.

For more terminology definitions, please read the [Disqus Glossary](#disqus-glossary).

#### Loading Disqus

Whether websites use Disqus plugins or just manually embed the script, the system is loaded onto the page in generally the same way.

When a user visits a webpage that includes Disqus (for example, a blog post), the page makes a request to Disqus. Disqus uses the information defined on the page, called configuration variables, to locate the correct thread. Disqus will look up the associated thread and, if found, embed the correct threads with all the right comments onto the page. If an associated thread was not found, Disqus will create a new page with the data provided (again, in the configuration variables), and environment metadata such as page URL, page title, and current datetime.

When a thread is located or created, the Disqus script continues and generates the appropriate HTML, JavaScript, and CSS for your page and embeds it in the right location.

#### Posting to Disqus

The core of Disqus is posting comments and we strive to handle this in a smooth and secure way. Nearly all of Disqus is dynamically rendered HTML through JavaScript. However, the system makes use of iframes to directly communicate back to the Disqus servers when users post content. We use iframes to ensure security and to protect users against websites maliciously posting comments on their behalf.

When a user posts his or her comment, it is done within an iframe and sent directly to the Disqus servers. To the user, this entire experience is seamless and feels native to the website.

#### Further reading

This page describes how the core Disqus system works. If you'd like to learn about how comment counts are calculated and displayed, read How do comment counts work?.

### What is Disqus? {#what-is-disqus}

Disqus is a networked community platform used by hundreds of thousands of sites all over the web. With Disqus, your website gains a feature-rich comment system complete with social network integration, advanced administration and moderation options, and other extensive community functions. Most importantly, by utilizing Disqus, you are instantly plugging into our web-wide community network, connecting millions of global users to your small blog or large media hub.

Disqus works on just about any type of website or blog and can be installed either with a drop-in code snippet or by using one of the plugins available on our Install page. You can also customize and tweak Disqus for your website with extensive APIs and JavaScript hooks. Check out the Quick Start Guide or visit our homepage for more information and a demo of Disqus.

Spark engagement with comments! Disqus is the world's most trusted comments plugin. It makes communities easier for publishers to manage, and readers love using it.

-   **Looks good.** Automatically adapts to your sites design and colors, or you can set it to your own liking.

-   **Works everywhere.** Supports devices from desktop to mobile.

-   **Used across the world.** 70 languages supported and counting.

#### Disqus Ads

Earn money with native ads! Native advertising made simple. Disqus Ads helps eligible publishers generate revenue from your growing audience.

-   **Flexible options.** Native ad units are placed around comments. Pick the type of ad most suited to your site. You're in complete control.

-   **Adapts to your page.** Disqus ads are responsive -- they adapt to the look of your site and change layout based on device and width.

-   **Revenue analytics.** Understand how your website's traffic and audience engagement affect the revenue you're generating.

Ready to try Disqus Ads? Head over to the **[Disqus Advertising](#ads-faq)** page for more details!

## Miscellaneous {#cat-miscellaneous}

### Channel Help {#channel-help}

*Please note that Disqus Channels as they have previously existed have been sunset, and are no longer in operation. At this time, Discuss Disqus is our only operating channel, used for Disqus product feedback.*

If you run into any issues while using Disqus, we recommend posting about them in our community support channel, Discuss Disqus. This allows for discussion and resolution with our most experienced users, and is monitored by the Disqus community team and staff.
​
To create a discussion, you'll need to add the following information:

**Title**

Please be as clear and concise as possible about the issue you are currently facing. This will help others add information if they are getting the same experience or have the same question.
​

**Topics**

Topics are indexed in the channel, and help older discussions to be found and categorized correctly. At least one Topic will be required to post your discussion. You may find a list of commonly used Topics in the left sidebar while on the main Discuss Disqus page.
​

**Description**

Please add additional details as to what you are experiencing, this will help others assist with your issue. First, describe the behavior or experience as accurately as possible, and what you would expect to occur instead.
​
Additionally, please note whether you are experiencing your issue across multiple devices, and if so, please provide the Operating System (OS) and Browser versions you are getting this experience on, as sometimes issues can appear only on certain configurations. This information can typically be obtained by looking for the "About this computer" or "About this browser" options.

An example for a Mac computer using Chrome is as follows:

OS: Mac Sequoia 15.1
Browser: Chrome Version 146.0.7680.178

#### FAQ

##### Can I create a new channel?

No, at this time no new channels can be created and old channels cannot be accessed. Instead, please create a new site with Disqus that most closely matches what you wish to talk about.

##### How do I report a user or discussion?

While we want to harbor an environment of diverse opinions and communities, we do have a set of rules that all users and communities are expected to abide by, our Basic Rules. If you find that a user or discussion is in violation of any of these rules and needs to be reported, check out our documentation on how to report abuse.
​
For any issues with the Disqus platform or your account, please visit Discuss Disqus. If you are unable to log into your Disqus account to post to the channel, you may also reach out to us at publisher-success@disqus.com, or DM the \@DisqusSupport twitter account

### Disqus Advertising Content Guidelines {#disqus-advertising-content-guidelines}

Disqus Ads provides content marketers with reach into active Disqus communities all over the web. It’s a unique opportunity to capture the attention of a highly engaged audience that is vocal about the things they care about. However, the nature of this audience and the subject-matter diversity across sites using Disqus requires a broad sensibility of acceptance. At Disqus, publishers and users trust us with their space to convene and interact. To maintain that trust, we’re continually enhancing the quality and relevance of content delivered through our Ads features.

The purpose of this document is to provide content marketers with parameters and direction to inform the development of campaign content and headlines that will resonate with the Disqus audience. Disqus will continually refine these parameters based on feedback from users, publishers and advertisers alike. We’ll also modify these as we introduce new tools to customize campaign reach.

Disqus reserves the right to reject content deemed impermissible and inappropriate for its publisher partners. Disqus also reserves the right to suspend campaigns that generate negative feedback after activation. In all cases, Disqus account team members will work with each client to repurpose content and campaign headlines to ensure it both generates engagement and resonates with the Disqus audience. More specific guidance is included in the following.

As a general rule, content intended to provoke negativity, sensationalize or instigate will not be accepted. Explicitly negative headlines directed to a single individual or organization will not be accepted. Content that uses crass or sexually explicit language will also not be accepted. More specifically, content that fits any of the following criteria will not be accepted or sent back for revision:

-   Attempts to capitalize on global, national or local crisis such as natural disasters, political unrest or social issues in poor taste

-   Defames or provokes groups or individuals on the basis of sex, race, religion, social beliefs or national origin

-   Leverages celebrity rumor or scandal in an egregious manner

-   Makes a pejorative claim about a company competitor or individual

-   Uses sexually explicit language

-   Refers to sexual(ized) body parts or sexual or bathroom activities

-   Contains profanity or crude or off-putting slang

-   Solicits funds or makes promises or claims of profit

-   Refers to violence or violent acts

In addition, companies and organizations operating in fields of frequent controversy may not be accepted depending on the nature and substance of content. These fields include adult entertainment, firearms, political advocacy, government, defense contracting, energy, tobacco and finance.

#### II. Disclosure and Attribution

Advertisers must clearly attribute their company or brand name as the source of the content within the Disqus Ads feature. This attribution serves to help users distinguish the advertiser content from other content as well as ensure the brand name is featured. The use of misleading or vague content source names will not be accepted.

#### III. Types of Content

All promoted content links must take the user directly to the content described in the headline. Promoted links must deliver on the user expectation that they’ll be taken to a page to read more about the subject matter included in the headline. It cannot direct them to a conversion or lead generation vehicle such as a whitepaper, direct response promotion or webinar registration page. It cannot link to an automatic download page (including a PDF), e-commerce site, homepage, pop-up ad, or product sell-sheet.

Forms of content to link to include: blog posts, videos, slideshows, articles and third-party reviews. In short, it must be content and content that is represented in a transparent manner.

#### IV. Subject Matter that Will Work

Disqus users are more engaged and more likely to share than the average Internet audience. Ultimately, they want to discover more content worth talking about. But it need not be sensationalist in nature to engage them. Following are recommendations for the kinds of content that have been found to be effective at reaching the Disqus audience:

-   Instructional or educational. With clear descriptors. (Example: How to Make the Perfect Margarita)

-   Thought provoking. With headlines that prompt a question. (Example: What NASA Could Teach the Energy Industry)

-   Material. With a callout to intended audience. (Example: Acme CEO: A Letter to Our Shareholders)

-   Punchy. With something at stake. (Example: 5 Health Facts that Could Save Your Life)

-   Current. With timely relevance. (Example: Plan Your Last Minute 4th of July BBQ)

-   Playful. Without being overt or offensive. (Example: The Miley Cyrus Dress that Has Everyone Talking)

-   Video. With “video” in the headline. (Ex: Video: Beyonce Rocks the Superbowl)

For more guidance, please read our Editorial Playbook.
