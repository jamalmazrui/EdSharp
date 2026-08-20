---
title: "Chrome DevTools Protocol: Documentation and Tutorials"
date: 2026-08-11
lang: en
---

# Chrome DevTools Protocol: Documentation and Tutorials

This volume gathers thirteen documentation and tutorial pages about the Chrome DevTools Protocol, the interface a program uses to drive Chrome or Microsoft Edge over a WebSocket. It was assembled on 2026-08-11 from snapshots taken by getCdpDocs.py on 11 August 2026, and it carries each publisher's own words rearranged for reading by heading.

Two things a reader should know before starting. First, this is developer documentation rather than help for the user of an application, so it sits outside the scope of the app help guides and is offered on its own. Second, three of the thirteen sources are third parties writing about the protocol rather than the protocol's own publisher, and each article says at its head who published it.

The two official reference pages are the weak point of the collection. Their domain, command and event reference is drawn in the browser from a JSON file, so a saved copy holds only the list of domain names. That list is reproduced here because it is a useful map, but the reference itself must be read another way, and the note under those two articles says how.

## Contents

- Introductions and tutorials
    - [Getting started with the Chrome DevTools Protocol](#getting-started-with-the-chrome-devtools-protocol)
    - [An introduction to the Chrome DevTools Protocol](#an-introduction-to-the-chrome-devtools-protocol)
- Libraries and bindings
    - [chrome-remote-interface recipes](#chrome-remote-interface-recipes)
    - [PyCDP overview](#pycdp-overview)
- Playwright
    - [BrowserType and connect_over_cdp in Playwright for Python](#browsertype-and-connect_over_cdp-in-playwright-for-python)
    - [BrowserType and connectOverCDP in Playwright for JavaScript](#browsertype-and-connectovercdp-in-playwright-for-javascript)
    - [BrowserType and ConnectOverCDPAsync in Playwright for .NET](#browsertype-and-connectovercdpasync-in-playwright-for-net)
    - [CDPSession in Playwright for .NET](#cdpsession-in-playwright-for-net)
    - [CDPSession in Playwright for JavaScript](#cdpsession-in-playwright-for-javascript)
    - [CDPSession in Playwright for Python](#cdpsession-in-playwright-for-python)
- Protocol reference
    - [Protocol domains, latest (tip-of-tree)](#protocol-domains-latest-tip-of-tree)
    - [Protocol domains, stable 1.3](#protocol-domains-stable-13)
- Miscellaneous
    - [Awesome Chrome DevTools: protocol resources](#awesome-chrome-devtools-protocol-resources)

## Introductions and tutorials

### Getting started with the Chrome DevTools Protocol

Published by Andrey Lushnikov. Source page: [Getting started with the Chrome DevTools Protocol](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/README.md).

Think twice before using CDP directly for browser automation. You'll be better off with [Playwright](https://github.com/microsoft/playwright)

Not convinced? At least use [Puppeteer's CDPSession](https://github.com/aslushnikov/getting-started-with-cdp#using-puppeteers-cdpsession).

See also [Contributing to Chrome DevTools Protocol](https://docs.google.com/document/d/1c-COD2kaK__5iMM5SEx-PzNA7HFmgttcYfOHHX0HaOM/edit#heading=h.e6mz7k1mw34a)

NOTE: An interactive protocol viewer is available at [vanilla.aslushnikov.com](https://vanilla.aslushnikov.com).

The Chrome DevTools Protocol allows for tools to instrument, inspect, debug and profile Chromium, Chrome and other Blink-based browsers. Many existing projects currently use the protocol. The Chrome DevTools uses this protocol and the team maintains its API.

To run scripts locally, clone this repository and make sure to install
dependencies:

```
git clone https://github.com/aslushnikov/getting-started-with-cdp
cd getting-started-with-cdp
npm i
```

When Chromium is started with a `--remote-debugging-port=0` flag, it starts a Chrome DevTools Protocol server and prints its WebSocket URL to STDERR. The output looks something like this:

```
DevTools listening on ws://127.0.0.1:36775/devtools/browser/a292f96c-7332-4ce8-82a9-7411f3bd280a
```

If you don't get any output, make sure to add `--enable-logging` to see the WebSocket URL.

Clients can create a WebSocket to connect to the URL and start sending CDP commands. Chrome DevTools protocol is mostly based on [JSONRPC](https://www.jsonrpc.org/specification): each comand is a JavaScript struct with an `id`, a `method`, and an optional `params`.

The following example launches Chromium with a remote debugging port enabled and attaches to it via a WebSocket:

File: ./wsclient.js

```
const WebSocket = require('ws');
const puppeteer = require('puppeteer');

(async () => {
  // Puppeteer launches browser with a --remote-debugging-port=0 flag,
  // parses Remote Debugging URL from Chromium's STDOUT and exposes
  // it as |browser.wsEndpoint()|.
  const browser = await puppeteer.launch();

  // Create a websocket to issue CDP commands.
  const ws = new WebSocket(browser.wsEndpoint(), {perMessageDeflate: false});
  await new Promise(resolve => ws.once('open', resolve));
  console.log('connected!');

  ws.on('message', msg => console.log(msg));

  console.log('Sending Target.setDiscoverTargets');
  ws.send(JSON.stringify({
    id: 1,
    method: 'Target.setDiscoverTargets',
    params: {
      discover: true
    },
  }));
})();
```

This script sends a [`Targets.setDiscoverTargets`](https://vanilla.aslushnikov.com/#Target.setDiscoverTargets) command over the DevTools protocol. The browser will first emit a [`Target.targetCreated`](https://vanilla.aslushnikov.com/#Target.targetCreated) event for every existing target and then respond to the command:

```
connected!
Sending Target.setDiscoverTargets
{"method":"Target.targetCreated","params":{"targetInfo":{"targetId":"38555cfe-5ef3-44a5-a4e9-024ee6ebde5f","type":"browser","title":"","url":"","attached":true}}}
{"method":"Target.targetCreated","params":{"targetInfo":{"targetId":"52CA0FEA80FB0B98BCDB759E535B21E4","type":"page","title":"","url":"about:blank","attached":false,"browserContextId":"339D5F1CCABEFE8545E15F3C2FA5F505"}}}
{"id":1,"result":{}}
```

A few things to notice:

- [Line 19](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/wsclient.js#L19): Every command that is sent over to CDP must have a unique `'id'` parameter. Message responses will be delivered over websocket and will have the same `'id'`.
- Incoming WebSocket messages without `'id'` parameter are protocol events.
- Message order is important in CDP. In case of `Target.setDiscoverTargets`, it is (implicitly) guaranteed that all current targets will be reported before the response.
- There's a top-level "browser" target that always exists.
Before advancing any further, consider a simple helper function to send DevTools protocol commands and wait for their responses:

File: ./SEND.js

```
// Send a command over the WebSocket and return a promise
// that resolves with the command response.
module.exports = function SEND(ws, command) {
  ws.send(JSON.stringify(command));
  return new Promise(resolve => {
    ws.on('message', function(text) {
      const response = JSON.parse(text);
      if (response.id === command.id) {
        ws.removeListener('message', arguments.callee);
        resolve(response);
      }
    });
  });
}
```

NOTE: this `SEND` implementation is very inefficient - don't use it as-is! Check out Puppeteer's [Connection.ts](https://github.com/puppeteer/puppeteer/blob/main/packages/puppeteer-core/src/cdp/Connection.ts) for a
better version.

Chrome DevTools protocol has APIs to interact with many different parts of the browser - such as pages, serviceworkers and extensions. These parts are called Targets and can be fetched/tracked using [Target domain](https://vanilla.aslushnikov.com/#Target).

When client wants to interact with a target using CDP, it has to first attach to the target using [Target.attachToTarget](https://vanilla.aslushnikov.com/#Target.attachToTarget) command. The command will establish a protocol session to the given target and return a sessionId.

In order to submit a CDP command to the target, every message should also include a `sessionId` parameter next to the usual JSONRPC’s `'id'`.

The following example uses CDP to attach to a page and navigate it to a web site:

File: ./sessions.js

```
const WebSocket = require('ws');
const puppeteer = require('puppeteer');
const SEND = require('./SEND');

(async () => {
  // Launch a headful browser so that we can see the page navigating.
  const browser = await puppeteer.launch({headless: false});

  // Create a websocket to issue CDP commands.
  const ws = new WebSocket(browser.wsEndpoint(), {perMessageDeflate: false});
  await new Promise(resolve => ws.once('open', resolve));

  // Get list of all targets and find a "page" target.
  const targetsResponse = await SEND(ws, {
    id: 1,
    method: 'Target.getTargets',
  });
  const pageTarget = targetsResponse.result.targetInfos.find(info => info.type === 'page');

  // Attach to the page target.
  const sessionId = (await SEND(ws, {
    id: 2,
    method: 'Target.attachToTarget',
    params: {
      targetId: pageTarget.targetId,
      flatten: true,
    },
  })).result.sessionId;

  // Navigate the page using the session.
  await SEND(ws, {
    sessionId,
    id: 1, // Note that IDs are independent between sessions.
    method: 'Page.navigate',
    params: {
      url: 'https://pptr.dev',
    },
  });
})();
```

Things to notice:

- [Lines 22](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L22) and [33](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L33): clients must provide unique `'id'` for commands inside the session, but different sessions might use the same ids.
- [Line 26](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L26): the `"flatten"` flag is the preffered mode of operation; the non-flattened mode will be removed eventually. Flattened mode allows us to pass `sessionId` as a part of the message (line 32).
- [Line 32](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L32): include the `sessionId` of the page as a part of the message to send it to the page.
Some commands set state which is stored per session, e.g. `Runtime.enable` and `Targets.setDiscoverTargets`. Each session is initialized with a set of domains, the exact set depends on the attached target and can be [found somewhere in the Chromium source](https://source.chromium.org/search?q=%22session-%3ECreateAndAddHandler%22%20f:devtools&ss=chromium). For example, sessions connected to a browser don't have a "Page" domain, but sessions connected to pages do.

We call sessions attached to a Browser target browser sessions. Similarly, there are page sessions, worker sessions and so on. In fact, the WebSocket connection is an implicitly created browser session.

When a client connects over the WebSocket to the launched Chromium browser (sessions.js:10), a root browser session is created.
This session is the one that receives commands if there's no `sessionId` specified ([sessions.js:14](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L14-L17)). Later on, when the root browser session is used to attach to a page target ([sessions.js:21](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/sessions.js#L21-L28)), a new page session created.

The page session is created from inside the browser session and thus is a child of the browser session. When a parent session closes, e.g. via [`Target.detachFromTarget`](https://vanilla.aslushnikov.com/#Target.detachFromTarget), all of its child sessions are closed as well.

The Chrome DevTools Protocol has stable and experimental parts. Events, methods, and sometimes whole domains
might be marked as experimental. DevTools team doesn't commit to maintaining experimental APIs and changes/removes them regularly.

!!! USE EXPERIMENTAL APIS AT YOUR OWN RISK !!!

As history has shown, experimental APIs do change quite often. If possible, stick to the stable protocol or use [Puppeteer](https://github.com/puppeteer/puppeteer).

NOTE: The Chrome DevTools team maintains [Puppeteer](https://github.com/puppeteer/puppeteer) as a reliable high-level API to control a browser. Internally, Puppeteer does use experimental CDP methods, but the team makes sure to update the library as the underlying protocol changes.

[Vanilla protocol viewer](https://vanilla.aslushnikov.com/) aggressively highlights experimental bits with red background.

It is very convenient to use Puppeteer to experiment with the raw protocol.
The following example creates a raw protocol session to the page to speed up animations.

File: ./cdpsession.js

```
const puppeteer = require('puppeteer');

(async() => {
  // Use Puppeteer to launch a browser and open a page.
  const browser = await puppeteer.launch({headless: false});
  const page = await browser.newPage();

  // Create a raw DevTools protocol session to talk to the page.
  // Use CDP to set the animation playback rate.
  const session = await page.target().createCDPSession();
  await session.send('Animation.enable');
  session.on('Animation.animationCreated', () => {
    console.log('Animation created!');
  });
  await session.send('Animation.setPlaybackRate', {
    playbackRate: 2,
  });

  // Check it out! Fast animations on the "loading..." screen!
  await page.goto('https://pptr.dev');
})();
```

It's easy to monitor all CDP messages that Puppeteer exchanges with Chromium.

```
# Use DEBUG env variable to dump CDP traffic.
$ DEBUG=*protocol node simple.js
```

You can also monitor CDP messages from DevTools: [Chrome DevTools Protocol Monitor](https://umaar.com/dev-tips/166-protocol-monitor/).

### An introduction to the Chrome DevTools Protocol

Published by Reflect, a third party. Source page: [An introduction to the Chrome DevTools Protocol](https://reflect.run/articles/introduction-to-chrome-devtools-protocol/).

[Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/) (CDP) is a set of APIs that allows
developers to communicate with Chromium-based browsers, including Google Chrome. CDP was originally developed to power
the Developer Tools features within Chrome, but since its introduction its usage has extended to much more than this
initial use-case.

In this article, we’ll provide some practical examples for interacting with the Chrome DevTools Protocol, as well as
cover how some popular testing libraries utilize CDP.

The code examples in this article assume that the following programs are installed on your development machine:

- Python3
- Node
- npm
- Chrome or another Chromium-based browser

Chrome DevTools Protocol is divided into domains. Each domain has a set of commands and events that it supports.

For example, the Network domain contains APIs for accessing the HTTP requests and responses made when rendering a page.

Another useful domain is the DOM (Document Object Model) domain. It exposes APIs for reading from and writing to the
DOM. You can access query selectors, get element attributes, manipulate nodes and even scroll to selected nodes.

Apart from powering Developer Tools in Chrome, Chrome DevTools Protocol provides some of the underlying functionality
used in popular testing libraries like Playwright, Puppeteer, and Selenium.

We can execute a page in Chrome headless mode and use Chrome DevTools APIs to debug. Chrome DevTools Protocol works with
any language that supports WebSockets.

To get started, install `PyChromeDevTools`. `PyChromeDevTools` is a library that provides wrappers for events, types,
and commands specified in Chrome DevTools Protocol.

```
pip3 install PyChromeDevTools
```

Next, run Chrome in headless mode:

```
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=5678 --disable-gpu --headless
```

Once this command completes, you are ready to start interacting with the browser through CDP.

In this first example, we’ll write a Python script that navigates to a page and waits until it has been loaded
successfully:

```
import PyChromeDevTools
import time

chrome = PyChromeDevTools.ChromeInterface()
chrome.Network.enable()
chrome.Page.enable()

start_time=time.time()

chrome.Page.navigate(url="https://example.com/")

# Wait for Page.loadEventFired before continuing  program
chrome.wait_event("Page.loadEventFired", timeout=60)

end_time=time.time()

print("Page Loading Time:", end_time-start_time)
```

After enabling the `Network` and `Page` domain, we navigate to `[example.com](https://example.com)`, wait for the
[`Page.loadEventFired`](https://chromedevtools.github.io/devtools-protocol/tot/Page/#event-loadEventFired) event to be
sent, and finally measure the time it took to load the page.

The page loads in about 1.6 seconds.

In this next example, we’ll use the `getCookies()` command in the Network domain to extract the cookies associated with
a page.

```
import PyChromeDevTools
import time

chrome = PyChromeDevTools.ChromeInterface()
chrome.Network.enable()
chrome.Page.enable()

chrome.Page.navigate(url="https://www.google.com")

chrome.wait_event("Page.frameStoppedLoading", timeout=60)

cookies, messages = chrome.Network.getCookies()

print(cookies["result"]["cookies"])
```

There are many other useful commands in CDP that we won’t be able to cover in this article. Check the
[official documentation](https://chromedevtools.github.io/devtools-protocol/) for more information.

Puppeteer is a browser automation tool that runs on NodeJS. It provides an API for creating automated tests and scripts
that use the Chrome DevTools Protocol under the covers.

Since interacting with the low-level Chrome DevTools Protocol can be tedious, using a higher-level library like
Puppeteer to do most of the heavy lifting can be a great time-saver. Below is an architectural diagram describing how
Puppeteer works (credit: [devdocs.io](https://devdocs.io/puppeteer/))

Puppeteer ships in two packages:

- `puppeteer-core` is the main library that handles all communications with Chrome DevTools Protocol APIs.
- `puppeteer` downloads and installs a version of Chromium and uses the `puppeteer-core` library to interact with the
browser.
If you’re building a library or another end-user product where there’s no need to download another Chromium binary, it’s
better to use `puppeteer-core`.

Pretty much anything you can manually do in your browser can be automated with `puppeteer`. This includes:

- Form field entry
- Clicks / Taps
- Page navigation
- Extracting text displayed on the page
In addition to replicating actions that a user can perform in the browser, Puppeteer can also perform actions that are
included as part of Chrome Developer Tools. This includes:

- Generate screenshots and PDFs of pages.
- Recording load time and runtime performances.
- Emulating various mobile devices, including using their proper user agent, device dimensions, and pixel density.
One exciting feature of Puppeteer is utilizing headless mode to enable server-side rendering (SSR). Most search engines
rely on static HTML to index content, while more javascript-centric applications are getting created. Prerendering pages
using headless Chrome and Puppeteer is a great way to generate static HTML pages.

Install Puppeteer and Puppeteer-Core:

```
npm i puppeteer
npm i puppeteer-core
```

This script will use `puppeteer` to extract all anchor tags on a page.

```
const puppeteer = require("puppeteer");

(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  await page.goto("https://reflect.run/");

  const hrefs = await page.evaluate(() => {
    let links = [];

    let elements2 = document.querySelectorAll("a");

    for (let element2 of elements2) links.push(element2.href);

    return links;
  });

  console.log(hrefs);

  await browser.close();
})();
```

Run the script:

```
node puppeteer-script.js
```

The output is an array of all the hypertext references on the page.

Microsoft released the public version of Playwright in July 2020. Playwright is similar to Puppeteer in many ways, and
that’s not a coincidence: it was developed at Microsoft by the same team that initially developed Puppeteer at Google.

Playwright also uses Chrome DevTools Protocol to interact with Chromium-based browsers. One exciting feature in
Playwright is `BrowserContexts`. `BrowserContexts` lets you operate many independent browser sessions. If a page opens
another window, that page gets added to the parent context. A browser context can have multiple pages(tabs).

If you don’t want to use the high-level methods provided by Playwright, you can use `CDPSession` to directly interact
with Chrome DevTools Protocol.

Install Playwright using pip:

```
pip3 install playwright
playwright install
```

Let’s run a script to screenshot a website on an iPhone 12 device:

```
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch()
    iphone = p.devices["iPhone 12"]

    context = browser.new_context(**iphone)

    page = context.new_page()
    page.goto("https://reflect.run/")
    page.screenshot(path="example.png")

    browser.close()
```

Run the script:

```
python3 chrome_script.py
```

The script will create a new example.png file in the directory.

We can also monitor console output with Playwright.

```
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch()

    page = browser.new_page()
    page.on("console", lambda a: print(a))
    page.evaluate("console.log('hello', 5, {foo: 'bar'})")

    page.goto("https://reflect.run/")
```

Run the script:

```
python3 chrome_script.py
```

Selenium uses the WebDriver Protocol. [WebDriver](https://www.w3.org/TR/webdriver/) provides a set of interfaces to
manipulate DOM elements in web documents and user agents. For this to work, an intermediary server is required. To test
across different browsers, you’ll need separate drivers:

- ChromeDriver for Chrome
- GeckoDriver for Firefox
- SafariDriver for Safari
Version 4 of Selenium includes a new protocol called
[WebDriver BiDi](https://www.selenium.dev/documentation/webdriver/bidirectional/). The WebDriver BiDi (short for
“bi-directional”) interface is in its early stages, but its purpose is to provide a stable bidirectional API for
cross-browser automation and testing. Right now WebDriver BiDi is simply a wrapper around a subset of the Chrome
DevTools Protocol, however there is an effort underway to define a [W3C spec](https://w3c.github.io/webdriver-bidi/) for
WebDriver BiDi and have other browser vendors implement that spec.

Chrome DevTools Protocol makes it possible to have the same set of tests work across Chromium-based browsers. Besides
browser test automation, Chrome DevTools Protocol can also help with server-side rendering.

Although it’s powerful, it can be cumbersome to interact directly with the Chrome DevTools Protocol. Using tools like
Puppeteer, Playwright, and Selenium which provide higher-level abstractions over CDP is something you should consider
before deciding to interact directly with this protocol.

## Libraries and bindings

### chrome-remote-interface recipes

Published by cyrus-and, project wiki. Source page: [chrome-remote-interface recipes](https://github.com/cyrus-and/chrome-remote-interface/wiki).

This wiki is meant to collect the most common use cases and frequently asked questions in order to provide newcomers some boilerplate code to start with, hereafter called recipes. For a detailed API reference refer to the [README](https://github.com/cyrus-and/chrome-remote-interface/blob/master/README.md) instead.

All the available recipes will appear in the sidebar.

Each recipe must be placed in a separate Markdown file identified by a meaningful title.

Code recipes must adhere to a common format:

````
{description}

```js
const CDP = require('chrome-remote-interface');

{implementation}
```

- {optional external reference}...
````

The code must be able to run as is, that is, it must provide any additional requires and must not rely on external files.

### PyCDP overview

Published by PyCDP. Source page: [PyCDP overview](https://py-cdp.readthedocs.io/en/latest/overview.html).

Python Chrome DevTools Protocol (shortened to PyCDP) is a library that provides
Python wrappers for the types, commands, and events specified in the [Chrome
DevTools Protocol](https://github.com/ChromeDevTools/devtools-protocol/).

The Chrome DevTools Protocol provides for remote control of a web browser by
sending JSON messages over a WebSocket. That JSON format is described by a
machine-readable specification. This specification is used to automatically
generate the classes and methods found in this library.

You could write a CDP client by connecting a WebSocket and then sending JSON
objects, but this would be tedious and error-prone: the Python interpreter would
not catch any typos in your JSON objects, and you wouldn’t get autocomplete for
any parts of the JSON data structure. By providing a set of native Python
wrappers, this project makes it easier and faster to write CDP client code.

Two usage modes are available:

- Sans-I/O mode (original): The core library provides type wrappers without
performing any I/O. This maximises flexibility and allows integration with any
async framework such as
[trio-chrome-devtools-protocol](https://github.com/hyperiongray/trio-chrome-devtools-protocol).
- I/O mode: The `cdp.connection` module handles the WebSocket lifecycle,
JSON-RPC message framing, and command multiplexing. The
`cdp.browser_control` module layers a high-level automation API on top of
this, providing Playwright-style helpers for navigation, element interaction,
screenshots, and more. See [Browser Control](browser_control.html) for details.
This package provides Chrome DevTools Protocol r678025. Download a compatible
Chrome package:

- [Linux](https://storage.googleapis.com/chromium-browser-snapshots/Linux_x64/678025/chrome-linux.zip)
- [Mac](https://storage.googleapis.com/chromium-browser-snapshots/Mac/678025/chrome-mac.zip)
- [Windows 32-bit](https://storage.googleapis.com/chromium-browser-snapshots/Win/678025/chrome-win.zip)
- [Windows 64-bit](https://storage.googleapis.com/chromium-browser-snapshots/Win_x64/678025/chrome-win.zip)
Install from PyPI (requires Python ≥3.8):

```
$ pip install chrome-devtools-protocol        # Sans-I/O only
$ pip install chrome-devtools-protocol[io]    # With WebSocket / browser-control support
```

Quick example (Sans-I/O mode):

```
from cdp import page

frame_id = page.FrameId('my id')
assert repr(frame_id) == "FrameId('my id')"
```

Quick example (browser control):

```
import asyncio
from cdp.connection import CDPConnection
from cdp import browser_control as bc
from cdp import page

async def main():
    async with CDPConnection("ws://localhost:9222/devtools/page/ID") as conn:
        await conn.execute(page.enable())
        await bc.navigate(conn, "https://example.com")
        await bc.wait_for_load(conn)
        print(await bc.get_text(conn, "h1"))

asyncio.run(main())
```

## Playwright

### BrowserType and connect_over_cdp in Playwright for Python

Published by Playwright. Source page: [BrowserType and connect_over_cdp in Playwright for Python](https://playwright.dev/python/docs/api/class-browsertype).

BrowserType provides methods to launch a specific browser instance or connect to an existing one. The following is a typical example of using Playwright to drive automation:

- Sync
- Async

```
from playwright.sync_api import sync_playwright, Playwright

def run(playwright: Playwright):
    chromium = playwright.chromium
    browser = chromium.launch()
    page = browser.new_page()
    page.goto("https://example.com")
    # other actions...
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
```

```
import asyncio
from playwright.async_api import async_playwright, Playwright

async def run(playwright: Playwright):
    chromium = playwright.chromium
    browser = await chromium.launch()
    page = await browser.new_page()
    await page.goto("https://example.com")
    # other actions...
    await browser.close()

async def main():
    async with async_playwright() as playwright:
        await run(playwright)
asyncio.run(main())
```

Added before v1.9
browserType.connect

This method attaches Playwright to an existing browser instance created via `BrowserType.launchServer` in Node.js.

note

The major and minor version of the Playwright instance that connects needs to match the version of Playwright that launches the browser (1.2.3 → is compatible with 1.2.x).

Usage

```
browser_type.connect(endpoint)
browser_type.connect(endpoint, **kwargs)
```

Arguments

- `endpoint` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) Added in: v1.10#
- A Playwright browser websocket endpoint to connect to. You obtain this endpoint via `BrowserServer.wsEndpoint`.
- `expose_network` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional) Added in: v1.37#
- This option exposes network available on the connecting client to the browser being connected to. Consists of a list of rules separated by comma.
- Available rules:
    - Hostname pattern, for example: `example.com`, `*.org:99`, `x.*.y.com`, `*foo.org`.
    - IP literal, for example: `127.0.0.1`, `0.0.0.0:99`, `[::1]`, `[0:0::1]:99`.
    - `<loopback>` that matches local loopback interfaces: `localhost`, `*.localhost`, `127.0.0.1`, `[::1]`.
- Some common examples:
    - `"*"` to expose all network.
    - `"<loopback>"` to expose localhost network.
    - `"*.test.internal-domain,*.staging.internal-domain,<loopback>"` to expose test/staging deployments and localhost.
- `headers` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional) Added in: v1.11#
- Additional HTTP headers to be sent with web socket connect request. Optional.
- `slow_mo` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional) Added in: v1.10#
- Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
- `timeout` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional) Added in: v1.10#
- Maximum time in milliseconds to wait for the connection to be established. Defaults to `0` (no timeout).
Returns

- Browser#

Added in: v1.9
browserType.connect_over_cdp

This method attaches Playwright to an existing browser instance using the Chrome DevTools Protocol.

The default browser context is accessible via browser.contexts.

note

Connecting over the Chrome DevTools Protocol is only supported for Chromium-based browsers.

note

This connection is significantly lower fidelity than the Playwright protocol connection via browser_type.connect(). If you are experiencing issues or attempting to use advanced functionality, you probably want to use browser_type.connect().

warning

Playwright maintains a curated list of arguments for launching the browser. If you launch the browser without Playwright and do not pass the exact same arguments, some of Playwright functionality may be broken upon connecting to the browser.

Usage

- Sync
- Async

```
browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
default_context = browser.contexts[0]
page = default_context.pages[0]
```

```
browser = await playwright.chromium.connect_over_cdp("http://localhost:9222")
default_context = browser.contexts[0]
page = default_context.pages[0]
```

Arguments

- `endpoint_url` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) Added in: v1.11#
- A CDP websocket endpoint or http url to connect to. For example `[localhost:9222](http://localhost:9222/)` or `ws://127.0.0.1:9222/devtools/browser/387adf4c-243f-4051-a181-46798f4a46f4`.
- `artifacts_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional) Added in: v1.61#
- If specified, browser artifacts (such as traces and downloads) are saved into this directory.
- `headers` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional) Added in: v1.11#
- Additional HTTP headers to be sent with connect request. Optional.
- `is_local` [bool](https://docs.python.org/3/library/stdtypes.html) (optional) Added in: v1.58#
- Tells Playwright that it runs on the same host as the CDP server. It will enable certain optimizations that rely upon the file system being the same between Playwright and the Browser.
- `no_defaults` [bool](https://docs.python.org/3/library/stdtypes.html) (optional) Added in: v1.60#
- When true, Playwright will not apply its default overrides to the existing default browser context. Specifically, accept_downloads is left at the browser's setting, focus emulation is not enabled, and media emulation options (such as color_scheme, reduced_motion, forced_colors, and contrast) are not applied. Useful when attaching to a user's daily-driver browser where these overrides would interfere with existing browser state. New contexts created via browser.new_context() are not affected. Defaults to `false`.
- `slow_mo` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional) Added in: v1.11#
- Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
- `timeout` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional) Added in: v1.11#
- Maximum time in milliseconds to wait for the connection to be established. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
Returns

- Browser#

Added before v1.9
browserType.launch

Returns the browser instance.

Usage

You can use ignore_default_args to filter out `--mute-audio` from default arguments:

- Sync
- Async

```
browser = playwright.chromium.launch( # or "firefox" or "webkit".
    ignore_default_args=["--mute-audio"]
)
```

```
browser = await playwright.chromium.launch( # or "firefox" or "webkit".
    ignore_default_args=["--mute-audio"]
)
```

Chromium-only Playwright can also be used to control the Google Chrome or Microsoft Edge browsers, but it works best with the version of Chromium it is bundled with. There is no guarantee it will work with any other version. Use executable_path option with extreme caution.

If Google Chrome (rather than Chromium) is preferred, a [Chrome Canary](https://www.google.com/chrome/browser/canary.html) or [Dev Channel](https://www.chromium.org/getting-involved/dev-channel) build is suggested.

Stock browsers like Google Chrome and Microsoft Edge are suitable for tests that require proprietary media codecs for video playback. See [this article](https://www.howtogeek.com/202825/what%E2%80%99s-the-difference-between-chromium-and-chrome/) for other differences between Chromium and Chrome. [This article](https://chromium.googlesource.com/chromium/src/+/lkgr/docs/chromium_browser_vs_google_chrome.md) describes some differences for Linux users.

Arguments

- `args` [List](https://docs.python.org/3/library/typing.html#typing.List)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- warning
- Use custom browser args at your own risk, as some of them may break Playwright functionality.
- Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
- `artifacts_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
- `channel` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- Browser distribution channel.
- Use "chromium" to opt in to new headless mode.
- Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
- `chromium_sandbox` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Enable Chromium sandboxing. Defaults to `false`.
- `downloads_path` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
- `env` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) | [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) | [bool](https://docs.python.org/3/library/stdtypes.html)] (optional)#
- Specify environment variables that will be visible to the browser. Defaults to `process.env`.
- `executable_path` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- Path to a browser executable to run instead of the bundled one. If executable_path is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
- `firefox_user_prefs` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) | [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) | [bool](https://docs.python.org/3/library/stdtypes.html)] (optional)#
- Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
- You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
- `handle_sighup` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on SIGHUP. Defaults to `true`.
- `handle_sigint` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on Ctrl-C. Defaults to `true`.
- `handle_sigterm` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on SIGTERM. Defaults to `true`.
- `headless` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
- `ignore_default_args` [bool](https://docs.python.org/3/library/stdtypes.html) | [List](https://docs.python.org/3/library/typing.html#typing.List)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- If `true`, Playwright does not pass its own configurations args and only uses the ones from args. If an array is given, then filters out the given default arguments. Dangerous option; use with care. Defaults to `false`.
- `proxy` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `server` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)
    - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
    - `bypass` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
    - `username` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional username to use if HTTP proxy requires authentication.
    - `password` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional password to use if HTTP proxy requires authentication.
- Network proxy settings.
- `slow_mo` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)#
- Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
- `timeout` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)#
- Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
- `traces_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, traces are saved into this directory.
Returns

- Browser#

Added before v1.9
browserType.launch_persistent_context

Returns the persistent browser context instance.

Launches browser that uses persistent storage located at user_data_dir and returns the only context. Closing this context will automatically close the browser.

Usage

```
browser_type.launch_persistent_context(user_data_dir)
browser_type.launch_persistent_context(user_data_dir, **kwargs)
```

Arguments

- `user_data_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)]#
- Path to a User Data Directory, which stores browser session data like cookies and local storage. Pass an empty string to create a temporary directory.
- More details for [Chromium](https://chromium.googlesource.com/chromium/src/+/master/docs/user_data_dir.md#introduction) and [Firefox](https://wiki.mozilla.org/Firefox/CommandLineOptions#User_profile). Chromium's user data directory is the parent directory of the "Profile Path" seen at `chrome://version`.
- Note that browsers do not allow launching multiple instances with the same User Data Directory.
- warning
- Chromium/Chrome: Due to recent Chrome policy changes, automating the default Chrome user profile is not supported. Pointing `userDataDir` to Chrome's main "User Data" directory (the profile used for your regular browsing) may result in pages not loading or the browser exiting. Create and use a separate directory (for example, an empty folder) as your automation profile instead. See [developer.chrome.com](https://developer.chrome.com/blog/remote-debugging-port) for details.
- `accept_downloads` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether to automatically download all the attachments. Defaults to `true` where all the downloads are accepted.
- `args` [List](https://docs.python.org/3/library/typing.html#typing.List)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- warning
- Use custom browser args at your own risk, as some of them may break Playwright functionality.
- Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
- `artifacts_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
- `base_url` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- When using page.goto(), page.route(), page.wait_for_url(), page.expect_request(), or page.expect_response() it takes the base URL in consideration by using the [`URL()`](https://developer.mozilla.org/en-US/docs/Web/API/URL/URL) constructor for building the corresponding URL. Unset by default. Examples:
    - baseURL: `[localhost:3000](http://localhost:3000)` and navigating to `/bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
    - baseURL: `[localhost:3000](http://localhost:3000/foo/)` and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/foo/bar.html)`
    - baseURL: `[localhost:3000](http://localhost:3000/foo)` (without trailing slash) and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
- `bypass_csp` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Toggles bypassing page's Content-Security-Policy. Defaults to `false`.
- `channel` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- Browser distribution channel.
- Use "chromium" to opt in to new headless mode.
- Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
- `chromium_sandbox` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Enable Chromium sandboxing. Defaults to `false`.
- `client_certificates` [List](https://docs.python.org/3/library/typing.html#typing.List)[[Dict](https://docs.python.org/3/library/typing.html#typing.Dict)] (optional) Added in: 1.46#
    - `origin` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)
    - Exact origin that the certificate is valid for. Origin includes `https` protocol, a hostname and optionally a port.
    - `certPath` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)
    - Path to the file with the certificate in PEM format.
    - `cert` [bytes](https://docs.python.org/3/library/stdtypes.html#bytes) (optional)
    - Direct value of the certificate in PEM format.
    - `keyPath` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)
    - Path to the file with the private key in PEM format.
    - `key` [bytes](https://docs.python.org/3/library/stdtypes.html#bytes) (optional)
    - Direct value of the private key in PEM format.
    - `pfxPath` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)
    - Path to the PFX or PKCS12 encoded private key and certificate chain.
    - `pfx` [bytes](https://docs.python.org/3/library/stdtypes.html#bytes) (optional)
    - Direct value of the PFX or PKCS12 encoded private key and certificate chain.
    - `passphrase` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Passphrase for the private key (PEM or PFX).
- TLS Client Authentication allows the server to request a client certificate and verify it.
- Details
- An array of client certificates to be used. Each certificate object must have either both `certPath` and `keyPath`, a single `pfxPath`, or their corresponding direct value equivalents (`cert` and `key`, or `pfx`). Optionally, `passphrase` property should be provided if the certificate is encrypted. The `origin` property should be provided with an exact match to the request origin that the certificate is valid for.
- Client certificate authentication is only active when at least one client certificate is provided. If you want to reject all client certificates sent by the server, you need to provide a client certificate with an `origin` that does not match any of the domains you plan to visit.
- note
- When using WebKit on macOS, accessing `localhost` will not pick up client certificates. You can make it work by replacing `localhost` with `local.playwright`.
- `color_scheme` "light" | "dark" | "no-preference" | "null" (optional)#
- Emulates [prefers-colors-scheme](https://developer.mozilla.org/en-US/docs/Web/CSS/@media/prefers-color-scheme) media feature, supported values are `'light'` and `'dark'`. See page.emulate_media() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'light'`.
- `contrast` "no-preference" | "more" | "null" (optional)#
- Emulates `'prefers-contrast'` media feature, supported values are `'no-preference'`, `'more'`. See page.emulate_media() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'no-preference'`.
- `device_scale_factor` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)#
- Specify device scale factor (can be thought of as dpr). Defaults to `1`. Learn more about emulating devices with device scale factor.
- `downloads_path` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
- `env` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) | [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) | [bool](https://docs.python.org/3/library/stdtypes.html)] (optional)#
- Specify environment variables that will be visible to the browser. Defaults to `process.env`.
- `executable_path` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- Path to a browser executable to run instead of the bundled one. If executable_path is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
- `extra_http_headers` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- An object containing additional HTTP headers to be sent with every request. Defaults to none.
- `firefox_user_prefs` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) | [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) | [bool](https://docs.python.org/3/library/stdtypes.html)] (optional) Added in: v1.40#
- Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
- You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
- `forced_colors` "active" | "none" | "null" (optional)#
- Emulates `'forced-colors'` media feature, supported values are `'active'`, `'none'`. See page.emulate_media() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'none'`.
- `geolocation` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `latitude` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - Latitude between -90 and 90.
    - `longitude` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - Longitude between -180 and 180.
    - `accuracy` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)
    - Non-negative accuracy value. Defaults to `0`.
- `handle_sighup` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on SIGHUP. Defaults to `true`.
- `handle_sigint` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on Ctrl-C. Defaults to `true`.
- `handle_sigterm` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Close the browser process on SIGTERM. Defaults to `true`.
- `has_touch` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Specifies if viewport supports touch events. Defaults to false. Learn more about mobile emulation.
- `headless` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
- `http_credentials` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `username` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)
    - `password` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)
    - `origin` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Restrain sending http credentials on specific origin (scheme://host:port).
    - `send` "unauthorized" | "always" (optional)
    - This option only applies to the requests sent from corresponding APIRequestContext and does not affect requests sent from the browser. `'always'` - `Authorization` header with basic authentication credentials will be sent with the each API request. `'unauthorized` - the credentials are only sent when 401 (Unauthorized) response with `WWW-Authenticate` header is received. Defaults to `'unauthorized'`.
- Credentials for [HTTP authentication](https://developer.mozilla.org/en-US/docs/Web/HTTP/Authentication). If no origin is specified, the username and password are sent to any servers upon unauthorized responses.
- `ignore_default_args` [bool](https://docs.python.org/3/library/stdtypes.html) | [List](https://docs.python.org/3/library/typing.html#typing.List)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- If `true`, Playwright does not pass its own configurations args and only uses the ones from args. If an array is given, then filters out the given default arguments. Dangerous option; use with care. Defaults to `false`.
- `ignore_https_errors` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether to ignore HTTPS errors when sending network requests. Defaults to `false`.
- `is_mobile` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether the `meta viewport` tag is taken into account and touch events are enabled. isMobile is a part of device, so you don't actually need to set it manually. Defaults to `false` and is not supported in Firefox. Learn more about mobile emulation.
- `java_script_enabled` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether or not to enable JavaScript in the context. Defaults to `true`. Learn more about disabling JavaScript.
- `locale` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- Specify user locale, for example `en-GB`, `de-DE`, etc. Locale will affect `navigator.language` value, `Accept-Language` request header value as well as number and date formatting rules. Defaults to the system default locale. Learn more about emulation in our emulation guide.
- `no_viewport` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Does not enforce fixed viewport, allows resizing window in the headed mode.
- `offline` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Whether to emulate network being offline. Defaults to `false`. Learn more about network emulation.
- `permissions` [List](https://docs.python.org/3/library/typing.html#typing.List)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)] (optional)#
- A list of permissions to grant to all pages in this context. See browser_context.grant_permissions() for more details. Defaults to none.
- `proxy` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `server` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)
    - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
    - `bypass` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
    - `username` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional username to use if HTTP proxy requires authentication.
    - `password` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)
    - Optional password to use if HTTP proxy requires authentication.
- Network proxy settings.
- `record_har_content` "omit" | "embed" | "attach" (optional)#
- Optional setting to control resource content management. If `omit` is specified, content is not persisted. If `attach` is specified, resources are persisted as separate files and all of these files are archived along with the HAR file. Defaults to `embed`, which stores content inline the HAR file as per HAR specification.
- `record_har_mode` "full" | "minimal" (optional)#
- When set to `minimal`, only record information necessary for routing from HAR. This omits sizes, timing, page, cookies, security and other types of HAR information that are not used when replaying from HAR. Defaults to `full`.
- `record_har_omit_content` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- Optional setting to control whether to omit request content from the HAR. Defaults to `false`.
- `record_har_path` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- Enables [HAR](http://www.softwareishard.com/blog/har-12-spec) recording for all pages into the specified HAR file on the filesystem. If not specified, the HAR is not recorded. Make sure to call browser_context.close() for the HAR to be saved.
- `record_har_url_filter` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) | [Pattern](https://docs.python.org/3/library/re.html) (optional)#
- `record_video_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- Enables video recording for all pages into the specified directory. If not specified videos are not recorded. Make sure to call browser_context.close() for videos to be saved.
- `record_video_size` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `width` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - Video frame width.
    - `height` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - Video frame height.
- Dimensions of the recorded videos. If not specified the size will be equal to `viewport` scaled down to fit into 800x800. If `viewport` is not configured explicitly the video size defaults to 800x450. Actual picture of each page will be scaled down if necessary to fit the specified size.
- `reduced_motion` "reduce" | "no-preference" | "null" (optional)#
- Emulates `'prefers-reduced-motion'` media feature, supported values are `'reduce'`, `'no-preference'`. See page.emulate_media() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'no-preference'`.
- `screen` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `width` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - page width in pixels.
    - `height` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - page height in pixels.
- Emulates consistent window screen size available inside web page via `window.screen`. Is only used when the viewport is set.
- `service_workers` "allow" | "block" (optional)#
- Whether to allow sites to register Service workers. Defaults to `'allow'`.
    - `'allow'`: [Service Workers](https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API) can be registered.
    - `'block'`: Playwright will block all registration of Service Workers.
- `slow_mo` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)#
- Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
- `strict_selectors` [bool](https://docs.python.org/3/library/stdtypes.html) (optional)#
- If set to true, enables strict selectors mode for this context. In the strict selectors mode all operations on selectors that imply single target DOM element will throw when more than one element matches the selector. This option does not affect any Locator APIs (Locators are always strict). Defaults to `false`. See Locator to learn more about the strict mode.
- `timeout` [float](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex) (optional)#
- Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
- `timezone_id` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- Changes the timezone of the context. See [ICU's metaZones.txt](https://cs.chromium.org/chromium/src/third_party/icu/source/data/misc/metaZones.txt?rcl=faee8bc70570192d82d2978a71e2a615788597d1) for a list of supported timezone IDs. Defaults to the system timezone.
- `traces_dir` [Union](https://docs.python.org/3/library/typing.html#typing.Union)[[str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str), [pathlib.Path](https://realpython.com/python-pathlib/)] (optional)#
- If specified, traces are saved into this directory.
- `user_agent` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str) (optional)#
- Specific user agent to use in this context.
- `viewport` [NoneType](https://docs.python.org/3/library/constants.html#None) | [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
    - `width` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - page width in pixels.
    - `height` [int](https://docs.python.org/3/library/stdtypes.html#numeric-types-int-float-complex)
    - page height in pixels.
- Sets a consistent viewport for each page. Defaults to an 1280x720 viewport. `no_viewport` disables the fixed viewport. Learn more about viewport emulation.
Returns

- BrowserContext#

Added before v1.9
browserType.executable_path

A path where Playwright expects to find a bundled browser executable.

Usage

```
browser_type.executable_path
```

Returns

- [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)#

Added before v1.9
browserType.name

Returns browser name. For example: `'chromium'`, `'webkit'` or `'firefox'`.

Usage

```
browser_type.name
```

Returns

- [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)#

### BrowserType and connectOverCDP in Playwright for JavaScript

Published by Playwright. Source page: [BrowserType and connectOverCDP in Playwright for JavaScript](https://playwright.dev/docs/api/class-browsertype).

BrowserType provides methods to launch a specific browser instance or connect to an existing one. The following is a typical example of using Playwright to drive automation:

```
const { chromium } = require('playwright');  // Or 'firefox' or 'webkit'.

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('https://example.com');
  // other actions...
  await browser.close();
})();
```

Added before v1.9
browserType.connect

This method attaches Playwright to an existing browser instance created via `BrowserType.launchServer` in Node.js.

note

The major and minor version of the Playwright instance that connects needs to match the version of Playwright that launches the browser (1.2.3 → is compatible with 1.2.x).

Usage

```
await browserType.connect(endpoint);
await browserType.connect(endpoint, options);
```

Arguments

- `endpoint` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) Added in: v1.10#
- A Playwright browser websocket endpoint to connect to. You obtain this endpoint via `BrowserServer.wsEndpoint`.
- `options` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - `exposeNetwork` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional) Added in: v1.37#
    - This option exposes network available on the connecting client to the browser being connected to. Consists of a list of rules separated by comma.
    - Available rules:
        - Hostname pattern, for example: `example.com`, `*.org:99`, `x.*.y.com`, `*foo.org`.
        - IP literal, for example: `127.0.0.1`, `0.0.0.0:99`, `[::1]`, `[0:0::1]:99`.
        - `<loopback>` that matches local loopback interfaces: `localhost`, `*.localhost`, `127.0.0.1`, `[::1]`.
    - Some common examples:
        - `"*"` to expose all network.
        - `"<loopback>"` to expose localhost network.
        - `"*.test.internal-domain,*.staging.internal-domain,<loopback>"` to expose test/staging deployments and localhost.
    - `headers` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional) Added in: v1.11#
    - Additional HTTP headers to be sent with web socket connect request. Optional.
    - `slowMo` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional) Added in: v1.10#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
    - `timeout` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional) Added in: v1.10#
    - Maximum time in milliseconds to wait for the connection to be established. Defaults to `0` (no timeout).
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<Browser>#

Added in: v1.9
browserType.connectOverCDP

This method attaches Playwright to an existing browser instance using the Chrome DevTools Protocol.

The default browser context is accessible via browser.contexts().

note

Connecting over the Chrome DevTools Protocol is only supported for Chromium-based browsers.

note

This connection is significantly lower fidelity than the Playwright protocol connection via browserType.connect(). If you are experiencing issues or attempting to use advanced functionality, you probably want to use browserType.connect().

warning

Playwright maintains a curated list of arguments for launching the browser. If you launch the browser without Playwright and do not pass the exact same arguments, some of Playwright functionality may be broken upon connecting to the browser.

Usage

```
const browser = await playwright.chromium.connectOverCDP('http://localhost:9222');
const defaultContext = browser.contexts()[0];
const page = defaultContext.pages()[0];
```

Arguments

- `endpointURL` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) Added in: v1.11#
- A CDP websocket endpoint or http url to connect to. For example `[localhost:9222](http://localhost:9222/)` or `ws://127.0.0.1:9222/devtools/browser/387adf4c-243f-4051-a181-46798f4a46f4`.
- `options` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - `artifactsDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional) Added in: v1.61#
    - If specified, browser artifacts (such as traces and downloads) are saved into this directory.
    - `endpointURL` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional) Added in: v1.14#
    - Deprecated
    - Use the first argument instead.
    - `headers` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional) Added in: v1.11#
    - Additional HTTP headers to be sent with connect request. Optional.
    - `isLocal` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional) Added in: v1.58#
    - Tells Playwright that it runs on the same host as the CDP server. It will enable certain optimizations that rely upon the file system being the same between Playwright and the Browser.
    - `noDefaults` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional) Added in: v1.60#
    - When true, Playwright will not apply its default overrides to the existing default browser context. Specifically, acceptDownloads is left at the browser's setting, focus emulation is not enabled, and media emulation options (such as colorScheme, reducedMotion, forcedColors, and contrast) are not applied. Useful when attaching to a user's daily-driver browser where these overrides would interfere with existing browser state. New contexts created via browser.newContext() are not affected. Defaults to `false`.
    - `slowMo` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional) Added in: v1.11#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
    - `timeout` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional) Added in: v1.11#
    - Maximum time in milliseconds to wait for the connection to be established. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<Browser>#

Added before v1.9
browserType.executablePath

A path where Playwright expects to find a bundled browser executable.

Usage

```
browserType.executablePath();
```

Returns

- [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)#

Added before v1.9
browserType.launch

Returns the browser instance.

Usage

You can use ignoreDefaultArgs to filter out `--mute-audio` from default arguments:

```
const browser = await chromium.launch({  // Or 'firefox' or 'webkit'.
  ignoreDefaultArgs: ['--mute-audio']
});
```

Chromium-only Playwright can also be used to control the Google Chrome or Microsoft Edge browsers, but it works best with the version of Chromium it is bundled with. There is no guarantee it will work with any other version. Use executablePath option with extreme caution.

If Google Chrome (rather than Chromium) is preferred, a [Chrome Canary](https://www.google.com/chrome/browser/canary.html) or [Dev Channel](https://www.chromium.org/getting-involved/dev-channel) build is suggested.

Stock browsers like Google Chrome and Microsoft Edge are suitable for tests that require proprietary media codecs for video playback. See [this article](https://www.howtogeek.com/202825/what%E2%80%99s-the-difference-between-chromium-and-chrome/) for other differences between Chromium and Chrome. [This article](https://chromium.googlesource.com/chromium/src/+/lkgr/docs/chromium_browser_vs_google_chrome.md) describes some differences for Linux users.

Arguments

- `options` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - `args` [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - warning
    - Use custom browser args at your own risk, as some of them may break Playwright functionality.
    - Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
    - `artifactsDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
    - `channel` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Browser distribution channel.
    - Use "chromium" to opt in to new headless mode.
    - Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
    - `chromiumSandbox` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Enable Chromium sandboxing. Defaults to `false`.
    - `downloadsPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
    - `env` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [undefined]> (optional)#
    - `executablePath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Path to a browser executable to run instead of the bundled one. If executablePath is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
    - `firefoxUserPrefs` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) | [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type)> (optional)#
    - Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
    - You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
    - `handleSIGHUP` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGHUP. Defaults to `true`.
    - `handleSIGINT` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on Ctrl-C. Defaults to `true`.
    - `handleSIGTERM` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGTERM. Defaults to `true`.
    - `headless` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
    - `ignoreDefaultArgs` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) | [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from args. If an array is given, then filters out the given default arguments. Dangerous option; use with care. Defaults to `false`.
    - `logger` Logger (optional)#
    - Deprecated
    - The logs received by the logger are incomplete. Please use tracing instead.
    - Logger sink for Playwright logging.
    - `proxy` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `server` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
        - `bypass` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
        - `username` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional username to use if HTTP proxy requires authentication.
        - `password` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional password to use if HTTP proxy requires authentication.
    - Network proxy settings.
    - `slowMo` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
    - `timeout` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
    - `tracesDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, traces are saved into this directory.
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<Browser>#

Added before v1.9
browserType.launchPersistentContext

Returns the persistent browser context instance.

Launches browser that uses persistent storage located at userDataDir and returns the only context. Closing this context will automatically close the browser.

Usage

```
await browserType.launchPersistentContext(userDataDir);
await browserType.launchPersistentContext(userDataDir, options);
```

Arguments

- `userDataDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)#
- Path to a User Data Directory, which stores browser session data like cookies and local storage. Pass an empty string to create a temporary directory.
- More details for [Chromium](https://chromium.googlesource.com/chromium/src/+/master/docs/user_data_dir.md#introduction) and [Firefox](https://wiki.mozilla.org/Firefox/CommandLineOptions#User_profile). Chromium's user data directory is the parent directory of the "Profile Path" seen at `chrome://version`.
- Note that browsers do not allow launching multiple instances with the same User Data Directory.
- warning
- Chromium/Chrome: Due to recent Chrome policy changes, automating the default Chrome user profile is not supported. Pointing `userDataDir` to Chrome's main "User Data" directory (the profile used for your regular browsing) may result in pages not loading or the browser exiting. Create and use a separate directory (for example, an empty folder) as your automation profile instead. See [developer.chrome.com](https://developer.chrome.com/blog/remote-debugging-port) for details.
- `options` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - `acceptDownloads` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to automatically download all the attachments. Defaults to `true` where all the downloads are accepted.
    - `args` [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - warning
    - Use custom browser args at your own risk, as some of them may break Playwright functionality.
    - Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
    - `artifactsDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
    - `baseURL` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - When using page.goto(), page.route(), page.waitForURL(), page.waitForRequest(), or page.waitForResponse() it takes the base URL in consideration by using the [`URL()`](https://developer.mozilla.org/en-US/docs/Web/API/URL/URL) constructor for building the corresponding URL. Unset by default. Examples:
        - baseURL: `[localhost:3000](http://localhost:3000)` and navigating to `/bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
        - baseURL: `[localhost:3000](http://localhost:3000/foo/)` and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/foo/bar.html)`
        - baseURL: `[localhost:3000](http://localhost:3000/foo)` (without trailing slash) and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
    - `bypassCSP` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Toggles bypassing page's Content-Security-Policy. Defaults to `false`.
    - `channel` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Browser distribution channel.
    - Use "chromium" to opt in to new headless mode.
    - Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
    - `chromiumSandbox` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Enable Chromium sandboxing. Defaults to `false`.
    - `clientCertificates` [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)> (optional) Added in: 1.46#
        - `origin` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - Exact origin that the certificate is valid for. Origin includes `https` protocol, a hostname and optionally a port.
        - `certPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Path to the file with the certificate in PEM format.
        - `cert` [Buffer](https://nodejs.org/api/buffer.html#buffer_class_buffer) (optional)
        - Direct value of the certificate in PEM format.
        - `keyPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Path to the file with the private key in PEM format.
        - `key` [Buffer](https://nodejs.org/api/buffer.html#buffer_class_buffer) (optional)
        - Direct value of the private key in PEM format.
        - `pfxPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Path to the PFX or PKCS12 encoded private key and certificate chain.
        - `pfx` [Buffer](https://nodejs.org/api/buffer.html#buffer_class_buffer) (optional)
        - Direct value of the PFX or PKCS12 encoded private key and certificate chain.
        - `passphrase` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Passphrase for the private key (PEM or PFX).
    - TLS Client Authentication allows the server to request a client certificate and verify it.
    - Details
    - An array of client certificates to be used. Each certificate object must have either both `certPath` and `keyPath`, a single `pfxPath`, or their corresponding direct value equivalents (`cert` and `key`, or `pfx`). Optionally, `passphrase` property should be provided if the certificate is encrypted. The `origin` property should be provided with an exact match to the request origin that the certificate is valid for.
    - Client certificate authentication is only active when at least one client certificate is provided. If you want to reject all client certificates sent by the server, you need to provide a client certificate with an `origin` that does not match any of the domains you plan to visit.
    - note
    - When using WebKit on macOS, accessing `localhost` will not pick up client certificates. You can make it work by replacing `localhost` with `local.playwright`.
    - `colorScheme` [null](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/null) | "light" | "dark" | "no-preference" (optional)#
    - Emulates [prefers-colors-scheme](https://developer.mozilla.org/en-US/docs/Web/CSS/@media/prefers-color-scheme) media feature, supported values are `'light'` and `'dark'`. See page.emulateMedia() for more details. Passing `null` resets emulation to system defaults. Defaults to `'light'`.
    - `contrast` [null](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/null) | "no-preference" | "more" (optional)#
    - Emulates `'prefers-contrast'` media feature, supported values are `'no-preference'`, `'more'`. See page.emulateMedia() for more details. Passing `null` resets emulation to system defaults. Defaults to `'no-preference'`.
    - `deviceScaleFactor` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Specify device scale factor (can be thought of as dpr). Defaults to `1`. Learn more about emulating devices with device scale factor.
    - `downloadsPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
    - `env` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [undefined]> (optional)#
    - `executablePath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Path to a browser executable to run instead of the bundled one. If executablePath is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
    - `extraHTTPHeaders` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - An object containing additional HTTP headers to be sent with every request. Defaults to none.
    - `firefoxUserPrefs` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) | [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type)> (optional) Added in: v1.40#
    - Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
    - You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
    - `forcedColors` [null](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/null) | "active" | "none" (optional)#
    - Emulates `'forced-colors'` media feature, supported values are `'active'`, `'none'`. See page.emulateMedia() for more details. Passing `null` resets emulation to system defaults. Defaults to `'none'`.
    - `geolocation` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `latitude` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - Latitude between -90 and 90.
        - `longitude` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - Longitude between -180 and 180.
        - `accuracy` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)
        - Non-negative accuracy value. Defaults to `0`.
    - `handleSIGHUP` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGHUP. Defaults to `true`.
    - `handleSIGINT` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on Ctrl-C. Defaults to `true`.
    - `handleSIGTERM` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGTERM. Defaults to `true`.
    - `hasTouch` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Specifies if viewport supports touch events. Defaults to false. Learn more about mobile emulation.
    - `headless` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
    - `httpCredentials` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `username` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - `password` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - `origin` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Restrain sending http credentials on specific origin (scheme://host:port).
        - `send` "unauthorized" | "always" (optional)
        - This option only applies to the requests sent from corresponding APIRequestContext and does not affect requests sent from the browser. `'always'` - `Authorization` header with basic authentication credentials will be sent with the each API request. `'unauthorized` - the credentials are only sent when 401 (Unauthorized) response with `WWW-Authenticate` header is received. Defaults to `'unauthorized'`.
    - Credentials for [HTTP authentication](https://developer.mozilla.org/en-US/docs/Web/HTTP/Authentication). If no origin is specified, the username and password are sent to any servers upon unauthorized responses.
    - `ignoreDefaultArgs` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) | [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from args. If an array is given, then filters out the given default arguments. Dangerous option; use with care. Defaults to `false`.
    - `ignoreHTTPSErrors` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to ignore HTTPS errors when sending network requests. Defaults to `false`.
    - `isMobile` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether the `meta viewport` tag is taken into account and touch events are enabled. isMobile is a part of device, so you don't actually need to set it manually. Defaults to `false` and is not supported in Firefox. Learn more about mobile emulation.
    - `javaScriptEnabled` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether or not to enable JavaScript in the context. Defaults to `true`. Learn more about disabling JavaScript.
    - `locale` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Specify user locale, for example `en-GB`, `de-DE`, etc. Locale will affect `navigator.language` value, `Accept-Language` request header value as well as number and date formatting rules. Defaults to the system default locale. Learn more about emulation in our emulation guide.
    - `logger` Logger (optional)#
    - Deprecated
    - The logs received by the logger are incomplete. Please use tracing instead.
    - Logger sink for Playwright logging.
    - `offline` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to emulate network being offline. Defaults to `false`. Learn more about network emulation.
    - `permissions` [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - A list of permissions to grant to all pages in this context. See browserContext.grantPermissions() for more details. Defaults to none.
    - `proxy` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `server` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
        - `bypass` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
        - `username` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional username to use if HTTP proxy requires authentication.
        - `password` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional password to use if HTTP proxy requires authentication.
    - Network proxy settings.
    - `recordHar` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `omitContent` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)
        - Optional setting to control whether to omit request content from the HAR. Defaults to `false`. Deprecated, use `content` policy instead.
        - `content` "omit" | "embed" | "attach" (optional)
        - Optional setting to control resource content management. If `omit` is specified, content is not persisted. If `attach` is specified, resources are persisted as separate files or entries in the ZIP archive. If `embed` is specified, content is stored inline the HAR file as per HAR specification. Defaults to `attach` for `.zip` output files and to `embed` for all other file extensions.
        - `path` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - Path on the filesystem to write the HAR file to. If the file name ends with `.zip`, `content: 'attach'` is used by default.
        - `mode` "full" | "minimal" (optional)
        - When set to `minimal`, only record information necessary for routing from HAR. This omits sizes, timing, page, cookies, security and other types of HAR information that are not used when replaying from HAR. Defaults to `full`.
        - `urlFilter` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [RegExp](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/RegExp) (optional)
        - A glob or regex pattern to filter requests that are stored in the HAR. When a baseURL via the context options was provided and the passed URL is a path, it gets merged via the [`new URL()`](https://developer.mozilla.org/en-US/docs/Web/API/URL/URL) constructor. Defaults to none.
    - Enables [HAR](http://www.softwareishard.com/blog/har-12-spec) recording for all pages into `recordHar.path` file. If not specified, the HAR is not recorded. Make sure to await browserContext.close() for the HAR to be saved.
    - `recordVideo` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `dir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Path to the directory to put videos into. If not specified, the videos will be stored in `artifactsDir` (see browserType.launch() options).
        - `size` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
            - `width` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
            - Video frame width.
            - `height` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
            - Video frame height.
        - Optional dimensions of the recorded videos. If not specified the size will be equal to `viewport` scaled down to fit into 800x800. If `viewport` is not configured explicitly the video size defaults to 800x450. Actual picture of each page will be scaled down if necessary to fit the specified size.
        - `showActions` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
            - `duration` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)
            - How long each annotation is displayed in milliseconds. Defaults to `500`.
            - `position` "top-left" | "top" | "top-right" | "bottom-left" | "bottom" | "bottom-right" (optional)
            - Position of the action title overlay. Defaults to `"top-right"`.
            - `fontSize` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)
            - Font size of the action title in pixels. Defaults to `24`.
            - `cursor` "none" | "pointer" (optional)
            - Cursor decoration shown for pointer actions. `"pointer"` (the default) renders a mouse pointer that animates from the previous action point to the next one. `"none"` disables the cursor decoration.
        - If specified, enables visual annotations on interacted elements during video recording.
    - Enables video recording for all pages into `recordVideo.dir` directory. If not specified videos are not recorded. Make sure to await browserContext.close() for videos to be saved.
    - `reducedMotion` [null](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/null) | "reduce" | "no-preference" (optional)#
    - Emulates `'prefers-reduced-motion'` media feature, supported values are `'reduce'`, `'no-preference'`. See page.emulateMedia() for more details. Passing `null` resets emulation to system defaults. Defaults to `'no-preference'`.
    - `screen` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `width` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - page width in pixels.
        - `height` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - page height in pixels.
    - Emulates consistent window screen size available inside web page via `window.screen`. Is only used when the viewport is set.
    - `serviceWorkers` "allow" | "block" (optional)#
    - Whether to allow sites to register Service workers. Defaults to `'allow'`.
        - `'allow'`: [Service Workers](https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API) can be registered.
        - `'block'`: Playwright will block all registration of Service Workers.
    - `slowMo` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
    - `strictSelectors` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - If set to true, enables strict selectors mode for this context. In the strict selectors mode all operations on selectors that imply single target DOM element will throw when more than one element matches the selector. This option does not affect any Locator APIs (Locators are always strict). Defaults to `false`. See Locator to learn more about the strict mode.
    - `timeout` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
    - `timezoneId` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Changes the timezone of the context. See [ICU's metaZones.txt](https://cs.chromium.org/chromium/src/third_party/icu/source/data/misc/metaZones.txt?rcl=faee8bc70570192d82d2978a71e2a615788597d1) for a list of supported timezone IDs. Defaults to the system timezone.
    - `tracesDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, traces are saved into this directory.
    - `userAgent` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Specific user agent to use in this context.
    - `viewport` [null](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/null) | [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `width` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - page width in pixels.
        - `height` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type)
        - page height in pixels.
    - Emulates consistent viewport for each page. Defaults to an 1280x720 viewport. Use `null` to disable the consistent viewport emulation. Learn more about viewport emulation.
    - note
    - The `null` value opts out from the default presets, makes viewport depend on the host window size defined by the operating system. It makes the execution of the tests non-deterministic.
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<BrowserContext>#

Added before v1.9
browserType.launchServer

Returns the browser app instance. You can connect to it via browserType.connect(), which requires the major/minor client/server version to match (1.2.3 → is compatible with 1.2.x).

Usage

Launches browser server that client can connect to. An example of launching a browser executable and connecting to it later:

```
const { chromium } = require('playwright');  // Or 'webkit' or 'firefox'.

(async () => {
  const browserServer = await chromium.launchServer();
  const wsEndpoint = browserServer.wsEndpoint();
  // Use web socket endpoint later to establish a connection.
  const browser = await chromium.connect(wsEndpoint);
  // Close browser instance.
  await browserServer.close();
})();
```

Arguments

- `options` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - `args` [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - warning
    - Use custom browser args at your own risk, as some of them may break Playwright functionality.
    - Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
    - `artifactsDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
    - `channel` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Browser distribution channel.
    - Use "chromium" to opt in to new headless mode.
    - Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
    - `chromiumSandbox` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Enable Chromium sandboxing. Defaults to `false`.
    - `downloadsPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
    - `env` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [undefined]> (optional)#
    - `executablePath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - Path to a browser executable to run instead of the bundled one. If executablePath is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
    - `firefoxUserPrefs` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type), [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) | [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) | [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type)> (optional)#
    - Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
    - You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
    - `handleSIGHUP` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGHUP. Defaults to `true`.
    - `handleSIGINT` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on Ctrl-C. Defaults to `true`.
    - `handleSIGTERM` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Close the browser process on SIGTERM. Defaults to `true`.
    - `headless` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) (optional)#
    - Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
    - `host` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional) Added in: v1.45#
    - Host to use for the web socket. It is optional and defaults to `localhost`, accepting connections only from the loopback interface. Pass an explicit address (e.g. `0.0.0.0`) to accept connections from the network — be aware this exposes the browser RPC to anything that can reach the listening port.
    - `ignoreDefaultArgs` [boolean](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Boolean_type) | [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)<[string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)> (optional)#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from args. If an array is given, then filters out the given default arguments. Dangerous option; use with care. Defaults to `false`.
    - `logger` Logger (optional)#
    - Deprecated
    - The logs received by the logger are incomplete. Please use tracing instead.
    - Logger sink for Playwright logging.
    - `port` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Port to use for the web socket. Defaults to 0 that picks any available port.
    - `proxy` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
        - `server` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
        - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
        - `bypass` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
        - `username` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional username to use if HTTP proxy requires authentication.
        - `password` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)
        - Optional password to use if HTTP proxy requires authentication.
    - Network proxy settings.
    - `timeout` [number](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#Number_type) (optional)#
    - Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
    - `tracesDir` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional)#
    - If specified, traces are saved into this directory.
    - `wsPath` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type) (optional) Added in: v1.15#
    - Path at which to serve the Browser Server. For security, this defaults to an unguessable string.
    - warning
    - Any process or web page (including those running in Playwright) with knowledge of the `wsPath` can take control of the OS user. For this reason, you should use an unguessable token when using this option.
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<BrowserServer>#

Added before v1.9
browserType.name

Returns browser name. For example: `'chromium'`, `'webkit'` or `'firefox'`.

Usage

```
browserType.name();
```

Returns

- [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)#

### BrowserType and ConnectOverCDPAsync in Playwright for .NET

Published by Playwright. Source page: [BrowserType and ConnectOverCDPAsync in Playwright for .NET](https://playwright.dev/dotnet/docs/api/class-browsertype).

BrowserType provides methods to launch a specific browser instance or connect to an existing one. The following is a typical example of using Playwright to drive automation:

```
using Microsoft.Playwright;
using System.Threading.Tasks;

class BrowserTypeExamples
{
    public static async Task Run()
    {
        using var playwright = await Playwright.CreateAsync();
        var chromium = playwright.Chromium;
        var browser = await chromium.LaunchAsync();
        var page = await browser.NewPageAsync();
        await page.GotoAsync("https://www.bing.com");
        // other actions
        await browser.CloseAsync();
    }
}
```

Added before v1.9
browserType.ConnectAsync

This method attaches Playwright to an existing browser instance created via `BrowserType.launchServer` in Node.js.

note

The major and minor version of the Playwright instance that connects needs to match the version of Playwright that launches the browser (1.2.3 → is compatible with 1.2.x).

Usage

```
await BrowserType.ConnectAsync(endpoint, options);
```

Arguments

- `endpoint` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string) Added in: v1.10#
- A Playwright browser websocket endpoint to connect to. You obtain this endpoint via `BrowserServer.wsEndpoint`.
- `options` `BrowserTypeConnectOptions?` (optional)
    - `ExposeNetwork` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional) Added in: v1.37#
    - This option exposes network available on the connecting client to the browser being connected to. Consists of a list of rules separated by comma.
    - Available rules:
        - Hostname pattern, for example: `example.com`, `*.org:99`, `x.*.y.com`, `*foo.org`.
        - IP literal, for example: `127.0.0.1`, `0.0.0.0:99`, `[::1]`, `[0:0::1]:99`.
        - `<loopback>` that matches local loopback interfaces: `localhost`, `*.localhost`, `127.0.0.1`, `[::1]`.
    - Some common examples:
        - `"*"` to expose all network.
        - `"<loopback>"` to expose localhost network.
        - `"*.test.internal-domain,*.staging.internal-domain,<loopback>"` to expose test/staging deployments and localhost.
    - `Headers` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional) Added in: v1.11#
    - Additional HTTP headers to be sent with web socket connect request. Optional.
    - `SlowMo` [float]? (optional) Added in: v1.10#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
    - `Timeout` [float]? (optional) Added in: v1.10#
    - Maximum time in milliseconds to wait for the connection to be established. Defaults to `0` (no timeout).
Returns

- Browser#

Added in: v1.9
browserType.ConnectOverCDPAsync

This method attaches Playwright to an existing browser instance using the Chrome DevTools Protocol.

The default browser context is accessible via Browser.Contexts.

note

Connecting over the Chrome DevTools Protocol is only supported for Chromium-based browsers.

note

This connection is significantly lower fidelity than the Playwright protocol connection via BrowserType.ConnectAsync(). If you are experiencing issues or attempting to use advanced functionality, you probably want to use BrowserType.ConnectAsync().

warning

Playwright maintains a curated list of arguments for launching the browser. If you launch the browser without Playwright and do not pass the exact same arguments, some of Playwright functionality may be broken upon connecting to the browser.

Usage

```
var browser = await playwright.Chromium.ConnectOverCDPAsync("http://localhost:9222");
var defaultContext = browser.Contexts[0];
var page = defaultContext.Pages[0];
```

Arguments

- `endpointURL` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string) Added in: v1.11#
- A CDP websocket endpoint or http url to connect to. For example `[localhost:9222](http://localhost:9222/)` or `ws://127.0.0.1:9222/devtools/browser/387adf4c-243f-4051-a181-46798f4a46f4`.
- `options` `BrowserTypeConnectOverCDPOptions?` (optional)
    - `ArtifactsDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional) Added in: v1.61#
    - If specified, browser artifacts (such as traces and downloads) are saved into this directory.
    - `Headers` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional) Added in: v1.11#
    - Additional HTTP headers to be sent with connect request. Optional.
    - `IsLocal` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional) Added in: v1.58#
    - Tells Playwright that it runs on the same host as the CDP server. It will enable certain optimizations that rely upon the file system being the same between Playwright and the Browser.
    - `NoDefaults` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional) Added in: v1.60#
    - When true, Playwright will not apply its default overrides to the existing default browser context. Specifically, AcceptDownloads is left at the browser's setting, focus emulation is not enabled, and media emulation options (such as ColorScheme, ReducedMotion, ForcedColors, and Contrast) are not applied. Useful when attaching to a user's daily-driver browser where these overrides would interfere with existing browser state. New contexts created via Browser.NewContextAsync() are not affected. Defaults to `false`.
    - `SlowMo` [float]? (optional) Added in: v1.11#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on. Defaults to 0.
    - `Timeout` [float]? (optional) Added in: v1.11#
    - Maximum time in milliseconds to wait for the connection to be established. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
Returns

- Browser#

Added before v1.9
browserType.ExecutablePath

A path where Playwright expects to find a bundled browser executable.

Usage

```
BrowserType.ExecutablePath
```

Returns

- [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)#

Added before v1.9
browserType.LaunchAsync

Returns the browser instance.

Usage

You can use IgnoreDefaultArgs to filter out `--mute-audio` from default arguments:

```
var browser = await playwright.Chromium.LaunchAsync(new() {
    IgnoreDefaultArgs = new[] { "--mute-audio" }
});
```

Chromium-only Playwright can also be used to control the Google Chrome or Microsoft Edge browsers, but it works best with the version of Chromium it is bundled with. There is no guarantee it will work with any other version. Use ExecutablePath option with extreme caution.

If Google Chrome (rather than Chromium) is preferred, a [Chrome Canary](https://www.google.com/chrome/browser/canary.html) or [Dev Channel](https://www.chromium.org/getting-involved/dev-channel) build is suggested.

Stock browsers like Google Chrome and Microsoft Edge are suitable for tests that require proprietary media codecs for video playback. See [this article](https://www.howtogeek.com/202825/what%E2%80%99s-the-difference-between-chromium-and-chrome/) for other differences between Chromium and Chrome. [This article](https://chromium.googlesource.com/chromium/src/+/lkgr/docs/chromium_browser_vs_google_chrome.md) describes some differences for Linux users.

Arguments

- `options` `BrowserTypeLaunchOptions?` (optional)
    - `Args` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - warning
    - Use custom browser args at your own risk, as some of them may break Playwright functionality.
    - Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
    - `ArtifactsDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
    - `Channel` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Browser distribution channel.
    - Use "chromium" to opt in to new headless mode.
    - Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
    - `ChromiumSandbox` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Enable Chromium sandboxing. Defaults to `false`.
    - `DownloadsPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
    - `Env` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - Specify environment variables that will be visible to the browser. Defaults to `process.env`.
    - `ExecutablePath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Path to a browser executable to run instead of the bundled one. If ExecutablePath is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
    - `FirefoxUserPrefs` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [object]> (optional)#
    - Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
    - You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
    - `HandleSIGHUP` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on SIGHUP. Defaults to `true`.
    - `HandleSIGINT` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on Ctrl-C. Defaults to `true`.
    - `HandleSIGTERM` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on SIGTERM. Defaults to `true`.
    - `Headless` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
    - `IgnoreAllDefaultArgs` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional) Added in: v1.9#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from Args. Dangerous option; use with care. Defaults to `false`.
    - `IgnoreDefaultArgs` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from Args. Dangerous option; use with care.
    - `Proxy` Proxy? (optional)#
        - `Server` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)
        - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
        - `Bypass` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
        - `Username` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional username to use if HTTP proxy requires authentication.
        - `Password` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional password to use if HTTP proxy requires authentication.
    - Network proxy settings.
    - `SlowMo` [float]? (optional)#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
    - `Timeout` [float]? (optional)#
    - Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
    - `TracesDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, traces are saved into this directory.
Returns

- Browser#

Added before v1.9
browserType.LaunchPersistentContextAsync

Returns the persistent browser context instance.

Launches browser that uses persistent storage located at userDataDir and returns the only context. Closing this context will automatically close the browser.

Usage

```
await BrowserType.LaunchPersistentContextAsync(userDataDir, options);
```

Arguments

- `userDataDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)#
- Path to a User Data Directory, which stores browser session data like cookies and local storage. Pass an empty string to create a temporary directory.
- More details for [Chromium](https://chromium.googlesource.com/chromium/src/+/master/docs/user_data_dir.md#introduction) and [Firefox](https://wiki.mozilla.org/Firefox/CommandLineOptions#User_profile). Chromium's user data directory is the parent directory of the "Profile Path" seen at `chrome://version`.
- Note that browsers do not allow launching multiple instances with the same User Data Directory.
- warning
- Chromium/Chrome: Due to recent Chrome policy changes, automating the default Chrome user profile is not supported. Pointing `userDataDir` to Chrome's main "User Data" directory (the profile used for your regular browsing) may result in pages not loading or the browser exiting. Create and use a separate directory (for example, an empty folder) as your automation profile instead. See [developer.chrome.com](https://developer.chrome.com/blog/remote-debugging-port) for details.
- `options` `BrowserTypeLaunchPersistentContextOptions?` (optional)
    - `AcceptDownloads` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether to automatically download all the attachments. Defaults to `true` where all the downloads are accepted.
    - `Args` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - warning
    - Use custom browser args at your own risk, as some of them may break Playwright functionality.
    - Additional arguments to pass to the browser instance. The list of Chromium flags can be found [here](https://peter.sh/experiments/chromium-command-line-switches/).
    - `ArtifactsDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, artifacts (traces, videos, downloads, HAR files, etc.) are saved into this directory. The directory is not cleaned up when the browser closes. If not specified, a temporary directory is used and cleaned up when the browser closes.
    - `BaseURL` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - When using Page.GotoAsync(), Page.RouteAsync(), Page.WaitForURLAsync(), Page.RunAndWaitForRequestAsync(), or Page.RunAndWaitForResponseAsync() it takes the base URL in consideration by using the [`URL()`](https://developer.mozilla.org/en-US/docs/Web/API/URL/URL) constructor for building the corresponding URL. Unset by default. Examples:
        - baseURL: `[localhost:3000](http://localhost:3000)` and navigating to `/bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
        - baseURL: `[localhost:3000](http://localhost:3000/foo/)` and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/foo/bar.html)`
        - baseURL: `[localhost:3000](http://localhost:3000/foo)` (without trailing slash) and navigating to `./bar.html` results in `[localhost:3000](http://localhost:3000/bar.html)`
    - `BypassCSP` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Toggles bypassing page's Content-Security-Policy. Defaults to `false`.
    - `Channel` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Browser distribution channel.
    - Use "chromium" to opt in to new headless mode.
    - Use "chrome", "chrome-beta", "chrome-dev", "chrome-canary", "msedge", "msedge-beta", "msedge-dev", or "msedge-canary" to use branded Google Chrome and Microsoft Edge.
    - `ChromiumSandbox` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Enable Chromium sandboxing. Defaults to `false`.
    - `ClientCertificates` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<ClientCertificates> (optional) Added in: 1.46#
        - `Origin` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)
        - Exact origin that the certificate is valid for. Origin includes `https` protocol, a hostname and optionally a port.
        - `CertPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Path to the file with the certificate in PEM format.
        - `Cert` [byte](https://docs.microsoft.com/en-us/dotnet/api/system.byte)[]? (optional)
        - Direct value of the certificate in PEM format.
        - `KeyPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Path to the file with the private key in PEM format.
        - `Key` [byte](https://docs.microsoft.com/en-us/dotnet/api/system.byte)[]? (optional)
        - Direct value of the private key in PEM format.
        - `PfxPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Path to the PFX or PKCS12 encoded private key and certificate chain.
        - `Pfx` [byte](https://docs.microsoft.com/en-us/dotnet/api/system.byte)[]? (optional)
        - Direct value of the PFX or PKCS12 encoded private key and certificate chain.
        - `Passphrase` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Passphrase for the private key (PEM or PFX).
    - TLS Client Authentication allows the server to request a client certificate and verify it.
    - Details
    - An array of client certificates to be used. Each certificate object must have either both `certPath` and `keyPath`, a single `pfxPath`, or their corresponding direct value equivalents (`cert` and `key`, or `pfx`). Optionally, `passphrase` property should be provided if the certificate is encrypted. The `origin` property should be provided with an exact match to the request origin that the certificate is valid for.
    - Client certificate authentication is only active when at least one client certificate is provided. If you want to reject all client certificates sent by the server, you need to provide a client certificate with an `origin` that does not match any of the domains you plan to visit.
    - note
    - When using WebKit on macOS, accessing `localhost` will not pick up client certificates. You can make it work by replacing `localhost` with `local.playwright`.
    - `ColorScheme` `enum ColorScheme { Light, Dark, NoPreference, Null }?` (optional)#
    - Emulates [prefers-colors-scheme](https://developer.mozilla.org/en-US/docs/Web/CSS/@media/prefers-color-scheme) media feature, supported values are `'light'` and `'dark'`. See Page.EmulateMediaAsync() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'light'`.
    - `Contrast` `enum Contrast { NoPreference, More, Null }?` (optional)#
    - Emulates `'prefers-contrast'` media feature, supported values are `'no-preference'`, `'more'`. See Page.EmulateMediaAsync() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'no-preference'`.
    - `DeviceScaleFactor` [float]? (optional)#
    - Specify device scale factor (can be thought of as dpr). Defaults to `1`. Learn more about emulating devices with device scale factor.
    - `DownloadsPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, accepted downloads are downloaded into this directory. Otherwise, temporary directory is created and is deleted when browser is closed. In either case, the downloads are deleted when the browser context they were created in is closed.
    - `Env` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - Specify environment variables that will be visible to the browser. Defaults to `process.env`.
    - `ExecutablePath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Path to a browser executable to run instead of the bundled one. If ExecutablePath is a relative path, then it is resolved relative to the current working directory. Note that Playwright only works with the bundled Chromium, Firefox or WebKit, use at your own risk.
    - `ExtraHTTPHeaders` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - An object containing additional HTTP headers to be sent with every request. Defaults to none.
    - `FirefoxUserPrefs` [IDictionary](https://docs.microsoft.com/en-us/dotnet/api/system.collections.idictionary)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), [object]> (optional) Added in: v1.40#
    - Firefox user preferences. Learn more about the Firefox user preferences at [`about:config`](https://support.mozilla.org/en-US/kb/about-config-editor-firefox).
    - You can also provide a path to a custom [`policies.json` file](https://mozilla.github.io/policy-templates/) via `PLAYWRIGHT_FIREFOX_POLICIES_JSON` environment variable.
    - `ForcedColors` `enum ForcedColors { Active, None, Null }?` (optional)#
    - Emulates `'forced-colors'` media feature, supported values are `'active'`, `'none'`. See Page.EmulateMediaAsync() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'none'`.
    - `Geolocation` Geolocation? (optional)#
        - `Latitude` [float]
        - Latitude between -90 and 90.
        - `Longitude` [float]
        - Longitude between -180 and 180.
        - `Accuracy` [float]? (optional)
        - Non-negative accuracy value. Defaults to `0`.
    - `HandleSIGHUP` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on SIGHUP. Defaults to `true`.
    - `HandleSIGINT` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on Ctrl-C. Defaults to `true`.
    - `HandleSIGTERM` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Close the browser process on SIGTERM. Defaults to `true`.
    - `HasTouch` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Specifies if viewport supports touch events. Defaults to false. Learn more about mobile emulation.
    - `Headless` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether to run browser in headless mode. More details for [Chromium](https://developers.google.com/web/updates/2017/04/headless-chrome) and [Firefox](https://hacks.mozilla.org/2017/12/using-headless-mode-in-firefox/). Defaults to `true`.
    - `HttpCredentials` HttpCredentials? (optional)#
        - `Username` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)
        - `Password` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)
        - `Origin` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Restrain sending http credentials on specific origin (scheme://host:port).
        - `Send` `enum HttpCredentialsSend { Unauthorized, Always }?` (optional)
        - This option only applies to the requests sent from corresponding APIRequestContext and does not affect requests sent from the browser. `'always'` - `Authorization` header with basic authentication credentials will be sent with the each API request. `'unauthorized` - the credentials are only sent when 401 (Unauthorized) response with `WWW-Authenticate` header is received. Defaults to `'unauthorized'`.
    - Credentials for [HTTP authentication](https://developer.mozilla.org/en-US/docs/Web/HTTP/Authentication). If no origin is specified, the username and password are sent to any servers upon unauthorized responses.
    - `IgnoreAllDefaultArgs` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional) Added in: v1.9#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from Args. Dangerous option; use with care. Defaults to `false`.
    - `IgnoreDefaultArgs` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - If `true`, Playwright does not pass its own configurations args and only uses the ones from Args. Dangerous option; use with care.
    - `IgnoreHTTPSErrors` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether to ignore HTTPS errors when sending network requests. Defaults to `false`.
    - `IsMobile` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether the `meta viewport` tag is taken into account and touch events are enabled. isMobile is a part of device, so you don't actually need to set it manually. Defaults to `false` and is not supported in Firefox. Learn more about mobile emulation.
    - `JavaScriptEnabled` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether or not to enable JavaScript in the context. Defaults to `true`. Learn more about disabling JavaScript.
    - `Locale` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Specify user locale, for example `en-GB`, `de-DE`, etc. Locale will affect `navigator.language` value, `Accept-Language` request header value as well as number and date formatting rules. Defaults to the system default locale. Learn more about emulation in our emulation guide.
    - `Offline` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Whether to emulate network being offline. Defaults to `false`. Learn more about network emulation.
    - `Permissions` [IEnumerable](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable)?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string)> (optional)#
    - A list of permissions to grant to all pages in this context. See BrowserContext.GrantPermissionsAsync() for more details. Defaults to none.
    - `Proxy` Proxy? (optional)#
        - `Server` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)
        - Proxy to be used for all requests. HTTP and SOCKS proxies are supported, for example `[myproxy.com:3128](http://myproxy.com:3128)` or `socks5://myproxy.com:3128`. Short form `myproxy.com:3128` is considered an HTTP proxy.
        - `Bypass` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional comma-separated domains to bypass proxy, for example `".com, chromium.org, .domain.com"`.
        - `Username` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional username to use if HTTP proxy requires authentication.
        - `Password` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)
        - Optional password to use if HTTP proxy requires authentication.
    - Network proxy settings.
    - `RecordHarContent` `enum HarContentPolicy { Omit, Embed, Attach }?` (optional)#
    - Optional setting to control resource content management. If `omit` is specified, content is not persisted. If `attach` is specified, resources are persisted as separate files and all of these files are archived along with the HAR file. Defaults to `embed`, which stores content inline the HAR file as per HAR specification.
    - `RecordHarMode` `enum HarMode { Full, Minimal }?` (optional)#
    - When set to `minimal`, only record information necessary for routing from HAR. This omits sizes, timing, page, cookies, security and other types of HAR information that are not used when replaying from HAR. Defaults to `full`.
    - `RecordHarOmitContent` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - Optional setting to control whether to omit request content from the HAR. Defaults to `false`.
    - `RecordHarPath` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Enables [HAR](http://www.softwareishard.com/blog/har-12-spec) recording for all pages into the specified HAR file on the filesystem. If not specified, the HAR is not recorded. Make sure to call BrowserContext.CloseAsync() for the HAR to be saved.
    - `RecordHarUrlFilter|RecordHarUrlFilterRegex` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? | [Regex](https://docs.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex)? (optional)#
    - `RecordVideoDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Enables video recording for all pages into the specified directory. If not specified videos are not recorded. Make sure to call BrowserContext.CloseAsync() for videos to be saved.
    - `RecordVideoSize` RecordVideoSize? (optional)#
        - `Width` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - Video frame width.
        - `Height` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - Video frame height.
    - Dimensions of the recorded videos. If not specified the size will be equal to `viewport` scaled down to fit into 800x800. If `viewport` is not configured explicitly the video size defaults to 800x450. Actual picture of each page will be scaled down if necessary to fit the specified size.
    - `ReducedMotion` `enum ReducedMotion { Reduce, NoPreference, Null }?` (optional)#
    - Emulates `'prefers-reduced-motion'` media feature, supported values are `'reduce'`, `'no-preference'`. See Page.EmulateMediaAsync() for more details. Passing `'null'` resets emulation to system defaults. Defaults to `'no-preference'`.
    - `ScreenSize` ScreenSize? (optional)#
        - `Width` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - page width in pixels.
        - `Height` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - page height in pixels.
    - Emulates consistent window screen size available inside web page via `window.screen`. Is only used when the ViewportSize is set.
    - `ServiceWorkers` `enum ServiceWorkerPolicy { Allow, Block }?` (optional)#
    - Whether to allow sites to register Service workers. Defaults to `'allow'`.
        - `'allow'`: [Service Workers](https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API) can be registered.
        - `'block'`: Playwright will block all registration of Service Workers.
    - `SlowMo` [float]? (optional)#
    - Slows down Playwright operations by the specified amount of milliseconds. Useful so that you can see what is going on.
    - `StrictSelectors` [bool](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)? (optional)#
    - If set to true, enables strict selectors mode for this context. In the strict selectors mode all operations on selectors that imply single target DOM element will throw when more than one element matches the selector. This option does not affect any Locator APIs (Locators are always strict). Defaults to `false`. See Locator to learn more about the strict mode.
    - `Timeout` [float]? (optional)#
    - Maximum time in milliseconds to wait for the browser instance to start. Defaults to `30000` (30 seconds). Pass `0` to disable timeout.
    - `TimezoneId` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Changes the timezone of the context. See [ICU's metaZones.txt](https://cs.chromium.org/chromium/src/third_party/icu/source/data/misc/metaZones.txt?rcl=faee8bc70570192d82d2978a71e2a615788597d1) for a list of supported timezone IDs. Defaults to the system timezone.
    - `TracesDir` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - If specified, traces are saved into this directory.
    - `UserAgent` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)? (optional)#
    - Specific user agent to use in this context.
    - `ViewportSize` ViewportSize? (optional)#
        - `Width` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - page width in pixels.
        - `Height` [int](https://docs.microsoft.com/en-us/dotnet/api/system.int32)
        - page height in pixels.
    - Emulates consistent viewport for each page. Defaults to an 1280x720 viewport. Use `ViewportSize.NoViewport` to disable the consistent viewport emulation. Learn more about viewport emulation.
    - note
    - The `ViewportSize.NoViewport` value opts out from the default presets, makes viewport depend on the host window size defined by the operating system. It makes the execution of the tests non-deterministic.
Returns

- BrowserContext#

Added before v1.9
browserType.Name

Returns browser name. For example: `'chromium'`, `'webkit'` or `'firefox'`.

Usage

```
BrowserType.Name
```

Returns

- [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)#

### CDPSession in Playwright for .NET

Published by Playwright. Source page: [CDPSession in Playwright for .NET](https://playwright.dev/dotnet/docs/api/class-cdpsession).

The `CDPSession` instances are used to talk raw Chrome Devtools Protocol:

- protocol methods can be called with `session.send` method.
- protocol events can be subscribed to with `session.on` method.
Useful links:

- Documentation on DevTools Protocol can be found here: [DevTools Protocol Viewer](https://chromedevtools.github.io/devtools-protocol/).
- Getting Started with DevTools Protocol: [github.com](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/README.md)

```
var client = await Page.Context.NewCDPSessionAsync(Page);
await client.SendAsync("Runtime.enable");
client.Event("Animation.animationCreated").OnEvent += (_, _) => Console.WriteLine("Animation created!");
var response = await client.SendAsync("Animation.getPlaybackRate");
var playbackRate = response.Value.GetProperty("playbackRate").GetDouble();
Console.WriteLine("playback rate is " + playbackRate);
await client.SendAsync("Animation.setPlaybackRate", new() { { "playbackRate", playbackRate / 2 } });
```

Added before v1.9
cdpSession.DetachAsync

Detaches the CDPSession from the target. Once detached, the CDPSession object won't emit any events and can't be used to send messages.

Usage

```
await CdpSession.DetachAsync();
```

Returns

- [void](https://docs.microsoft.com/en-us/dotnet/csharp/language-reference/builtin-types/void)#

Added in: v.1.30
cdpSession.Event

Returns an event emitter for the given CDP event name.

Usage

```
CdpSession.Event(eventName);
```

Arguments

- `eventName` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string) Added in: v1.30#
- CDP event name.
Returns

- CDPSessionEvent#

Added before v1.9
cdpSession.SendAsync

Usage

```
await CdpSession.SendAsync(method, params);
```

Arguments

- `method` [string](https://docs.microsoft.com/en-us/dotnet/api/system.string)#
- Protocol method name.
- `args` [Map]?<[string](https://docs.microsoft.com/en-us/dotnet/api/system.string), Args> (optional) Added in: v1.30#
- Optional method parameters.
Returns

- [JsonElement?]#

Added in: v1.59
cdpSession.event Close

Emitted when the session is closed, either because the target was closed or `session.detach()` was called.

Usage

```
CdpSession.Close += async (_, cDPSession) => {};
```

Event data

- CDPSession

### CDPSession in Playwright for JavaScript

Published by Playwright. Source page: [CDPSession in Playwright for JavaScript](https://playwright.dev/docs/api/class-cdpsession).

The `CDPSession` instances are used to talk raw Chrome Devtools Protocol:

- protocol methods can be called with `session.send` method.
- protocol events can be subscribed to with `session.on` method.
Useful links:

- Documentation on DevTools Protocol can be found here: [DevTools Protocol Viewer](https://chromedevtools.github.io/devtools-protocol/).
- Getting Started with DevTools Protocol: [github.com](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/README.md)

```
const client = await page.context().newCDPSession(page);
await client.send('Animation.enable');
client.on('Animation.animationCreated', () => console.log('Animation created!'));
const response = await client.send('Animation.getPlaybackRate');
console.log('playback rate is ' + response.playbackRate);
await client.send('Animation.setPlaybackRate', {
  playbackRate: response.playbackRate / 2
});
```

Added before v1.9
cdpSession.detach

Detaches the CDPSession from the target. Once detached, the CDPSession object won't emit any events and can't be used to send messages.

Usage

```
await cdpSession.detach();
```

Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<[void](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/undefined)>#

Added before v1.9
cdpSession.send

Usage

```
await cdpSession.send(method);
await cdpSession.send(method, params);
```

Arguments

- `method` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)#
- Protocol method name.
- `params` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)#
- Optional method parameters.
Returns

- [Promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)<[Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)>#

Added in: v1.59
cdpSession.on('close')

Emitted when the session is closed, either because the target was closed or `session.detach()` was called.

Usage

```
cdpSession.on('close', data => {});
```

Event data

- CDPSession

Added in: v1.59
cdpSession.on('event')

Emitted for every CDP event received from the session. Allows subscribing to all CDP events at once without knowing their names ahead of time.

Usage

```
session.on('event', ({ method, params }) => {
  console.log(`CDP event: ${method}`, params);
});
```

Event data

- [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object)
    - `method` [string](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Data_structures#String_type)
    - CDP event name.
    - `params` [Object](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object) (optional)
    - CDP event parameters.

### CDPSession in Playwright for Python

Published by Playwright. Source page: [CDPSession in Playwright for Python](https://playwright.dev/python/docs/api/class-cdpsession).

The `CDPSession` instances are used to talk raw Chrome Devtools Protocol:

- protocol methods can be called with `session.send` method.
- protocol events can be subscribed to with `session.on` method.
Useful links:

- Documentation on DevTools Protocol can be found here: [DevTools Protocol Viewer](https://chromedevtools.github.io/devtools-protocol/).
- Getting Started with DevTools Protocol: [github.com](https://github.com/aslushnikov/getting-started-with-cdp/blob/master/README.md)
- Sync
- Async

```
client = page.context.new_cdp_session(page)
client.send("Animation.enable")
client.on("Animation.animationCreated", lambda: print("animation created!"))
response = client.send("Animation.getPlaybackRate")
print("playback rate is " + str(response["playbackRate"]))
client.send("Animation.setPlaybackRate", {
    "playbackRate": response["playbackRate"] / 2
})
```

```
client = await page.context.new_cdp_session(page)
await client.send("Animation.enable")
client.on("Animation.animationCreated", lambda: print("animation created!"))
response = await client.send("Animation.getPlaybackRate")
print("playback rate is " + str(response["playbackRate"]))
await client.send("Animation.setPlaybackRate", {
    "playbackRate": response["playbackRate"] / 2
})
```

Added before v1.9
cdpSession.detach

Detaches the CDPSession from the target. Once detached, the CDPSession object won't emit any events and can't be used to send messages.

Usage

```
cdp_session.detach()
```

Returns

- [NoneType](https://docs.python.org/3/library/constants.html#None)#

Added before v1.9
cdpSession.send

Usage

```
cdp_session.send(method)
cdp_session.send(method, **kwargs)
```

Arguments

- `method` [str](https://docs.python.org/3/library/stdtypes.html#text-sequence-type-str)#
- Protocol method name.
- `params` [Dict](https://docs.python.org/3/library/typing.html#typing.Dict) (optional)#
- Optional method parameters.
Returns

- [Dict](https://docs.python.org/3/library/typing.html#typing.Dict)#

Added in: v1.59
cdpSession.on("close")

Emitted when the session is closed, either because the target was closed or `session.detach()` was called.

Usage

```
cdp_session.on("close", handler)
```

Event data

- CDPSession

## Protocol reference

### Protocol domains, latest (tip-of-tree)

Published by Chrome DevTools team. Source page: [Protocol domains, latest (tip-of-tree)](https://chromedevtools.github.io/devtools-protocol/tot/).

The reference itself is drawn in the browser from a JSON file, so the saved page holds only this list of domains and the stability marking each one carries. To read the commands and events themselves, ask a browser already running with the remote debugging port open: the whole protocol it speaks is served at localhost:9222/json/protocol. The same definitions are published as browser_protocol.json and js_protocol.json in the ChromeDevTools devtools-protocol repository.

- Accessibility — experimental
- Animation — experimental
- Audits — experimental
- Autofill — experimental
- BackgroundService — experimental
- BluetoothEmulation — experimental
- Browser — in stable 1.3
- CacheStorage — experimental
- Cast — experimental
- Console — deprecated, v8 inspector
- CrashReportContext — experimental
- CSS — experimental
- Debugger — in stable 1.3, v8 inspector
- DeviceAccess — experimental
- DeviceOrientation — experimental
- DOM — in stable 1.3
- DOMDebugger — in stable 1.3
- DOMSnapshot — experimental
- DOMStorage — experimental
- Emulation — in stable 1.3
- EventBreakpoints — experimental
- Extensions — experimental
- FedCm — experimental
- Fetch — in stable 1.3
- FileSystem — experimental
- HeadlessExperimental — experimental
- HeapProfiler — experimental, v8 inspector
- IndexedDB — experimental
- Input — in stable 1.3
- Inspector — experimental
- IO — in stable 1.3
- LayerTree — experimental
- Log — in stable 1.3
- Media — experimental
- Memory — experimental
- Network — in stable 1.3
- Overlay — experimental
- Page — in stable 1.3
- Performance — in stable 1.3
- PerformanceTimeline — experimental
- Preload — experimental
- Profiler — in stable 1.3, v8 inspector
- PWA — experimental
- Runtime — in stable 1.3, v8 inspector
- Schema — deprecated, v8 inspector
- Security — in stable 1.3
- ServiceWorker — experimental
- SmartCardEmulation — experimental
- Storage — experimental
- SystemInfo — experimental
- Target — in stable 1.3
- Tethering — experimental
- Tracing — in stable 1.3
- WebAudio — experimental
- WebAuthn — experimental
- WebMCP — experimental

### Protocol domains, stable 1.3

Published by Chrome DevTools team. Source page: [Protocol domains, stable 1.3](https://chromedevtools.github.io/devtools-protocol/1-3/).

The reference itself is drawn in the browser from a JSON file, so the saved page holds only this list of domains and the stability marking each one carries. To read the commands and events themselves, ask a browser already running with the remote debugging port open: the whole protocol it speaks is served at localhost:9222/json/protocol. The same definitions are published as browser_protocol.json and js_protocol.json in the ChromeDevTools devtools-protocol repository.

- Accessibility — experimental
- Animation — experimental
- Audits — experimental
- Autofill — experimental
- BackgroundService — experimental
- BluetoothEmulation — experimental
- Browser — in stable 1.3
- CacheStorage — experimental
- Cast — experimental
- Console — deprecated, v8 inspector
- CrashReportContext — experimental
- CSS — experimental
- Debugger — in stable 1.3, v8 inspector
- DeviceAccess — experimental
- DeviceOrientation — experimental
- DOM — in stable 1.3
- DOMDebugger — in stable 1.3
- DOMSnapshot — experimental
- DOMStorage — experimental
- Emulation — in stable 1.3
- EventBreakpoints — experimental
- Extensions — experimental
- FedCm — experimental
- Fetch — in stable 1.3
- FileSystem — experimental
- HeadlessExperimental — experimental
- HeapProfiler — experimental, v8 inspector
- IndexedDB — experimental
- Input — in stable 1.3
- Inspector — experimental
- IO — in stable 1.3
- LayerTree — experimental
- Log — in stable 1.3
- Media — experimental
- Memory — experimental
- Network — in stable 1.3
- Overlay — experimental
- Page — in stable 1.3
- Performance — in stable 1.3
- PerformanceTimeline — experimental
- Preload — experimental
- Profiler — in stable 1.3, v8 inspector
- PWA — experimental
- Runtime — in stable 1.3, v8 inspector
- Schema — deprecated, v8 inspector
- Security — in stable 1.3
- ServiceWorker — experimental
- SmartCardEmulation — experimental
- Storage — experimental
- SystemInfo — experimental
- Target — in stable 1.3
- Tethering — experimental
- Tracing — in stable 1.3
- WebAudio — experimental
- WebAuthn — experimental
- WebMCP — experimental

## Miscellaneous

### Awesome Chrome DevTools: protocol resources

Published by Chrome DevTools team, community list. Source page: [Awesome Chrome DevTools: protocol resources](https://github.com/ChromeDevTools/awesome-chrome-devtools).

Awesome tooling and resources in the Chrome DevTools ecosystem

- Learning
- DevTools tooling and ecosystem
- Chrome DevTools Protocol
- Using DevTools frontend with other platforms
- Applications
- DevTools Extensions
- Alumni

- [Dev Tips](https://umaar.com/dev-tips/) - Large collection of tips as animated gifs.
- [DevTools Tips](https://devtoolstips.org/) - Collection of illustrated tips as mini tutorials.
- [Can I DevTools?](https://www.canidev.tools/) - Various workflows, documented. Also a weekly tips & tricks [newsletter](https://canidevtools.substack.com/).
- [Web cheatcodes](https://codepo8.github.io/web-cheatcodes/) - Browser developer tools for non-developers.
- [Dear Console](https://codepo8.github.io/dearconsole) - A collection of snippets to use in the browser console.
- [Chrome Secret Menus](https://github.com/sparkyrider/chrome-secret-menus) - Comprehensive guide to internal pages and diagnostic tools in Chrome.
- [Front-end Debugging Tools Handbook](https://github.com/lala-hakobyan/front-end-debugging-handbook) - Practical guide to mastering front-end debugging tools, from Chrome DevTools and framework extensions to AI-enhanced IDE debugging.

- [immutable-devtools](https://github.com/andrewdavey/immutable-devtools) - Custom formatter for Immutable-js values.

- [betwixt](https://github.com/kdzwinel/betwixt) - System level network proxy, providing inspection via Network panel.

- [call-trace](https://github.com/brendankenny/call-trace) - Can instrument your JS with hooks, and then generate a `.cpuprofile` of the of the complete (non-sampled) execution. View either time or call counts.
- [cpuprofilify](https://github.com/thlorenz/cpuprofilify) - Converts output of various profiling/sampling tools to the `.cpuprofile` format.
- [Wishbone Python framework](https://wishbone.readthedocs.io/en/latest/misc/profiling.html) - Profiling data can export as `.cpuprofile`.

- [snapline](https://github.com/pmdartus/snapline) - Converts timeline screenshots to gif.

- [DevTools Timeline Viewer](https://chromedevtools.github.io/timeline-viewer/) - Share URLs of your timeline recordings.

- [VS Code - Debugger for Chrome](https://github.com/Microsoft/vscode-chrome-debug/) - Breakpoint debugging in VS Code.
- [VS Code - Elements for Microsoft Edge](https://github.com/microsoft/vscode-edge-devtools) - Elements panel inside VS Code.
- [ChromeREPL](https://github.com/acarabott/ChromeREPL) - Within Sublime Text, use the Chrome console.
- [Sublime Web Inspector](http://sokolovstas.github.io/SublimeWebInspector/) - JavaScript Breakpoint debugging right in Sublime Text.
- [WebStorm/JetBrains Chrome Extension](https://www.jetbrains.com/help/webstorm/2017.1/configuring-javascript-debugger-and-jetbrains-chrome-extension.html) - The WebStorm IDE can debug JavaScript, view the DOM tree, and edit HTML, CSS and JS live.

- [ChromeDevTools/devtools-protocol](https://github.com/chromedevtools/devtools-protocol) - Canonical location of the protocol JSON. Issue tracker for protocol bugs. TypeScript types.
- [DevTools Protocol API Docs](https://chromedevtools.github.io/devtools-protocol/) - Easy browsable UI for exploring the protocol's domains, methods and events.

- [chrome-remote-interface Wiki](https://github.com/cyrus-and/chrome-remote-interface/wiki) - Many useful recipes.
- [Chrome Protocol Proxy](https://github.com/wendigo/chrome-protocol-proxy) - Tool for debugging clients using devtools protocol.

- [Puppeteer](https://github.com/GoogleChrome/puppeteer/) - Node.js offering a high-level API to control headless Chrome over the DevTools Protocol. See also [awesome-puppeteer](https://github.com/transitive-bullshit/awesome-puppeteer).
- [Playwright](https://github.com/microsoft/playwright) - Library to automate Chromium, Firefox and WebKit with a single API. Available for Node.js, Python, .Net, Java. See also [awesome-playwright](https://github.com/mxschmitt/awesome-playwright).

- JavaScript/Node.js: [chrome-remote-interface](https://github.com/cyrus-and/chrome-remote-interface)
- TypeScript/Node.js: [chrome-debugging-client](https://github.com/TracerBench/chrome-debugging-client)
- TypeScript/Node.js: [noice-json-rpc](https://www.npmjs.com/package/noice-json-rpc) - A proxy-based implementation to expose the CDP as its API.
- TypeScript/Node.js: [Taiko](https://github.com/getgauge/taiko/)
- TypeScript/Node.js: [Lumen](https://github.com/omxyz/lumen) - Vision-first browser agent with self-healing deterministic replay over CDP.
- Rust: [Rust Headless Chrome](https://github.com/atroche/rust-headless-chrome/)
- Java: [chrome-devtools-java-client](https://github.com/kklisura/chrome-devtools-java-client)
- Java: [jvppeteer](https://github.com/fanyong920/jvppeteer) - Headless Chrome For Java
- Python: [PyCDP](https://github.com/hyperiongray/python-chrome-devtools-protocol) - Pure-Python, sans-IO wrappers. See also the [Trio CDP driver](https://github.com/hyperiongray/trio-chrome-devtools-protocol)
- Python: [chromewhip](https://github.com/chuckus/chromewhip) - drop-in replacement for the `splash` service
- Python: [pyppeteer](https://github.com/pyppeteer/pyppeteer) - Puppeteer port
- Python: [ChromeController](https://github.com/fake-name/ChromeController) - high-level browser mgmt
- Go: [chromedp](https://github.com/chromedp/chromedp) - High-level actions and tasks for driving browsers
- Go: [cdp](https://github.com/mafredri/cdp)
- Go: [gcd](https://github.com/wirepair/gcd)
- Go: [godet](https://github.com/raff/godet)
- Go: [Rod](https://github.com/go-rod/rod)
- C#/.NET: [Puppeteer Sharp](https://github.com/hardkoded/puppeteer-sharp) - Puppeteer port
- C#/dotnet: [chrome-dev-tools](https://github.com/BaristaLabs/chrome-dev-tools) - Protocol wrapper generator that can be customized by editing handlebars templates. Includes .Net Core template.
- C#/.NET: [dotnet-chrome-protocol](https://github.com/seclerp/dotnet-chrome-protocol) - A runtime library and schema code generation tools for Chrome DevTools Protocol support in C#/.NET.
- Ruby: [Ferrum](https://github.com/route/ferrum) - high-level API to control Chrome in Ruby
- Ruby: [Cuprite](https://github.com/machinio/cuprite) - Capybara driver
- Kotlin: [chrome-reactive-kotlin](https://github.com/wendigo/chrome-reactive-kotlin) - reactive (rxjava 2.x), low-level client library in Kotlin
- Kotlin: [chrome-devtools-kotlin](https://github.com/joffrey-bion/chrome-devtools-kotlin) - A coroutine-based client library, providing low-level CDP primitives and high-level extensions.
- Clojure: [clj-chrome-devtools](https://github.com/tatut/clj-chrome-devtools) - The CDP wrapper API is autogenerated and will be updated when CDP protocol changes.
- Clojure: [cuic](https://github.com/milankinen/cuic) - Providing a high-level API for UI test automation over the DevTools Protocol.
- PHP: [chrome-devtools-protocol](https://github.com/jakubkulhan/chrome-devtools-protocol) - A PHP client library for the protocol.
- PHP: [PuPHPeteer](https://github.com/rialto-php/puphpeteer) - PHP bridge to node Puppeteer

- [devtools-remote-debugger](https://github.com/Nice-PLQ/devtools-remote-debugger) - Use devtools against a webpage; a CDP agent implemeted in client-side JS.
- [Inspect](https://inspect.dev/) - Use devtools against iOS and Android, easily. Browser and Webviews. (closed source)

- [Facebook Stetho](https://github.com/facebook/stetho) - Native Android debugging with Chrome DevTools.
- [j2v8-debugger](https://github.com/AlexTrotsenko/j2v8-debugger) - Debugging JavaScript running in [J2V8](https://github.com/eclipsesource/J2V8) with Chrome DevTools.

- [Dirac](https://github.com/binaryage/dirac) - Debugging of ClojsureScript.

- [PonyDebugger](https://github.com/square/PonyDebugger) - Remote network and data debugging iOS apps with Chrome DevTools.

- [ndb](https://github.com/GoogleChromeLabs/ndb) - An improved Node.js debugging experience with the DevTools Frontend.
- [Debugging Node.js with Chrome DevTools](https://medium.com/@paul_irish/debugging-node-js-nightlies-with-chrome-devtools-7c4a1b95ae27) - Guide on using the full debugging and profiling support in Node v6.3+.
- [thetool](https://github.com/sfninja/thetool) - CPU, memory, coverage, type profiling with Node.
- [chrome-devtools-frontend](https://www.npmjs.com/package/chrome-devtools-frontend) - Mirror of the frontend that ships in Chrome.

- [ruby/debug](https://github.com/ruby/debug) - Debugging functionality for Ruby.

- [BrowserBox](https://github.com/BrowserBox/BrowserBox) - Embed Chrome in a web page, largely powered by DevTools and supporting multiuser browsing, remote DevTools, audio, and documents like `.docx`, `.pdf`, and more.
- [Puppetromium](https://github.com/dosyago/puppetromium) - A proof-of-concept web browser built with Puppeteer, written in Node.js, HTML and CSS, with 0% client-side JavaScript.

- [dn](https://github.com/dosyago/dn) - Archive and index pages you browse for offline viewing and search, implemented using the `Fetch` domain's interceptions, and works with any Chromium-based browser.

- [Clockwork](https://chromewebstore.google.com/detail/clockwork/dmggabnehkmmfmdffgajcflpdjlnoemp?hl=en) - View PHP application profiling data.
- [RailsPanel](https://chromewebstore.google.com/detail/railspanel/gjpfobpafnhjhbajcjgccbbdofdckggg?hl=en-US) - View Ruby on Rails application profiling data.
- [React Developer Tools](https://chromewebstore.google.com/detail/react-developer-tools/fmkadmapgofadopljbjfkapdkoienihi) - Inspect the React component hierarchies.
- [Ember.js Inspector](https://chromewebstore.google.com/detail/ember-inspector/bmdblncegkenkacieihfhpjfppoconhi) - Allows you to inspect Ember.js objects in your application.
- [Vue.js Developer Tools](https://github.com/vuejs/vue-devtools) - Inspect Vue.js components and manipulate their data.
- [Angular DevTools](https://chromewebstore.google.com/detail/angular-devtools/ienfalfjdbdpebioblfackkekamfmbnh) - Debugging and Profiling for Angular applications.
- [Backbone Debugger](https://chromewebstore.google.com/detail/backbone-debugger/bhljhndlimiafopmmhjlgfpnnchjjbhd) - Inspect a Backbone application's views, models, events, and routes.
- [Redux Devtools](https://chromewebstore.google.com/detail/redux-devtools/lmhkpmbekcpmknklioeibfkpmmfibljd) - Inspect Redux with actions history, undo and replay.
- [Insight](https://github.com/3Dparallax/insight/) - A WebGL debugging toolkit which enables more productive WebGL development and more efficient WebGL applications.
- [BEM devtools](https://github.com/escaton/bem-chrome-devtools) - Inspect BEM entities expressed in `i-bem` framework.
- [Web Component DevTools](https://chromewebstore.google.com/detail/web-component-devtools/gdniinfdlmmmjpnhgnkmfpffipenjljo) - Inspect, modify and observe Web Components on page.

- [Material UI Theme](https://chromewebstore.google.com/detail/material-devtools-theme-c/jmefikbdhgocdjeejjnnepgnfkkbpgjo) - Provides various Material Design inspired themes.

- [sloth](https://github.com/denar90/sloth) - Chrome extension allows to enable and save CPU and network throttling for selected tabs.
- [TracerBench](https://github.com/TracerBench/tracerbench) - A controlled performance benchmarking tool for web applications, providing clear, actionable and usable insights into performance deltas.

- [Puppeteer IDE](https://github.com/gajananpp/puppeteer-ide-extension) - Standalone Puppeteer playground in browser's developer tools.
- [k6 browser](https://github.com/grafana/xk6-browser) - Browser automation and end-to-end web testing tool that interacts with browsers and collects frontend performance metrics.

Old projects, likely not maintained any longer… But still cool.

- [Remote Debug Gateway](https://github.com/RemoteDebug/remotedebug-gateway) - Allows you to connect a client to multiple browsers at once.
    - Multiuser DevTools: [DevTools Remote](https://github.com/auchenberg/devtools-remote) - Remotely debug someone else's browser.
- [DevTools Backend](https://github.com/christian-bromann/devtools-backend) - Standalone implementation of the Chrome DevTools backend to debug arbitrary web environments.
- Python CDP driver: [pychrome](https://github.com/fate0/pychrome) - low level CDP transport handler
- [ios-webkit-debug-proxy](https://github.com/google/ios-webkit-debug-proxy) - Exposes Mobile Safari & UIWebView instances via the CDP.
    - [Remote Debug iOS WebKit adapter](https://github.com/RemoteDebug/remotedebug-ios-webkit-adapter) - Builts upon ios-webkit-debug-proxy and translates WebKit's Remote Debugging Protocol API to the CDP.
- [IE Diagnostics Adapter](https://github.com/Microsoft/IEDiagnosticsAdapter) - Protocol adaptor for Microsoft IE 11 to CDP.
- [go-debugger-devtools](https://github.com/allada/go-debugger-devtools)
