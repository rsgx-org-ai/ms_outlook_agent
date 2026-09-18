# RSGx Ideas — Outlook add-in

A small Outlook add-in that adds an **RSGx Ideas** button to the ribbon. Click it
and a side pane opens where you can jot down an idea or a note without leaving
your inbox. Each entry is timestamped, can carry the subject and sender of the
email you were reading, and can be copied straight into the
[AI Hub idea funnel](https://ai.rsgx.com/submit-request).

It is a **prototype**: plain HTML/CSS/JS in [`addin/`](addin/), no build step,
no server of its own. The files are published to GitHub Pages and Outlook loads
the pane from there.

**Manifest URL (this is what you give Outlook):**

```
https://rsgx-org-ai.github.io/ms_outlook_agent/manifest.xml
```

| File | Purpose |
| --- | --- |
| `addin/manifest.xml` | What you give Outlook. Declares the ribbon button and where the pane lives. |
| `addin/taskpane.html` / `.css` / `.js` | The pane itself. Uses [Office.js](https://learn.microsoft.com/office/dev/add-ins/) from Microsoft's CDN. |
| `addin/assets/icon-*.png` | Ribbon and store icons (16/32/64/80/128 px), rendered from `icon.svg`. |
| `.github/workflows/pages.yml` | Publishes `addin/` to GitHub Pages on every push to `main`. |

## Quick start

1. Make sure GitHub Pages is on for this repository (see [Hosting](#hosting-github-pages)).
2. In Outlook, add a custom add-in **from URL** and paste the manifest URL above.
3. Open any email and click **RSGx Ideas** on the ribbon.

## Turn it on yourself (sideload)

You do not need IT for this. A custom add-in installed this way is visible only
to you.

**Classic Outlook for Windows** (the ribbon with *File · Home · Send/Receive …*)

1. In Outlook: **Home → Get Add-ins**. On some builds the button is
   **All Apps → Add apps**, or **File → Manage Add-ins**, which opens the same
   dialog in a browser.
2. In the dialog choose **My add-ins** (left menu), scroll to **Custom Addins**,
   then **+ Add a custom add-in → Add from URL…**.
3. Paste `https://rsgx-org-ai.github.io/ms_outlook_agent/manifest.xml`, click
   **OK**, accept the warning, and click **Install**.
   (If your build only offers *Add from file…*, open that URL in a browser,
   save it as `manifest.xml`, and pick the file instead.)
4. Open any email. **RSGx Ideas** appears on the Home ribbon (in the
   *RSGx AI Hub* group). Click it to open the pane; on newer builds you can
   pin the pane so it stays open as you move between emails.

**New Outlook for Windows / Outlook on the web**

1. Open <https://outlook.office.com>, open an email, and click the **Apps**
   icon (or **…** → *Get Add-ins*).
2. **My add-ins → Custom Addins → + Add a custom add-in → Add from URL…**,
   paste the manifest URL, **Install**.
3. The button shows up under **Apps** when reading a message.

Sideloaded add-ins can take a minute to appear; restarting Outlook helps.

## Roll it out to everyone (Microsoft 365 admin)

An admin can push it to a group or the whole company so nobody has to sideload:

1. **Microsoft 365 admin center → Settings → Integrated apps → Deploy Add-in**
   (on some tenants: *Upload custom apps*, or *Settings → Add-ins → Deploy Add-in*).
2. Choose **Office Add-in → Provide link to manifest file** and paste the
   manifest URL (or download it and use *Upload manifest file from device*).
3. Choose who gets it (a test group first is sensible) and whether it is
   *Fixed* (always on), *Available* (users opt in) or *Optional*.
4. It appears for those users within a few hours.

Updating the add-in later: bump `<Version>` in `addin/manifest.xml` and
re-deploy the manifest. Changes to the pane itself (`taskpane.*`) go live as
soon as the Pages workflow finishes; Outlook simply loads the new files.

## Hosting (GitHub Pages)

`.github/workflows/pages.yml` publishes the `addin/` folder to
`https://rsgx-org-ai.github.io/ms_outlook_agent/` on every push to `main`.

Two things the workflow cannot do for you:

- **A repo admin must turn Pages on once:** *Settings → Pages → Build and
  deployment → Source = "GitHub Actions"*. Until then the deploy job fails with
  a "Pages not enabled" error; re-run it (or push again) once the setting is
  saved.
- **This repository is private.** GitHub Pages on a private repository is only
  available on GitHub Enterprise Cloud. If the org is not on that plan, either
  make the repository public (there is nothing sensitive in it: the pane is
  static and Outlook has to fetch it anonymously anyway) or host `addin/` on
  any other HTTPS origin and change the URLs in `manifest.xml` to match.

Outlook requires every URL in the manifest to be HTTPS and reachable without a
sign-in. GitHub Pages satisfies both.

## Local development

Outlook only loads add-ins over **HTTPS**, so a plain `http://localhost` will
not do. The simplest route is the Office Add-ins dev tooling, which provisions
a trusted localhost certificate:

```bash
npx office-addin-dev-certs install          # one-off: trusted cert for localhost
npx office-addin-manifest validate addin/manifest.xml
npx http-server addin -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```

Then **swap the URL in the manifest**: replace every
`https://rsgx-org-ai.github.io/ms_outlook_agent` with `https://localhost:3000`
(the pane, the icons and the first `<AppDomain>`), save it as
`addin/manifest.dev.xml` (git-ignored) and sideload that copy. The pane is at
`https://localhost:3000/taskpane.html`.

Opening `addin/taskpane.html` in an ordinary browser also works: without Outlook
it falls back to `localStorage` and shows "preview mode" in the footer.

`npx office-addin-debugging start addin/manifest.dev.xml` will sideload and
open a debugger, if you have the Office tooling installed.

## How it relates to the AI Hub

- **Idea funnel.** "Send to AI Hub" opens
  <https://ai.rsgx.com/submit-request> with the entry (and its email context)
  already in the description box, via `?problem=…&source=outlook`. The hub
  reads both keys and shows a short "brought in from Outlook" note. The text
  is still copied to the clipboard as a fallback in case a proxy strips the
  query string. The user picks a function and submits; nothing is posted
  automatically.
- **Links open in the system browser** via `Office.context.ui.openBrowserWindow`
  when Outlook provides it, falling back to `window.open` in a plain browser.

## Known limits of the prototype

- **Entries live in your mailbox, not in the hub.** They are stored in Office
  *roaming settings*, which follow you across Outlook clients on the same
  mailbox but are not visible to anyone else and are not backed up anywhere.
  Outlook gives each add-in a **32 KB** roaming-settings budget, which is
  roughly 8 long notes or many short ones; when it is full, Save reports an
  error and you need to delete older entries.
- **No hub API write yet.** Nothing is posted to the funnel automatically; the
  prefilled-link flow above is the bridge, and the user still presses Submit.
  A signed-in `POST` from the pane would need an auth story for the Outlook
  webview (SSO token exchange), deliberately out of scope for this prototype.
- **Long entries are truncated in the link** at 1800 characters, to stay inside
  the URL length older Outlook webviews and proxies handle. The full text is on
  the clipboard, and the hub itself accepts up to 4000.
- **Context is subject + sender only.** No email body, attachments or links
  are captured.
- **Icons are placeholders** (a red tile with a lightbulb).
- **Unverified in a real Outlook client.** Tested in a browser with a mocked
  Office host only; Microsoft's online manifest validator was not reachable
  from the build environment. The first real-client run may surface layout
  tweaks for the classic Outlook task pane width (~320 px).
