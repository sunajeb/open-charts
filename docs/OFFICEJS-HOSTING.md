# Hosting the Office.js add-in

The Office.js add-in is a task pane: a set of static web files plus `manifest.xml`, the file Office
reads to find them. To roll it out to an organization, you host the files on a https server and
deploy the manifest from the Microsoft 365 admin center. Nothing is installed on the computers.
For trying it on one computer, see the [user guide](USER-GUIDE.md#officejs-add-in-windows-web-mac).
For the macro add-in and IT deployment in general, see [IT-DEPLOYMENT.md](IT-DEPLOYMENT.md).

## Build for your host

Pick the https address the files will live at, then build with it:

```sh
npm install
npm run build:officejs -- --base https://example.org/open-charts/
```

- `--base` must be an absolute `https://` address, without a query or fragment. A path is fine. The
  build stops with a message otherwise.
- `--out <folder>` writes somewhere other than `dist/officejs/`. The folder is replaced, so it must be
  new, empty, or an earlier build.
- Every URL in the manifest (task pane, icons, support page) points under the base. The add-in id
  is fixed, so it is the same add-in as a sideloaded copy.

The output folder is what you upload:

| File | What it is |
|---|---|
| `taskpane.html`, `taskpane.js` | the pane: datasheet editor and chart engine |
| `commands.html` | the ribbon button's function file |
| `support.html` | the support page the manifest links to |
| `fonts/` | Liberation Sans, used to measure text (SIL Open Font License, licence file included) |
| `assets/` | icons |
| `engine.js` | the engine as an ES module; the pane does not load it, so uploading it is optional |
| `manifest.xml` | for the admin center; it may be uploaded too, but nothing loads it from the host |

## Hosting requirements

- **https with a certificate the users' computers trust.** Office loads add-ins only over https.
  Any static host works: an internal web server, Azure Static Web Apps or Blob Storage, S3 with a
  CDN, GitHub Pages.
- **Upload the folder as it is**, at the `--base` address. File names are fixed and relative.
- **Content types:** `.html` as `text/html`, `.js` as `text/javascript`, `.png` as `image/png`,
  `.ttf` as `font/ttf` (or `application/octet-stream`), `.xml` as `application/xml`.
- **No CORS headers are needed.** The pane loads its own files from the same address. The only
  other file it loads is Microsoft's `office.js`, from
  `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. Users' computers must be able to
  reach that address.
- **No server code.** The pane asks `fonts/metrics?family=<font name>` once per font that is not
  Arial. On the development server this returns exact font measurements. On a static host it gets a
  404, and the pane then measures with the browser instead and does not ask again. The 404s in your
  logs are expected.
- **Let Office frame the pane.** PowerPoint on the web shows the pane in a frame. Do not send
  `X-Frame-Options` or a `Content-Security-Policy` with `frame-ancestors` that would block it. If
  you set a Content-Security-Policy, it must allow inline scripts (the editor page has them) and
  scripts from `appsforoffice.microsoft.com`.
- **Cache headers.** The files keep the same names in every version. Serve `taskpane.html`,
  `taskpane.js` and `commands.html` with `Cache-Control: no-cache`, so users get a new version as
  soon as you upload it. `fonts/` and `assets/` can be cached for a day or more.

## What the pane sends over the network

- It requests its own files from your host, and `office.js` from Microsoft.
- Your host's logs see those requests. That includes the `fonts/metrics?family=…` requests, which
  carry the names of fonts used in charts.
- Chart data never leaves the computer. Layout runs inside the pane, and the chart is written to
  the presentation through Office.
- On a host other than `localhost`, the pane sends no logs anywhere.

## Check the manifest

The unit tests (`test/officejs.test.ts`) check the manifest's structure, element order and URLs.
Microsoft's validator checks it against the schema and the store rules. It needs internet
access:

```sh
npx office-addin-manifest validate dist/officejs/manifest.xml
```

Then sideload the manifest on one computer (user guide) before you deploy it to everyone.

## Deploy to your organization (centralized deployment)

1. Upload the built folder to the `--base` address. Check in a browser that `<base>taskpane.html`
   and `<base>fonts/LiberationSans-Regular.ttf` load over https.
2. In the Microsoft 365 admin center, go to *Settings › Integrated apps › Upload custom apps*.
   Choose *Office Add-in* and upload `manifest.xml` from the build folder.
3. Assign it to users or groups and deploy.
4. It can take a few hours to reach users. They then find **open-charts** on the **Home** tab in
   PowerPoint for Windows, Mac and the web. It needs a PowerPoint that supports PowerPointApi 1.4.
   Features above that are detected when the pane runs.

**Updates.** Build again with the same `--base` and upload the files. The pane uses the new files
the next time it opens. Upload the new `manifest.xml` in the admin center only when the manifest
changes, for example after an open-charts version change.

**Removal.** Remove the add-in in *Integrated apps*. Charts already in presentations stay as they
are.

## AppSource (later)

Listing it in Microsoft AppSource needs a permanent public host and a privacy policy. It also needs
a Partner Center account and Microsoft's validation. None of that is set up yet. The manifest
already uses the fixed add-in id and https URLs that a listing needs.
