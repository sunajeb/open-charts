# Releases, installers and signing

How the installers are built, how to cut a release, what the owner needs for signing, how IT
deploys open-charts, and how updates work. Users install from the
[README](../README.md#install).

## The installers

| File | For | Built by | Installs |
|---|---|---|---|
| `open-charts-<version>.pkg` | Mac users | `npm run build:mac` (on a Mac) | per user, no admin password |
| `open-charts-<version>.msi` | Windows users | `npm run build:win && npm run build:msi` (on Windows) | per user, no admin rights; `/quiet` for silent |
| `open-charts-<version>-machine.msi` | Windows IT (Intune, SCCM) | the same | per machine, as SYSTEM |

None of them needs Node on the user's computer: the engine is a Node
[single executable application](https://nodejs.org/api/single-executable-applications.html), built
from the official Node binary of the build machine's Node version. The binary is downloaded from
nodejs.org and checked against that release's `SHASUMS256.txt` (`scripts/build-win.ts`).

### Mac: `open-charts-<version>.pkg` (`scripts/build-mac.ts`)

- **Payload**, installed into `~/Library/Application Support/open-charts`:
  - `engine/open-charts`: the engine, universal (arm64 and x86_64 slices joined with `lipo`).
  - `OpenChartsPopup.app`: the Edit data window, also universal.
  - `package/open-charts.ppam` and `package/addin.json`: the add-in and its callback prefix.
  - `package/Uninstall open-charts.app`: the uninstaller.
- **Per user only.** The distribution enables only the `currentUserHome` domain, and the component's
  `auth` is `none`, so Installer asks for no password. No bundle is relocatable, so Installer never
  updates some other copy of the pop-up app instead.
- **preinstall** stops a running engine. **postinstall** runs `engine/open-charts addin install`
  (`src/addin/mac-install.ts`). This writes the add-in to
  `~/Library/Application Support/Microsoft/open-charts/`, compiles the bridge script into
  `~/Library/Application Scripts/com.microsoft.Powerpoint/`, writes the `open-charts` command and
  copies the uninstaller to `~/Applications`. These are the same places an install from source uses.
  The bridge script then runs the executable instead of `node … src/cli.ts`.
- **Registration without starting PowerPoint.** If PowerPoint is running, the add-in is registered
  and loaded right away. Otherwise a LaunchAgent (`io.github.open-charts.register`) checks with
  `pgrep` every 30 s, without starting the engine. Once PowerPoint runs, it calls
  `open-charts addin register` and then removes itself. After 10 failed attempts (for example,
  Automation permission denied) it gives up and says so in `addin.log`.
- **Uninstall:** *Uninstall open-charts* in `~/Applications`, or `open-charts addin uninstall`.
  Either one removes all of the above, the LaunchAgent and the package receipt.
- **Check without installing:** `npm run check:mac-pkg` expands the `.pkg` and unpacks the payload
  into a temporary home folder. It runs the pre- and postinstall scripts the way Installer does,
  with no Node on `PATH` and PowerPoint left alone. It then checks every file, the bridge script, the
  `open-charts` command and the engine (`scripts/smoke-exe.ts`), and finally uninstalls.
  `OPEN_CHARTS_NO_POWERPOINT=1` keeps PowerPoint and launchd out of a test install.

The x86_64 slice runs on Intel Macs and under Rosetta. `build:mac` runs it under Rosetta when
Rosetta is installed, and skips that check otherwise.

### Windows: the MSIs (`scripts/build-msi.ts`, `addin/windows/open-charts.wxs`)

They are built with the [WiX Toolset](https://wixtoolset.org) v5:

```powershell
dotnet tool install --global wix --version 5.0.2
wix extension add --global WixToolset.Util.wixext/5.0.2 WixToolset.UI.wixext/5.0.2
npm run build:win; npm run build:msi
```

- **Per user** (`Scope=perUser`). It installs exactly what `install.ps1` installs, in the same
  places:
  - `%LOCALAPPDATA%\open-charts\open-charts.exe`, `addin.json` and `uninstall.ps1`;
  - `%APPDATA%\Microsoft\AddIns\open-charts.ppam`;
  - the `AddIns\open-charts` key (`Path`, `AutoLoad=1`) and the `Trusted Locations\open-charts` key
    under `HKCU\Software\Microsoft\Office\16.0\PowerPoint`.

  A running engine is ended first (`util:CloseApplication`). Uninstalling removes all of it,
  including what the engine wrote (the log and unpacked resources). The package asks for no
  elevation: its summary information sets the "no UAC elevation" bit, and CI checks for it.
- **Per machine** (`Scope=perMachine`). It installs the same files to
  `%ProgramFiles%\open-charts`. It also creates an Active Setup entry, so Windows runs
  `install.ps1 -SkipEngineCheck` once for each user at their next sign-in (and again when the
  version goes up).
- **CI** (`.github/workflows/test.yml`, windows job) builds both packages. It installs and removes the
  per-user MSI silently, with a running engine. It installs the machine MSI, runs its Active Setup
  step for the runner's user, then removes everything.

## Cutting a release

1. Set the new version in `package.json` and `src/version.ts` (a test keeps them in sync). Add the
   changes to `CHANGELOG.md` and commit.
2. Tag the commit and push the tag: `git tag v0.3.0 && git push origin v0.3.0`.
3. `.github/workflows/release.yml` builds the `.pkg` on macOS and both MSIs on Windows. It runs the
   package checks, signs everything if the secrets below exist, and creates a **draft** release
   with the three installers and `SHA256SUMS.txt`.
4. Install the draft's files on a Mac and a PC (the live acceptance runs, `docs/ACCEPTANCE.md`), then
   **publish** the release. A build without signing secrets is marked *UNSIGNED BUILD: do not
   publish*. The update check only ever sees published releases, so only signed packages are
   offered.

`workflow_dispatch` builds the same files without a release (download them from the run).
Building by hand works too: `npm run build:mac && npm run check:mac-pkg` on a Mac, and
`npm run build:win && npm run build:msi` on Windows.

## Signing: what the owner needs

Unsigned builds work, but macOS and Windows warn before opening them (the README explains how to
continue), and company laptops often block them. Both build scripts print **UNSIGNED** when they
cannot sign.

### Apple: Developer ID and notarization (Mac)

1. Join the **Apple Developer Program** ($99 per year): <https://developer.apple.com/programs/>.
   An organization account needs a D-U-N-S number; an individual account does not. The certificate
   name shows the account name.
2. In *Certificates, Identifiers & Profiles* (or Xcode › Settings › Accounts › Manage
   Certificates), create two certificates:
   - **Developer ID Application**, which signs the engine, the pop-up app and the uninstaller;
   - **Developer ID Installer**, which signs the `.pkg`.
3. Create an **app-specific password** at <https://account.apple.com> (Sign-In and Security), for
   notarization.
4. In Keychain Access, export both certificates with their private keys as one `.p12`. Then add
   these **repository secrets** (GitHub › Settings › Secrets and variables › Actions):

   | Secret | Value |
   |---|---|
   | `APPLE_CERTIFICATES_P12` | `base64 -i certs.p12` |
   | `APPLE_CERTIFICATES_PASSWORD` | the `.p12` password |
   | `APPLE_DEVELOPER_ID` | `Developer ID Application: Your Name (TEAMID)` |
   | `APPLE_INSTALLER_ID` | `Developer ID Installer: Your Name (TEAMID)` |
   | `APPLE_ID`, `APPLE_TEAM_ID`, `APPLE_APP_PASSWORD` | your Apple ID, team ID, app-specific password |

To sign on your own Mac, first store the notarization credentials once:

```sh
xcrun notarytool store-credentials open-charts --apple-id you@example.com --team-id TEAMID
```

Then build with the certificates in your login keychain:

```sh
APPLE_DEVELOPER_ID="Developer ID Application: …" APPLE_INSTALLER_ID="Developer ID Installer: …" \
APPLE_NOTARY_PROFILE=open-charts npm run build:mac
```

With these set, `build:mac` signs the engine with the hardened runtime and the entitlements Node
needs for its JIT (`addin/mac/engine.entitlements`), signs the pop-up app and uninstaller, and signs
the `.pkg`. It then notarizes the `.pkg` (`xcrun notarytool submit --wait`) and staples the ticket.
Check the result with `pkgutil --check-signature` and `spctl --assess --type install -vv`.

### Azure Trusted Signing (Windows)

Azure Trusted Signing (Microsoft calls it Artifact Signing) costs about $10 per month on the Basic
tier. Signatures from it build SmartScreen reputation without an EV certificate.

1. An Azure subscription: <https://portal.azure.com>.
2. Create a **Trusted Signing account** and note its region's endpoint, for example
   `https://eus.codesigning.azure.net/`.
3. **Identity validation** (Public): Microsoft verifies the organization, or the individual where
   individuals are accepted. This takes from hours to days. Check the current eligibility rules in
   the Azure documentation.
4. Create a **certificate profile** (Public Trust).
5. Create an **app registration** (service principal) with a client secret. Give it the role
   *Trusted Signing Certificate Profile Signer* on the account.
6. Add these repository secrets: `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`,
   `AZURE_TRUSTED_SIGNING_ENDPOINT`, `AZURE_TRUSTED_SIGNING_ACCOUNT` and
   `AZURE_TRUSTED_SIGNING_PROFILE`.

With them, CI uses `azure/trusted-signing-action` to sign `open-charts.exe` before the MSIs are
built, so the MSIs contain the signed exe. It then signs both MSIs. Signing never runs for pull
requests, and the release workflow checks every signature with `Get-AuthenticodeSignature`.

Signing the VBA project inside `open-charts.ppam`, so that "trust signed macros" policies accept it,
needs the same certificate and comes with task C3.

## IT deployment

The full guide for IT admins is [IT-DEPLOYMENT.md](IT-DEPLOYMENT.md). In short:

### Windows (Intune)

- **Per machine:** upload `open-charts-<version>-machine.msi` as a *line-of-business app*, or as a
  Win32 app with the install command `msiexec /i open-charts-<version>-machine.msi /quiet` and the
  uninstall command `msiexec /x {product code} /quiet`. Assign it to device groups. Each user gets
  the add-in at their next sign-in, through Active Setup running `install.ps1`. A user who is signed
  in during the install gets it after signing out and in again.
- **Per user:** assign `open-charts-<version>.msi` to user groups in *user* install context.
  It needs no admin rights.
- **Uninstall:** removing the machine MSI removes Program Files and the Active Setup entry. Each user's
  copy stays until `%LOCALAPPDATA%\open-charts\uninstall.ps1` runs in that user's context (an Intune
  script). The per-user MSI removes everything itself.
- **Macro and trust policy:**
  - The add-in is a macro add-in. `install.ps1` and the MSI add `%APPDATA%\Microsoft\AddIns\` as a
    PowerPoint trusted location under HKCU.
  - If policy manages trusted locations, add that folder by policy instead.
  - `VBAWarnings=4` ("disable all macros") blocks the add-in.
  - Application control (AppLocker, WDAC): allow `%LOCALAPPDATA%\open-charts\open-charts.exe`, by
    publisher once it is signed.

  More in [windows.md](windows.md#corporate-laptops).

### Mac (Jamf and other MDM)

The `.pkg` installs only into the user's home folder, so an MDM that runs `installer` as root
against `/` is refused. Run it as the signed-in user instead, for example in a Jamf policy script:

```sh
user=$(stat -f %Su /dev/console); uid=$(id -u "$user")
launchctl asuser "$uid" sudo -u "$user" installer -pkg /path/open-charts-<version>.pkg -target CurrentUserHomeDirectory
```

A machine-wide Mac package is not built yet (task C6).

## Beta releases for test users

Push a tag with a pre-release suffix, for example `git tag v1.0.0-beta.1 && git push origin
v1.0.0-beta.1`. The release workflow builds the same installers and publishes them at once as a
GitHub **pre-release**, signed or not, with `docs/BETA.md` as its notes: a test-user disclaimer and
the steps past macOS's and Windows' warnings for unsigned software. The update check never offers a
pre-release. Share the release page link with testers; feedback comes back through issues and
**Help › Report a problem**.

## Updates

**Today (client side, off by default).**

- `open-charts addin update` checks now. `open-charts addin update on` makes the add-in server check
  once a day as it starts; this sets `"updateCheck": true` in the add-in home's `settings.json`.
- The check reads the GitHub Releases API:
  `https://api.github.com/repos/sunajeb/open-charts/releases/latest`. That endpoint returns only the
  latest **published**, non-draft, non-pre-release release.
- The check compares that release's version with the engine's and picks this platform's installer:
  the `.pkg`, or the per-user `.msi`.
- It stores the result in `update.json`, logs `UPDATE … is available` to `addin.log`, and serves it
  at `GET /api/addin/update`.
- It never downloads or runs anything. The request carries only `User-Agent: open-charts/<version>`.
- Offline or rate-limited, it keeps the last result and tries again on the next start.
- Code: `src/update.ts`; tests: `test/installers.test.ts`.

**Signed packages only.** Releases are created as drafts, and the owner publishes only signed
builds (see Cutting a release). The endpoint never returns a draft, so the client can only be
offered a signed package.

**Installing an update (the user's own action).** `open-charts addin update install`:

1. checks for the latest release, as above;
2. downloads this platform's installer and `SHA256SUMS.txt`;
3. verifies the download before opening it:
   - the SHA-256 matches the line for the installer in `SHA256SUMS.txt`;
   - on a Mac, `pkgutil --check-signature` shows a Developer ID Installer certificate with our team
     ID, and `spctl --assess --type install` accepts the package;
   - on Windows, `Get-AuthenticodeSignature` is `Valid` with our certificate subject;
4. opens the verified installer (Installer on a Mac, `msiexec` on Windows), where the user confirms.

An installer that fails any check is deleted and never opened. The signing identity (`SIGNER` in
`src/update.ts`) is empty until the signing accounts exist (C3), so until then every download is
refused and users install from the release page.

**Managed machines.** IT turns the check off for everyone with a policy file that users cannot
write: `/Library/Application Support/open-charts/policy.json` on a Mac,
`%ProgramData%\open-charts\policy.json` on Windows, containing `{"updateCheck": false}`. The
policy wins over the user's setting: no check, no download (see docs/IT-DEPLOYMENT.md).

**In the add-in.** **Help › Check for updates** asks once, when clicked, and says what is new and
where to download it (or that IT manages updates). Once the signing identity is known, it can offer
the verified install above as well.
