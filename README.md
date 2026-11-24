# google-script

Imported Apps Script project. The latest available revision (version 6) is in
the repository root. Historical snapshots are in `versions/version-1` through
`versions/version-6`.

## Setup

1. Install dependencies:
   npm install
2. Authenticate with Google Apps Script:
   npm run login
3. Push the latest revision:
   npm run push

## Files

- `Code.gs` — latest Apps Script source
- `appsscript.json` — Apps Script project configuration
- `.clasp.json` — local clasp project link (ignored by Git)
- `versions/` — source and manifest snapshots for every available version
- `.claspignore` — keeps historical snapshots out of uploads

## Useful commands

- `npm run open` — open the project in the Apps Script editor
- `npm run deploy` — deploy the latest version

Each `versions/version-N` directory is linked to the same remote Apps Script
project and contains that revision's source and manifest. The root
`.claspignore` prevents these snapshots from being uploaded by mistake. To
inspect or pull a specific historical revision, work from its directory.
