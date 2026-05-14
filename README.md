# LSPD Dienstblatt

Node.js/Express Dienstblatt mit statischem Frontend und JSON-Speicher.

## Render Deployment

Build Command:

```bash
npm ci --omit=dev
```

Start Command:

```bash
npm start
```

Environment Variables:

```text
NODE_ENV=production
DATA_DIR=/opt/render/project/src/storage
```

Persistent Disk:

```text
Mount Path: /opt/render/project/src/storage
Size: 1 GB
```

Die Datei `storage/dienstblatt.json` ist absichtlich nicht im Repository enthalten. Render muss fuer diesen Ordner einen Persistent Disk verwenden, damit Daten dauerhaft gespeichert bleiben.
