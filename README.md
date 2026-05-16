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
DISCORD_APPLICATION_ID=
DISCORD_PUBLIC_KEY=
DISCORD_BOT_TOKEN=
DISCORD_SERVER_ID=
DISCORD_OAUTH_REDIRECT_URL=
```

Persistent Disk:

```text
Mount Path: /opt/render/project/src/storage
Size: 1 GB
```

Die Datei `storage/dienstblatt.json` ist absichtlich nicht im Repository enthalten. Render muss fuer diesen Ordner einen Persistent Disk verwenden, damit Daten dauerhaft gespeichert bleiben.

## Discord Sync

Die Discord-Werte koennen im IT-Reiter gepflegt werden. Lokal koennen sensible Werte alternativ in `.env` liegen; diese Datei wird nicht committed. Eine Vorlage liegt in `.env.example`.
