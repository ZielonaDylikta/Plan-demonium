# Martin's Plan-demonium

Clean GitHub-ready copy of the Planner Stats app.

## Included files
- `server.js`
- `public/` (frontend app)
- `package.json` / `package-lock.json`
- `.env.example`
- `data/.gitkeep` (keeps the folder; runtime snapshots are ignored)

## Not included
- `.env` (secrets)
- `node_modules/`
- `data/snapshots.json` (runtime-generated)
- local editor/system files

## Run locally
1. Copy env template:
   ```bash
   cp .env.example .env
   ```
2. Fill values in `.env`:
   - `CLIENT_ID`
   - `TENANT_ID`
   - `CLIENT_SECRET`
   - `GROUP_ID`
3. Install dependencies and run:
   ```bash
   npm install
   npm start
   ```

## Create and push a new GitHub repo
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-user>/<your-repo>.git
git push -u origin main
```

## License
Licensed under the Apache License 2.0. See [LICENSE](./LICENSE).

Copyright 2026 ZielonaDylikta.
