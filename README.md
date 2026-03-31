# Martin's Plan-demonium v2.0

GitHub-ready copy of the Plan-demonium Teams + Planner dashboard.

## What this app does
Plan-demonium connects to Microsoft Planner (via Microsoft Graph) and presents a focused dashboard inside Teams. It helps teams monitor plan health, track progress, and quickly spot workload or risk imbalances across tasks.

## Core features
- Teams tab experience with Teams SSO authentication
- Planner KPI and status overview across plans
- Workload and risk insights for faster triage
- Scope templates for saved plan selections
- Channel-to-template bindings for context-aware dashboards
- Admin unlock flow (PIN + session)
- Feedback capture with protected CSV export
- Optional daily snapshot scheduling

## What this repo contains
- `server.js` (Express + Microsoft Graph backend)
- `public/` (frontend app)
- `teams/manifest.json` (sanitized Teams app manifest template)
- `.env.example` (environment variable template)
- `package.json` / `package-lock.json`

## Prerequisites
- Node.js 20+
- Microsoft Entra app registration with Graph permissions for Planner and Teams channels
- A target Microsoft 365 Group / Team (`GROUP_ID`)

## Local setup
1. Install dependencies:
   ```bash
   npm install
   ```
2. Create local env file:
   ```bash
   cp .env.example .env
   ```
3. Fill required values in `.env`:
   - `CLIENT_ID`
   - `TENANT_ID`
   - `CLIENT_SECRET`
   - `GROUP_ID`
   - `AUTH_AUDIENCE` (usually same as app registration client ID)
   - `AUTH_ALLOWED_TENANTS` (tenant ID or `*`)
4. Optional but recommended:
   - `ADMIN_PIN`
   - `ADMIN_SESSION_SECRET`
   - `FEEDBACK_ENABLED`
   - `FEEDBACK_DEV_KEY`
5. Start the app:
   ```bash
   npm start
   ```

## Teams manifest (template)
`teams/manifest.json` is intentionally sanitized. Before packaging it for Teams, replace placeholder values:
- `id`
- `webApplicationInfo.id`
- `webApplicationInfo.resource`
- `developer.websiteUrl`, `privacyUrl`, `termsOfUseUrl`
- tab URLs (`configurationUrl`, `contentUrl`)
- `validDomains`

Ensure these values match your deployed domain and Entra app registration.

## Create Teams app package
1. Update `teams/manifest.json` placeholders.
2. Add icon files in `teams/`:
   - `color.png` (192x192)
   - `outline.png` (32x32)
3. Zip the contents of `teams/` (files inside, not the folder itself).
4. Upload zip to Teams Developer Portal / Apps.

## Notes
- Runtime data is generated under `data/` on first run.


## License
Licensed under Apache License 2.0. See [LICENSE](./LICENSE).

Copyright 2026 ZielonaDylikta.
