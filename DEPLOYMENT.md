# SCRouteReady Deployment Notes

SCRouteReady is a static browser-only proof of concept. Build output is generated from source and should not be committed.

## Build

```bash
npm install
npm run build
```

Deploy the generated `dist/` directory with the static host after source review, dependency review, and proof-of-concept approval.

## Access Control

Configure any allowlist, identity provider, or Cloudflare Access policy in the hosting dashboard. Do not store real email addresses, account identifiers, API tokens, or credentials in this repository.

Example placeholder emails:

- `your-email@example.com`
- `betty-email@example.com`

## Data Handling

The application processes selected workbook files locally in the browser. It has no backend, upload endpoint, database, analytics, telemetry, or intentional route-data persistence.

Do not commit route exports, generated PDFs, screenshots containing route/customer information, environment files, credentials, or local machine paths.
