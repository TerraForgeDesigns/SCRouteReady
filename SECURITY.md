# SCRouteReady Security Model

SCRouteReady is currently a static, browser-only proof of concept. It has no backend, no database, no authentication system, and no file-upload endpoint.

## Local Processing

- The employee deliberately selects a `.xlsx` P21 route export.
- The browser reads that file with `File.arrayBuffer()`.
- SheetJS parses the workbook locally inside the browser runtime.
- The original export remains on the employee's computer in its selected location.
- SCRouteReady does not modify the original workbook.

## No Intentional Retention

The application must not use `localStorage`, `sessionStorage`, `IndexedDB`, Cache Storage, service workers, cookies, analytics, telemetry, tracking pixels, or any route-data recovery feature.

Workbook references are held only in temporary JavaScript memory for the active task. Finish & Clear removes active file, workbook, buffer, metadata, error, and preview references controlled by the application and resets the native file input.

Active-session clearing is not a guarantee of cryptographic memory erasure. Browser engines and operating systems control lower-level memory allocation, reuse, and process cleanup.

## Network Model

The production build is intended to run as a fully static application with locally bundled dependencies. Route files and parsed workbook data must never be transmitted by the application.

Windows printer spooler behavior, browser print behavior, and any user-selected Save as PDF action are outside the site's control. Saving as PDF should only happen through the browser's print dialog when a user chooses it.

## Deployment Notes

Before corporate deployment, SCRouteReady should receive source review, dependency review, static hosting review, and security-header review. Internal hosting is recommended after proof-of-concept approval.
