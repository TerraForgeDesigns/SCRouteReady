# SCRouteReady

SCRouteReady is a proof-of-concept browser application for preparing P21 Sales Order and Transfer Staging Reports for Pull Sheet and Pallet Key printing.

Subtitle: Pull Sheet & Pallet Key Generator

This is not yet an officially approved Stylecrest application.

## Current Development Stage

This stage provides the secure application foundation, polished report-selection interface, local `.xlsx` workbook reading, automatic report detection, Pull Sheet generation, automatic ABC Pallet Key generation when Sales stops need it, browser preview, native printing, and active-session clearing.

The generated Pull Sheet and Pallet Key pages are temporary browser-rendered HTML. SCRouteReady does not generate Excel output, automatically create PDFs, or download files.

## Intended Workflow

Drop Report(s) -> Generate Pull Sheet and Keys -> Print -> Finish & Clear

The P21 Excel workbook is the input format. Employees may drop one Sales Order Staging Report, one Transfer Staging Report, or one of each for the same truck. SCRouteReady identifies the report type from the `Results` worksheet headers; no report-type selector, column mapping, or preset is required.

Sales-only trucks use the Sales Order staging columns. Transfer-only trucks use Transfer staging columns and render product rows with the visible Type value `Xfer`. The Pull Sheet `PT/Xfer #` column shows Sales `Pick Ticket No` values for Sales rows and `Transfer No` values for Transfer rows. Mixed Sales / Transfer trucks are merged only when the normalized `Route from Descartes` route set matches across both reports. The output is a temporary browser-rendered HTML document optimized for landscape printing, not a generated Excel workbook. The browser print dialog may allow a user to save as PDF, but SCRouteReady will not automatically create or save PDF files.

## Local Browser Processing

The selected staging report or matching pair of staging reports is read with `File.arrayBuffer()` and parsed locally with SheetJS Community Edition. The application has no backend, no database, no file-upload endpoint, no authentication, no analytics, and no telemetry.

SCRouteReady intentionally keeps no previous-task recovery feature and does not use browser storage APIs. Finish & Clear removes active application references and resets the file input, while browser memory cleanup remains controlled by the browser and operating system.

## Project Structure

- `src/components` - UI components
- `src/hooks` - active workbook session management
- `src/services` - workbook reading, report detection, and Pull Sheet generation services
- `src/utils` - validation and formatting helpers
- `src/types` - shared TypeScript types
- `src/styles` - screen and future print styling
- `vendor` - vendored SheetJS Community Edition package

## Development Setup

Install dependencies:

```bash
npm install
```

Run the development server:

```bash
npm run dev
```

Run linting:

```bash
npm run lint
```

Run TypeScript checks:

```bash
npm run typecheck
```

Create a production build:

```bash
npm run build
```

Preview the production build:

```bash
npm run preview
```

## Current Limitations

- Accepts only `.xlsx` files up to 20 MB each.
- Accepts at most two reports per selection: one Sales Order report and one Transfer report.
- Requires a worksheet named `Results`, case-insensitive.
- Pull Sheet generation requires the expected P21 Sales Order or Transfer Staging Report columns.
- Mixed generation requires the same normalized Descartes route set in both reports.
- Browser pagination can vary by printer driver and print-dialog settings.
- SCRouteReady cannot prevent a user from manually overriding Landscape in the browser print dialog.
- No Excel output, PDF generation, automatic downloads, backend, cloud storage, or user accounts are implemented.

## Future Work

Future development may refine route-specific formatting, operational edge cases, internal hosting, and corporate review feedback.

The final printed Pull Sheet and ABC Key pages must use landscape `@page` print CSS, printer-safe margins near 0.25 to 0.35 inches, spreadsheet-style borders, repeated column headers, large readable body text around 11 to 12 points, larger bold Stop and ABC labels, practical row page-break control, clear page breaks between document sections, and black-and-white readability.

## Security Review Notes

Review `SECURITY.md` and `security-headers.example.txt` before any internal deployment. Dependency and source review are recommended before proof-of-concept approval or corporate hosting.
