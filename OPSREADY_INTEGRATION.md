# Automatic OpsReady pull-sheet archiving

Current branch: feature/opsready-integration. The standalone entry and production generation/printing behavior remain unchanged. No commit, push, merge or deployment is performed.

Upload P21 and Generate normally. The Hub version automatically captures the completed .print-document DOM once for each successful generation, freezes it with the exact existing CSS and embedded logo, and saves it asynchronously using a stable generation ID. Print never waits for archival storage. Not saved is explicit, with one-click retry of the same already-generated artifact; there is no generated-file upload, manual metadata form or Finalize action.

Recent Pull Sheets defaults to saved dates from the past 30 days (server UTC). Click date, then document, then Print pull sheet. Archive retains older records and optional search. Existing PDF records remain readable. All save/list/date/retrieval operations enforce the current authenticated Hub role and active facility.

The artifact is self-contained version 1 HTML with frozen actual rendered tables, stop order, ABC keys, inline column widths, full original app/pull-sheet/print CSS and the original PNG logo bytes. Retrieval does not rerun a future generator/renderer or fetch future CSS/assets. The trusted print-control button is added outside the stored artifact and hidden during printing. Original landscape sizing, margins and page rules are unchanged. Browser/system-font/printer differences can affect physical pagination; this is not byte-for-byte PDF capture. Browser print APIs cannot return PDF bytes to JavaScript, and no manual workaround is required.

Route IDs come from the generated model. That model currently has no trailer/preparation/shipping fields, so those values remain null, never invented or re-entered. Raw workbooks stay local. Unsaved retry content lives only in the current tab and is lost on close/refresh; success is shown only after the server confirms private storage and metadata commit.

See ../SC-OpsReady/docs/routeready-integration.md for the exact capture boundary, APIs, immutable schema, CSP/sandbox, file manifest, isolated test results and rollout prerequisites.

Checks: npm.cmd run typecheck; npm.cmd run lint; npm.cmd test; npm.cmd run build; npm.cmd run build:hub.

Local review from SC-OpsReady: npm.cmd run dev:route-integrated. The disposable localhost:8793 runner creates synthetic accounts and ephemeral D1/R2, applies migrations only to that storage, and prints random review credentials. Sign in as DC_COORDINATOR, upload synthetic P21, Generate and observe automatic Saved without delaying Print. Open history in a separate SITE_MANAGER profile. Do not use ephemeral storage for real reports.

Existing persistent databases are untouched. Live/staging rollout still requires reviewed migrations 0012/0013, private durable R2 binding, explicit Hub mount enablement and browser/physical-print acceptance. Existing dependency advisories remain unchanged; no dependencies were introduced or upgraded by this correction.
