export function AppHeader() {
  return (
    <header className="app-header">
      <div className="header-inner">
        <div className="logo-lockup">
          <img
            src="/assets/SCRouteReady_logo_cropped.png"
            alt="SCRouteReady"
            loading="eager"
            className="app-logo"
          />
        </div>
        <div className="header-status">
          <p className="eyebrow">Local Browser Processing</p>
          <p className="header-subtitle">Pull Sheet &amp; Pallet Key Generator</p>
          <p className="supporting-line">Secure route preparation without manual spreadsheet cleanup</p>
        </div>
        <a
          className="sparrow-launcher"
          href="https://sparrow-labeler.pages.dev"
          target="_blank"
          rel="noopener noreferrer"
          aria-label="Open Sparrow Labeler in a new tab"
        >
          <span className="sparrow-title">Sparrow Labeler</span>
          <span className="sparrow-copy">Labels &bull; Barcodes &bull; Print Tools</span>
          <span className="sparrow-action">Open Sparrow &rarr;</span>
        </a>
      </div>
    </header>
  );
}
