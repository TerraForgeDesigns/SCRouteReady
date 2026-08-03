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
      </div>
    </header>
  );
}
