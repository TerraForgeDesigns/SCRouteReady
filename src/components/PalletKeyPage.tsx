import type { PalletKeyPage as PalletKeyPageModel } from '../types/pullSheet';

interface PalletKeyPageProps {
  page: PalletKeyPageModel;
}

export function PalletKeyPage({ page }: PalletKeyPageProps) {
  return (
    <article className={`preview-page pallet-key-page new-print-page ${page.stopColorClass}`}>
      <header className="pallet-key-header">
        <div className="print-brand-block">
          <img src="/assets/SCRouteReady_logo_cropped.png" alt="SCRouteReady" className="print-logo" />
          <p>Route from Descartes: {page.routeDisplay}</p>
          <h3>STOP {page.stopDisplay} - PALLET KEY</h3>
        </div>
      </header>
      <p className="pallet-instruction">Write the matching letter on each box and tape this key to the pallet.</p>
      <div className="pallet-key-list">
        {page.entries.map((entry) => (
          <section className="pallet-key-entry" key={entry.designation}>
            <div className={`designation-letter ${entry.keyShadeClass}`}>{entry.designation}</div>
            <div>
              <h4>{entry.shipToName}</h4>
              <p>{entry.shipToCity}{entry.shipToCity && entry.shipToState ? ', ' : ''}{entry.shipToState}</p>
              <p><strong>Pick Tickets:</strong> {entry.pickTicketNumbers.join(', ')}</p>
              {entry.poNumbers.length > 0 ? <p><strong>PO:</strong> {entry.poNumbers.join(', ')}</p> : null}
            </div>
          </section>
        ))}
      </div>
    </article>
  );
}
