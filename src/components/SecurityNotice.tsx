export function SecurityNotice() {
  return (
    <aside className="security-notice" aria-labelledby="security-title">
      <div className="notice-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24">
          <path d="M12 3 5 6v5c0 4.5 2.8 8.5 7 10 4.2-1.5 7-5.5 7-10V6l-7-3Z" />
          <path d="M9.8 12.2 11.3 14l3.2-4" />
        </svg>
      </div>
      <div>
        <h2 id="security-title">Your route file is processed only in this browser.</h2>
        <p>Nothing is transmitted or intentionally stored. Finishing the task or closing the page clears the active working session.</p>
      </div>
    </aside>
  );
}
