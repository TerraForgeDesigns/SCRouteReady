interface StatusMessageProps {
  message: string;
}

export function StatusMessage({ message }: StatusMessageProps) {
  return (
    <p className="sr-only" role="status" aria-live="polite" aria-atomic="true">
      {message}
    </p>
  );
}
