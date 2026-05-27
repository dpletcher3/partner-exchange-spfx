import * as React from 'react';
import styles from './PhillipsPersonalizedHero.module.scss';

export interface IDateTimeProps {
  // Inherits the greeting color so the two read as one block.
  color: string;
}

// Current date + time, e.g. "Tuesday, May 26, 2026 · 5:25 PM". Refreshes every
// minute, with the first tick aligned to the next wall-clock minute boundary so
// the displayed minute flips at the same instant as the real one. The interval
// is cleaned up on unmount so it doesn't leak across SharePoint page transitions.
export const DateTime: React.FC<IDateTimeProps> = ({ color }) => {
  const [now, setNow] = React.useState<Date>(() => new Date());

  React.useEffect(() => {
    let intervalId: number | undefined;

    const msToNextMinute = 60000 - (Date.now() % 60000);
    const timeoutId = window.setTimeout(() => {
      setNow(new Date());
      intervalId = window.setInterval(() => setNow(new Date()), 60000);
    }, msToNextMinute);

    return () => {
      window.clearTimeout(timeoutId);
      if (intervalId !== undefined) {
        window.clearInterval(intervalId);
      }
    };
  }, []);

  return (
    <div className={styles.dateTime} style={{ color }}>
      {formatDateTime(now)}
    </div>
  );
};

function formatDateTime(date: Date): string {
  try {
    // `undefined` locale uses the browser's default locale formatting.
    const datePart = date.toLocaleDateString(undefined, {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    const timePart = date.toLocaleTimeString(undefined, {
      hour: 'numeric',
      minute: '2-digit'
    });
    return `${datePart} · ${timePart}`;
  } catch {
    // Locale data unavailable: fall back to the environment default string.
    return date.toString();
  }
}
