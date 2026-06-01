import * as React from 'react';

import styles from './PhillipsWorldClock.module.scss';
import { ClockTile } from './ClockTile';

export interface IClockConfig {
  title: string;
  timezone: string;
}

export interface IPhillipsWorldClockProps {
  sectionTitle: string;
  clocks: IClockConfig[];
  unconfiguredMessage: string;
}

// Aligns the tick to the next wall-clock minute boundary so the display
// updates within ~250ms of when humans expect ("11:04" appears at :04:00,
// not at :04:45). A bare setInterval(60000) starts wherever the component
// mounts and would drift the display by up to 59s behind reality.
function useMinuteAlignedNow(): Date {
  const [now, setNow] = React.useState<Date>(() => new Date());

  React.useEffect(() => {
    const tick = (): void => setNow(new Date());
    const initial = new Date();
    const msUntilNextMinute = 60000 - (initial.getSeconds() * 1000 + initial.getMilliseconds());

    let intervalId: ReturnType<typeof setInterval> | undefined;
    const timeoutId = setTimeout(() => {
      tick();
      intervalId = setInterval(tick, 60000);
    }, msUntilNextMinute);

    return () => {
      clearTimeout(timeoutId);
      if (intervalId !== undefined) {
        clearInterval(intervalId);
      }
    };
  }, []);

  return now;
}

export const PhillipsWorldClock: React.FC<IPhillipsWorldClockProps> = ({
  sectionTitle,
  clocks,
  unconfiguredMessage
}) => {
  const now = useMinuteAlignedNow();

  if (!clocks || clocks.length === 0) {
    return (
      <section className={styles.container}>
        {sectionTitle && (
          <div className={styles.header}>
            <h2 className={styles.title}>{sectionTitle}</h2>
          </div>
        )}
        <div className={styles.unconfigured}>{unconfiguredMessage}</div>
      </section>
    );
  }

  return (
    <section className={styles.container}>
      {sectionTitle && (
        <div className={styles.header}>
          <h2 className={styles.title}>{sectionTitle}</h2>
        </div>
      )}
      <div className={styles.grid}>
        {clocks.map((clock, idx) => (
          <ClockTile
            key={`${clock.title}-${clock.timezone}-${idx}`}
            title={clock.title}
            timezone={clock.timezone}
            now={now}
          />
        ))}
      </div>
    </section>
  );
};
