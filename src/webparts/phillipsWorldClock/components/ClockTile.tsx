import * as React from 'react';

import styles from './ClockTile.module.scss';
import { formatTime, formatDate } from '../services/timeFormatter';

export interface IClockTileProps {
  title: string;
  timezone: string;
  // The "current" Date — owned by the parent so all tiles share one tick
  // and re-renders are batched.
  now: Date;
}

export const ClockTile: React.FC<IClockTileProps> = ({ title, timezone, now }) => {
  // Compute time + date per render. Both calls are cheap (a single
  // DateTimeFormat construction each) and tied to the parent's tick.
  const { time, period } = formatTime(timezone, now);
  const date = formatDate(timezone, now);

  return (
    <div className={styles.tile}>
      <div className={styles.title} title={title}>{title}</div>
      <div className={styles.timeRow}>
        <span className={styles.time}>{time}</span>
        {period && <span className={styles.period}>{period}</span>}
      </div>
      <div className={styles.date}>{date}</div>
    </div>
  );
};
