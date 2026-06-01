import * as React from 'react';
import styles from './PhillipsCelebrations.module.scss';

const LOG = '[Celebrations]';

export interface IPhillipsCelebrationsProps {
  hasList: boolean;
  weekStart: 'sunday' | 'monday';
  defaultTab: 'birthdays' | 'anniversaries';
}

type Tab = 'anniversaries' | 'birthdays';
type Status = 'configure' | 'loading' | 'error' | 'loaded';

const WEEKDAYS_SUN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WEEKDAYS_MON = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

// Turn 1: scatter a few placeholder people across the shell so the grid + avatar
// circles read as a calendar; most days empty. Turn 2 replaces with real data.
const PLACEHOLDER_CELLS = new Set(['0-1', '0-4', '1-2', '1-5']);

export const PhillipsCelebrations: React.FC<IPhillipsCelebrationsProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');
  const [tab, setTab] = React.useState<Tab>(props.defaultTab);

  // Turn 1: a list selection flips configure → loaded (placeholder shell). Turn 2
  // replaces this with the real Partner Profiles read + date logic.
  React.useEffect(() => {
    const next: Status = props.hasList ? 'loaded' : 'configure';
    setStatus(next);
    console.log(`${LOG} status=${next}`);
  }, [props.hasList]);

  React.useEffect(() => {
    setTab(props.defaultTab);
  }, [props.defaultTab]);

  const weekdays = props.weekStart === 'monday' ? WEEKDAYS_MON : WEEKDAYS_SUN;

  return (
    <section className={styles.celebrations}>
      <div className={styles.tabs} role="tablist" aria-label="Celebrations">
        <button
          type="button"
          role="tab"
          aria-selected={tab === 'anniversaries'}
          className={`${styles.tab} ${tab === 'anniversaries' ? styles.tabActive : ''}`}
          onClick={() => setTab('anniversaries')}
        >
          Anniversaries
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={tab === 'birthdays'}
          className={`${styles.tab} ${tab === 'birthdays' ? styles.tabActive : ''}`}
          onClick={() => setTab('birthdays')}
        >
          Birthdays
        </button>
      </div>

      {status === 'configure' && (
        <div className={styles.message}>
          Select a list and map its columns in the property pane to show celebrations.
        </div>
      )}
      {status === 'loading' && <div className={styles.message}>Loading…</div>}
      {status === 'error' && (
        <div className={styles.message} role="alert">
          Couldn’t load celebrations.
        </div>
      )}
      {status === 'loaded' && (
        <div className={styles.calendar} aria-label={`${tab}: current week and next week`}>
          <div className={styles.weekHeader}>
            {weekdays.map((d) => (
              <div key={d} className={styles.weekdayCol}>
                {d}
              </div>
            ))}
          </div>
          {[0, 1].map((week) => (
            <div key={week} className={styles.weekRow}>
              {weekdays.map((_, col) => {
                const key = `${week}-${col}`;
                return (
                  <div key={key} className={styles.dayCell}>
                    <div className={styles.dayNum} aria-hidden="true">
                      —
                    </div>
                    {PLACEHOLDER_CELLS.has(key) && (
                      <div className={styles.person}>
                        <div className={`${styles.avatar} ${styles.placeholder}`} aria-hidden="true" />
                        <div className={`${styles.personLine} ${styles.placeholder}`} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      )}
    </section>
  );
};
