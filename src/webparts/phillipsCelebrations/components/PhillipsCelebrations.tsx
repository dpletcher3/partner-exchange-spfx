import * as React from 'react';
import styles from './PhillipsCelebrations.module.scss';
import { CelebrationsDataService, ICelebrationPerson, ICelebrationsMapping } from '../services/CelebrationsDataService';
import { buildCelebrations, ICelebrationEvent, ICelebrationsResult, WeekStart } from '../services/celebrationsCalendar';

const LOG = '[Celebrations]';

export interface IPhillipsCelebrationsProps {
  service: CelebrationsDataService;
  siteUrl: string;
  listId: string;
  mapping: ICelebrationsMapping;
  weekStart: WeekStart;
  defaultTab: 'birthdays' | 'anniversaries';
  // Visibility of the (disabled placeholder) Send-a-Wish button; default hidden.
  showSendAWish?: boolean;
}

type Tab = 'anniversaries' | 'birthdays';
type Status = 'configure' | 'loading' | 'error' | 'loaded';

const WEEKDAYS_SUN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WEEKDAYS_MON = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

export const PhillipsCelebrations: React.FC<IPhillipsCelebrationsProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('configure');
  const [tab, setTab] = React.useState<Tab>(props.defaultTab);
  const [result, setResult] = React.useState<ICelebrationsResult | undefined>(undefined);
  const [peopleById, setPeopleById] = React.useState<Map<number, ICelebrationPerson>>(new Map());

  const m = props.mapping;
  const depsKey = `${props.listId}|${m.personField}|${m.birthdayField}|${m.hireField}|${props.weekStart}`;

  const load = React.useCallback(() => {
    if (!props.listId) {
      setStatus('configure');
      return;
    }
    setStatus('loading');
    props.service
      .getPeople(props.siteUrl, props.listId, props.mapping)
      .then((people) => {
        const map = new Map<number, ICelebrationPerson>(people.map((p) => [p.id, p]));
        const res = buildCelebrations(
          people.map((p) => ({ id: p.id, name: p.name, birthDate: p.birthDate, hireDate: p.hireDate })),
          new Date(),
          props.weekStart
        );
        res.eventsByDay.forEach((day) =>
          day.forEach((ev) =>
            console.log(
              `${LOG} ${ev.name}: ${ev.type} on ${ev.date.toDateString()}${ev.years ? ` (${ev.years}y)` : ''}`
            )
          )
        );
        setPeopleById(map);
        setResult(res);
        setStatus('loaded');
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} load failed`, err);
        setStatus('error');
      });
  }, [props.service, props.siteUrl, depsKey]);

  React.useEffect(() => {
    load();
  }, [load]);

  React.useEffect(() => {
    setTab(props.defaultTab);
  }, [props.defaultTab]);

  const weekdays = props.weekStart === 'monday' ? WEEKDAYS_MON : WEEKDAYS_SUN;
  const wantType = tab === 'birthdays' ? 'birthday' : 'anniversary';
  const filtered: ICelebrationEvent[][] = result
    ? result.eventsByDay.map((day) => day.filter((e) => e.type === wantType))
    : [];
  const total = filtered.reduce((n, d) => n + d.length, 0);

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
      {status === 'loaded' && result && total === 0 && (
        <div className={styles.message}>
          No {tab === 'birthdays' ? 'birthdays' : 'work anniversaries'} in the current or next week.
        </div>
      )}
      {status === 'loaded' && result && total > 0 && (
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
                const i = week * 7 + col;
                const day = result.window[i];
                const events = filtered[i] || [];
                return (
                  <div key={i} className={styles.dayCell}>
                    <div className={styles.dayNum} aria-hidden="true">
                      {day.getDate()}
                    </div>
                    {events.map((ev) => {
                      const person = peopleById.get(ev.personId);
                      return (
                        <div key={`${ev.personId}-${ev.type}`} className={styles.person}>
                          {person && person.photoUrl ? (
                            <img className={styles.avatar} src={person.photoUrl} alt={ev.name} />
                          ) : (
                            <div className={styles.avatar} aria-hidden="true" />
                          )}
                          <div className={styles.personName}>{ev.name}</div>
                          {ev.type === 'anniversary' && ev.years && (
                            <div className={styles.years}>{ev.years} {ev.years === 1 ? 'Year' : 'Years'}</div>
                          )}
                          {/* Send a Wish — disabled placeholder, hidden by default via the
                              showSendAWish toggle; Turn 3 wiring (handler + Viva Engage
                              deep-link) still pending. Kept intact so the toggle restores it. */}
                          {props.showSendAWish && (
                            <button type="button" className={styles.wishBtn} disabled>
                              Send a Wish
                            </button>
                          )}
                        </div>
                      );
                    })}
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
