import * as React from 'react';

import styles from './PhillipsDocumentCards.module.scss';
import { PhillipsDocumentCardsService } from '../services/PhillipsDocumentCardsService';
import { IFieldMapping, IDocCardItem, IDocColumnConfig } from '../services/models';

const LOG = '[DocumentCards]';
const MAX_COLUMNS = 4;

export interface IPhillipsDocumentCardsProps {
  service: PhillipsDocumentCardsService;
  siteUrl: string;
  listId: string;
  mapping: IFieldMapping;
  columns: IDocColumnConfig[];
}

// One on-screen column's fetched state.
interface IColumnState {
  config: IDocColumnConfig;
  items: IDocCardItem[];
  error?: string;
}

// Stage A render: deliberately UNSTYLED. Proves real per-column config drives a
// real server-side $filter against the real library — no cards, colors, icons-
// as-headers, or search yet (those land in Stage B). Each configured column
// shows its header and a plain list of its documents (title link + description +
// an icon placeholder).
export const PhillipsDocumentCards: React.FC<IPhillipsDocumentCardsProps> = (props) => {
  const [columns, setColumns] = React.useState<IColumnState[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);

  const m = props.mapping;
  // Re-fetch when the library, any mapped field, or the column set changes.
  // (Keyed on a primitive so re-renders that rebuild the mapping/columns
  // objects don't trigger spurious fetches — same approach as the Media Gallery.)
  const cfgKey =
    `${props.listId}|${m.titleField}|${m.descriptionField}|${m.iconField}|${m.sectionField}|` +
    props.columns.map((c) => `${c.filterValue}~${c.header}`).join(',');

  const load = React.useCallback(() => {
    if (!props.listId || props.columns.length === 0) {
      setColumns([]);
      return;
    }
    setLoading(true);
    Promise.all(
      props.columns.map((cfg) =>
        props.service
          .getItemsForColumn(props.siteUrl, props.listId, props.mapping, cfg.filterValue)
          .then((items): IColumnState => ({ config: cfg, items }))
          .catch((err: unknown): IColumnState => {
            const msg = err instanceof Error ? err.message : String(err);
            console.warn(`${LOG} column "${cfg.filterValue}" fetch failed`, err);
            return { config: cfg, items: [], error: msg };
          })
      )
    )
      .then((results) => {
        setColumns(results);
        setLoading(false);
      })
      .catch(() => {
        setLoading(false);
      });
  }, [props.service, props.siteUrl, cfgKey]);

  React.useEffect(() => {
    load();
  }, [load]);

  if (!props.listId) {
    return (
      <div className={styles.documentCards}>
        <p className={styles.message}>Select a document library in the property pane.</p>
      </div>
    );
  }
  if (props.columns.length === 0) {
    return (
      <div className={styles.documentCards}>
        <p className={styles.message}>Add one or more columns in the property pane (max {MAX_COLUMNS}).</p>
      </div>
    );
  }

  const colCount = Math.min(props.columns.length, MAX_COLUMNS);
  const gridStyle = { ['--phil-dc-cols']: String(colCount) } as React.CSSProperties;

  return (
    <div className={styles.documentCards}>
      {loading && <p className={styles.message}>Loading…</p>}
      <div className={styles.grid} style={gridStyle}>
        {columns.map((col, idx) => (
          <div key={idx} className={styles.column}>
            <h3 className={styles.columnHeader}>{col.config.header || col.config.filterValue}</h3>
            {col.error && <p className={styles.message}>Error: {col.error}</p>}
            {!col.error && col.items.length === 0 && (
              <p className={styles.message}>No documents in &quot;{col.config.filterValue}&quot;.</p>
            )}
            <ul className={styles.list}>
              {col.items.map((item) => (
                <li key={item.id} className={styles.item}>
                  {item.iconUrl ? (
                    <img className={styles.icon} src={item.iconUrl} alt="" />
                  ) : (
                    <span className={styles.iconPlaceholder}>[icon]</span>
                  )}
                  <a className={styles.title} href={item.docUrl}>
                    {item.title}
                  </a>
                  <div className={styles.description}>{item.description}</div>
                </li>
              ))}
            </ul>
          </div>
        ))}
      </div>
    </div>
  );
};
