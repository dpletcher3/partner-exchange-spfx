import * as React from 'react';

import styles from './PhillipsDocumentCards.module.scss';
import { PhillipsDocumentCardsService } from '../services/PhillipsDocumentCardsService';
import { IFieldMapping, IDocCardItem, IDocColumnConfig } from '../services/models';

const LOG = '[DocumentCards]';
const MAX_COLUMNS = 4;

// Generic document fallback glyph, shown in the card icon slot when an item has
// no CardIcon image. Inline SVG to match the repo's existing icon idiom (the
// Media Gallery play badge) — no Fluent React import, no FluentProvider. Fills
// with currentColor so it inherits the column accent set on the icon wrapper.
// Decorative: aria-hidden, the title link carries the accessible name.
const DocumentGlyph: React.FC = () => (
  <svg className={styles.fallbackGlyph} viewBox="0 0 24 24" focusable="false" aria-hidden="true">
    <path
      fill="currentColor"
      d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm0 2.5L17.5 8H14V4.5zM8 10h4v1.5H8V10zm0 3h8v1.5H8V13zm0 3h8v1.5H8V16z"
    />
  </svg>
);

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
        {columns.map((col, idx) => {
          // Per-column accent: set the custom property only when the author
          // picked a color; otherwise the SCSS var() fallback resolves to
          // --phil-red. Both --phil-dc-cols and this are custom properties, so
          // the cast matches the gridStyle pattern above.
          const columnStyle = (
            col.config.color ? { ['--phil-dc-accent']: col.config.color } : {}
          ) as React.CSSProperties;
          return (
            <div key={idx} className={styles.column} style={columnStyle}>
              <h3 className={styles.columnHeader}>
                {col.config.iconName && (
                  <i className={`ms-Icon ms-Icon--${col.config.iconName} ${styles.columnHeaderIcon}`} aria-hidden="true" />
                )}
                <span>{col.config.header || col.config.filterValue}</span>
              </h3>
              {col.error && <p className={styles.message}>Error: {col.error}</p>}
              {!col.error && col.items.length === 0 && (
                <p className={styles.message}>No documents in &quot;{col.config.filterValue}&quot;.</p>
              )}
              <ul className={styles.list}>
                {col.items.map((item) => (
                  <li key={item.id} className={styles.item}>
                    <a className={styles.card} href={item.docUrl}>
                      <span className={styles.iconWrap}>
                        {item.iconUrl ? (
                          <img className={styles.icon} src={item.iconUrl} alt="" />
                        ) : (
                          <DocumentGlyph />
                        )}
                      </span>
                      <span className={styles.body}>
                        <span className={styles.title}>{item.title}</span>
                        {item.description && <span className={styles.description}>{item.description}</span>}
                      </span>
                    </a>
                  </li>
                ))}
              </ul>
            </div>
          );
        })}
      </div>
    </div>
  );
};
