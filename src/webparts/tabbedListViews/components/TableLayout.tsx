import * as React from 'react';

import styles from './TableLayout.module.scss';
import { ITabData, IListRow } from '../services/models';

export interface ITableLayoutProps {
  data: ITabData;
}

export const TableLayout: React.FC<ITableLayoutProps> = ({ data }) => {
  // ID is in every RenderListDataAsStream response but it's never a meaningful
  // table column; drop it from the rendered columns.
  const columns = React.useMemo(
    () => data.viewFields.filter((f) => f !== 'ID'),
    [data.viewFields]
  );

  return (
    <div className={styles.tableWrapper}>
      <table className={styles.table}>
        <thead>
          <tr>
            {columns.map((field) => (
              <th key={field} scope="col" className={styles.headerCell}>
                {data.fieldDisplayNames[field] || field}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.rows.map((row, idx) => (
            <tr key={getRowKey(row, idx)} className={styles.row}>
              {columns.map((field) => (
                <td key={field} className={styles.cell}>
                  {renderCell(row[field])}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// ---------------------------------------------------------------------------
// Cell rendering
// ---------------------------------------------------------------------------

// Cells render the value as a plain string. Image/Thumbnail JSON strings fall
// back to '' since rendering thumbnails inside table cells isn't part of the
// spec for v1 (overlays are also off in the table layout).
function renderCell(value: unknown): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (typeof value === 'string') {
    // Image/Thumbnail JSON detected and dropped.
    if (value.charAt(0) === '{') {
      try {
        const parsed = JSON.parse(value);
        if (
          parsed &&
          typeof parsed === 'object' &&
          (parsed as { type?: unknown }).type === 'thumbnail'
        ) {
          return '';
        }
      } catch {
        // Not a JSON value — fall through and render as plain text.
      }
    }
    return value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (Array.isArray(value)) {
    const names: string[] = [];
    for (const entry of value) {
      if (entry && typeof entry === 'object') {
        const obj = entry as { [key: string]: unknown };
        const lookupValue = obj.lookupValue ?? obj.LookupValue ?? obj.title ?? obj.Title;
        if (typeof lookupValue === 'string' && lookupValue) {
          names.push(lookupValue);
        }
      }
    }
    return names.join(', ');
  }
  return '';
}

function getRowKey(row: IListRow, idx: number): string {
  const id = row.ID ?? row.Id;
  if (typeof id === 'string' || typeof id === 'number') {
    return String(id);
  }
  return `row-${idx}`;
}
