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

// Cells render the value as a plain string.
// - Strings, numbers, booleans pass through.
// - Image/Thumbnail JSON strings are dropped to '' (table cells don't render
//   thumbnails in v1; overlays are also off in this layout).
// - Lookup / User columns arrive as objects with a Title (OData $expand), or
//   arrays of similar entries; both shapes are unwrapped to display name(s).
function renderCell(value: unknown): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (typeof value === 'string') {
    if (value.charAt(0) === '{') {
      try {
        const parsed = JSON.parse(value);
        if (parsed && typeof parsed === 'object') {
          const obj = parsed as { [k: string]: unknown };
          // Image / Thumbnail column: drop from table cells.
          if (
            obj.type === 'thumbnail' ||
            typeof obj.fileName === 'string' ||
            typeof obj.serverRelativeUrl === 'string'
          ) {
            return '';
          }
        }
      } catch {
        // Not JSON — render the raw text.
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
      const name = displayNameFromObject(entry);
      if (name) {
        names.push(name);
      }
    }
    return names.join(', ');
  }
  if (typeof value === 'object') {
    return displayNameFromObject(value);
  }
  return '';
}

function displayNameFromObject(value: unknown): string {
  if (!value || typeof value !== 'object') {
    return '';
  }
  const obj = value as { [key: string]: unknown };
  const title = obj.Title ?? obj.title;
  if (typeof title === 'string' && title) {
    return title;
  }
  const lookupValue = obj.lookupValue ?? obj.LookupValue;
  if (typeof lookupValue === 'string' && lookupValue) {
    return lookupValue;
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
