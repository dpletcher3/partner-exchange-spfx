import * as React from 'react';

import styles from './GalleryLayout.module.scss';
import { ITabData, IListRow, OverlayPosition } from '../services/models';

export interface IGalleryLayoutProps {
  data: ITabData;
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
}

const OVERLAY_POSITION_CLASS: { [key in OverlayPosition]: string } = {
  'top-left': styles.overlayTopLeft,
  'top-right': styles.overlayTopRight,
  'bottom-left': styles.overlayBottomLeft,
  'bottom-right': styles.overlayBottomRight
};

interface ICardProps {
  row: IListRow;
  imageField: string | undefined;
  viewFields: string[];
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
}

const Card: React.FC<ICardProps> = ({
  row,
  imageField,
  viewFields,
  showOverlay,
  overlaySourceField,
  overlayLabelTemplate,
  overlayPosition
}) => {
  const title = getStringValue(row.Title);
  const imageUrl = imageField ? extractThumbnailUrl(row[imageField]) : undefined;
  const secondary = getSecondaryLine(row, viewFields, imageField);

  // Overlay only renders when configured AND the source field has a non-empty
  // value for THIS item. Empty values skip the badge silently.
  const overlayValue = showOverlay && overlaySourceField
    ? coerceToString(row[overlaySourceField])
    : '';
  const overlayLabel = overlayValue
    ? (overlayLabelTemplate || '{value}').replace(/\{value\}/g, overlayValue)
    : '';

  const positionClass = OVERLAY_POSITION_CLASS[overlayPosition] || styles.overlayBottomLeft;

  return (
    <div className={styles.card}>
      <div className={styles.thumb}>
        {imageUrl ? (
          <img className={styles.thumbImg} src={imageUrl} alt={title} />
        ) : (
          <div className={styles.fallback} aria-hidden="true" />
        )}
        {overlayLabel && (
          <div className={`${styles.overlay} ${positionClass}`}>{overlayLabel}</div>
        )}
      </div>
      <div className={styles.body}>
        {title && <div className={styles.cardTitle}>{title}</div>}
        {secondary && <div className={styles.secondary}>{secondary}</div>}
      </div>
    </div>
  );
};

export const GalleryLayout: React.FC<IGalleryLayoutProps> = ({
  data,
  showOverlay,
  overlaySourceField,
  overlayLabelTemplate,
  overlayPosition
}) => {
  // Detect the image field once per data load — the first field whose value
  // parses as a Thumbnail JSON. Done at component scope so all cards in the
  // grid share the same image field choice (consistent layout per tab).
  const imageField = React.useMemo<string | undefined>(
    () => findFirstThumbnailField(data.rows, data.viewFields),
    [data.rows, data.viewFields]
  );

  return (
    <div className={styles.grid}>
      {data.rows.map((row, idx) => (
        <Card
          key={getRowKey(row, idx)}
          row={row}
          imageField={imageField}
          viewFields={data.viewFields}
          showOverlay={showOverlay}
          overlaySourceField={overlaySourceField}
          overlayLabelTemplate={overlayLabelTemplate}
          overlayPosition={overlayPosition}
        />
      ))}
    </div>
  );
};

// ---------------------------------------------------------------------------
// Value extractors
// ---------------------------------------------------------------------------

// RenderListDataAsStream serializes Image / Thumbnail columns as a JSON string.
// Returns the field's InternalName, or undefined if no row carries a parseable
// thumbnail in any of the view's fields.
function findFirstThumbnailField(rows: IListRow[], viewFields: string[]): string | undefined {
  for (const field of viewFields) {
    for (const row of rows) {
      if (extractThumbnailUrl(row[field])) {
        return field;
      }
    }
  }
  return undefined;
}

// Pulls serverRelativeUrl out of an Image/Thumbnail column value, which arrives
// as a JSON string like {"type":"thumbnail","fileName":"x","serverRelativeUrl":"/sites/.../x.jpg",...}.
function extractThumbnailUrl(value: unknown): string | undefined {
  if (typeof value !== 'string' || !value) {
    return undefined;
  }
  // Image columns always start with a JSON object opening brace; cheap early
  // exit for plain text values.
  if (value.charAt(0) !== '{') {
    return undefined;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return undefined;
  }
  if (parsed && typeof parsed === 'object') {
    const obj = parsed as { [key: string]: unknown };
    const url = obj.serverRelativeUrl;
    if (typeof url === 'string' && url) {
      return url;
    }
  }
  return undefined;
}

// Picks the first non-title, non-image field in the view's order and returns
// its value as a display string. Returns '' when no usable value exists.
function getSecondaryLine(
  row: IListRow,
  viewFields: string[],
  imageField: string | undefined
): string {
  for (const field of viewFields) {
    if (field === 'Title' || field === imageField || field === 'ID' || field === 'LinkTitle') {
      continue;
    }
    const value = coerceToString(row[field]);
    if (value) {
      return value;
    }
  }
  return '';
}

// Lookup/Person columns arrive as arrays of { lookupId, lookupValue } pairs;
// text values arrive as strings. coerceToString handles both plus the empty
// cases and returns '' for anything it can't resolve.
function coerceToString(value: unknown): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (typeof value === 'string') {
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

function getStringValue(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

// Stable per-row key for React reconciliation. SharePoint always returns ID;
// the array index is the last-resort fallback.
function getRowKey(row: IListRow, idx: number): string {
  const id = row.ID ?? row.Id;
  if (typeof id === 'string' || typeof id === 'number') {
    return String(id);
  }
  return `row-${idx}`;
}
