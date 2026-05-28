import * as React from 'react';

import styles from './GalleryLayout.module.scss';
import { ITabData, IListRow, IFieldInfo, OverlayPosition } from '../services/models';

// Field TypeAsString values that SharePoint uses for image-bearing columns.
// "Thumbnail" is the modern Image column (Lists 2020+); "URL" with image
// hyperlink subtype falls back via row-value parsing below.
const IMAGE_FIELD_TYPES = new Set<string>(['Thumbnail', 'Image']);

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
  // Image-field detection order:
  //   1. Field metadata — first viewField whose TypeAsString matches a known
  //      image type ("Thumbnail" / "Image"). This is the reliable path.
  //   2. Fall back to row-value sniffing only if #1 finds nothing (e.g.
  //      hyperlink columns serving as ad-hoc image fields).
  const imageField = React.useMemo<string | undefined>(
    () =>
      findImageFieldByMetadata(data.fields, data.viewFields) ??
      findImageFieldByRowValues(data.rows, data.viewFields),
    [data.fields, data.rows, data.viewFields]
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

// Primary detector: walks the view's selected fields and returns the first
// whose TypeAsString identifies it as an image column. Reliable because it
// depends on schema metadata rather than the per-row JSON shape, which has
// surprised us before (1.0.1.1 used row-value sniffing only and failed when
// the Image column's serialized JSON didn't surface `serverRelativeUrl` at
// the expected key).
function findImageFieldByMetadata(
  fields: IFieldInfo[],
  viewFields: string[]
): string | undefined {
  const byName = new Map<string, IFieldInfo>();
  for (const f of fields) {
    byName.set(f.internalName, f);
  }
  for (const field of viewFields) {
    const meta = byName.get(field);
    if (meta && IMAGE_FIELD_TYPES.has(meta.typeAsString)) {
      return field;
    }
  }
  return undefined;
}

// Fallback detector: scans actual row values for anything that looks like an
// image payload. Catches Hyperlink-typed columns being used as ad-hoc image
// fields, or unusual response shapes.
function findImageFieldByRowValues(
  rows: IListRow[],
  viewFields: string[]
): string | undefined {
  for (const field of viewFields) {
    for (const row of rows) {
      if (extractThumbnailUrl(row[field])) {
        return field;
      }
    }
  }
  return undefined;
}

// Pulls a usable URL out of an Image/Thumbnail column value. The column's
// serialized shape varies — RenderListDataAsStream usually returns a JSON
// string like {"type":"thumbnail","serverRelativeUrl":"/sites/.../x.jpg",...},
// but the response can also arrive already parsed, or with the URL under a
// differently-cased key. This walks every plausible shape before giving up.
function extractThumbnailUrl(value: unknown): string | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }

  let obj: { [key: string]: unknown } | undefined;

  if (typeof value === 'string') {
    if (value.charAt(0) !== '{') {
      return undefined;
    }
    try {
      const parsed: unknown = JSON.parse(value);
      if (parsed && typeof parsed === 'object') {
        obj = parsed as { [key: string]: unknown };
      }
    } catch {
      return undefined;
    }
  } else if (typeof value === 'object') {
    obj = value as { [key: string]: unknown };
  }

  if (!obj) {
    return undefined;
  }

  // Try the common URL keys in priority order. serverRelativeUrl is the
  // canonical key for SP Image columns; the others handle edge cases observed
  // across hyperlink columns and verbose-metadata responses.
  const URL_KEYS = ['serverRelativeUrl', 'ServerRelativeUrl', 'fileServerRelativeUrl', 'Url', 'url'];
  for (const key of URL_KEYS) {
    const v = obj[key];
    if (typeof v === 'string' && v) {
      return v;
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
