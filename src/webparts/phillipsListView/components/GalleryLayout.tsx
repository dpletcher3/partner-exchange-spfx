import * as React from 'react';

import styles from './GalleryLayout.module.scss';
import { ITabData, IListRow, IFieldInfo, OverlayPosition } from '../services/models';

// Field TypeAsString values that SharePoint uses for image-bearing columns.
const IMAGE_FIELD_TYPES = new Set<string>(['Thumbnail', 'Image']);

// Field TypeAsString values for Person/User columns. We treat these as
// image-bearing because their value resolves to a profile photo via
// /_layouts/15/userphoto.aspx.
const PERSON_FIELD_TYPES = new Set<string>(['User', 'UserMulti']);

// Person-field internal-name priority. When a list has multiple Person fields
// the one whose internal name contains the earliest substring here wins,
// before falling through to "any other non-administrative Person field". This
// is how Partner Profiles' LinkedUser beats Author/Editor without the editor
// having to configure the field explicitly.
const PERSON_PRIORITY_SUBSTRINGS = ['LinkedUser', 'User', 'Person', 'Profile'];

// Administrative Person fields present on every SharePoint list. Skipped by
// auto-detect — they almost never represent "the person this row is about"
// semantically, and picking them would silently render an editor's photo
// instead of the subject's.
const PERSON_DEPRIORITIZED_NAMES = new Set<string>([
  'Author',
  'Editor',
  'CreatedBy',
  'ModifiedBy'
]);

// Descriptor for the auto-detected image field. `kind` tells the per-row
// extractor which path to use: 'image' = Image/Thumbnail column with file in
// AttachmentFiles; 'person' = Person/User column with email → userphoto.aspx;
// 'unknown' = field was found via row-value sniffing (Hyperlink etc.) and uses
// the generic URL extractor.
type ImageFieldKind = 'image' | 'person' | 'unknown';
interface IImageFieldDescriptor {
  name: string;
  kind: ImageFieldKind;
}

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
  imageField: IImageFieldDescriptor | undefined;
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
  const imageUrl = imageField
    ? resolveImageUrl(imageField, row)
    : undefined;
  const secondary = getSecondaryLine(row, viewFields, imageField?.name);

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
  //   1. Field metadata — explicit Image / Thumbnail columns first, then
  //      priority-ordered Person columns. This is the reliable path.
  //   2. Fall back to row-value sniffing only if #1 finds nothing (e.g.
  //      Hyperlink columns serving as ad-hoc image fields).
  const imageField = React.useMemo<IImageFieldDescriptor | undefined>(
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
// Image-field detection
// ---------------------------------------------------------------------------

// Primary detector: walks the view's selected fields and returns the first
// match by metadata. Image / Thumbnail columns beat Person columns; among
// Person columns, internal-name priority decides the winner.
function findImageFieldByMetadata(
  fields: IFieldInfo[],
  viewFields: string[]
): IImageFieldDescriptor | undefined {
  const byName = new Map<string, IFieldInfo>();
  for (const f of fields) {
    byName.set(f.internalName, f);
  }

  // 1. Explicit Image / Thumbnail columns first — if a list has one, it's
  //    almost always the intended "row photo".
  for (const field of viewFields) {
    const meta = byName.get(field);
    if (meta && IMAGE_FIELD_TYPES.has(meta.typeAsString)) {
      return { name: field, kind: 'image' };
    }
  }

  // 2. Priority-ordered Person columns. The first priority substring that
  //    matches any view-field's internal name wins. Substring match is
  //    case-insensitive so 'linkedUser' / 'LinkedUser' / 'LINKEDUSER' all work.
  const personFields = viewFields.filter((f) => {
    const meta = byName.get(f);
    return !!meta && PERSON_FIELD_TYPES.has(meta.typeAsString);
  });
  for (const substring of PERSON_PRIORITY_SUBSTRINGS) {
    const subLower = substring.toLowerCase();
    for (const field of personFields) {
      if (field.toLowerCase().indexOf(subLower) !== -1) {
        return { name: field, kind: 'person' };
      }
    }
  }

  // 3. Any other Person field whose internal name isn't on the deprioritized
  //    list (Author / Editor / CreatedBy / ModifiedBy).
  for (const field of personFields) {
    if (!PERSON_DEPRIORITIZED_NAMES.has(field)) {
      return { name: field, kind: 'person' };
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
): IImageFieldDescriptor | undefined {
  for (const field of viewFields) {
    for (const row of rows) {
      if (looksLikeImagePayload(row[field])) {
        return { name: field, kind: 'unknown' };
      }
    }
  }
  return undefined;
}

// Shape check used by the row-value fallback detector. Cheap parse without
// the AttachmentFiles lookup that the full extractor performs.
function looksLikeImagePayload(value: unknown): boolean {
  if (typeof value !== 'string' || value.charAt(0) !== '{') {
    return false;
  }
  try {
    const parsed: unknown = JSON.parse(value);
    if (parsed && typeof parsed === 'object') {
      const obj = parsed as { [k: string]: unknown };
      return typeof obj.fileName === 'string' || typeof obj.serverRelativeUrl === 'string';
    }
  } catch {
    // Not JSON — not an image payload.
  }
  return false;
}

// ---------------------------------------------------------------------------
// Per-row URL resolution
// ---------------------------------------------------------------------------

// Dispatches to the right extractor based on the descriptor's kind. Returns
// undefined if the row's value is empty or can't be resolved — the caller
// falls back to the gray placeholder.
function resolveImageUrl(
  descriptor: IImageFieldDescriptor,
  row: IListRow
): string | undefined {
  const value = row[descriptor.name];
  if (descriptor.kind === 'person') {
    const email = extractPersonEmail(value);
    if (!email) {
      return undefined;
    }
    // userphoto.aspx is host-relative; works across site collections without
    // siteUrl prefixing. size=L (~96px) is the right footprint for our 280px
    // card thumb at 4:3 aspect — smaller sizes pixelate noticeably.
    return `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=L`;
  }
  return extractThumbnailUrl(value, row.AttachmentFiles);
}

// Pulls the email out of a Person/User column value. RenderListDataAsStream
// returns Person fields as either an array of entry-objects or a JSON-encoded
// string of the same; the entry has `email` (lowercase) in current SP REST
// responses, with `Email` / `EMail` as fallbacks for older response shapes.
function extractPersonEmail(value: unknown): string | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }

  let candidates: unknown[] = [];
  if (typeof value === 'string') {
    const first = value.charAt(0);
    if (first !== '[' && first !== '{') {
      return undefined;
    }
    try {
      const parsed: unknown = JSON.parse(value);
      candidates = Array.isArray(parsed) ? parsed : [parsed];
    } catch {
      return undefined;
    }
  } else if (Array.isArray(value)) {
    candidates = value;
  } else if (typeof value === 'object') {
    candidates = [value];
  }

  for (const candidate of candidates) {
    if (!candidate || typeof candidate !== 'object') {
      continue;
    }
    const obj = candidate as { [k: string]: unknown };
    const email = obj.email ?? obj.Email ?? obj.EMail;
    if (typeof email === 'string' && email.indexOf('@') !== -1) {
      return email;
    }
  }
  return undefined;
}

// Pulls a usable URL out of an Image / Thumbnail column value. SharePoint
// Image columns serialize as a JSON string containing only a fileName plus
// the original display name — no URL — and the real file lives in the item's
// AttachmentFiles, where the filename matches the JSON's `fileName`.
function extractThumbnailUrl(value: unknown, attachmentFiles: unknown): string | undefined {
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

  // If the URL is inlined (modern hyperlink subtype, RenderListDataAsStream
  // responses, etc.), take it directly.
  const URL_KEYS = ['serverRelativeUrl', 'ServerRelativeUrl', 'fileServerRelativeUrl', 'Url', 'url'];
  for (const key of URL_KEYS) {
    const v = obj[key];
    if (typeof v === 'string' && v) {
      return v;
    }
  }

  // Otherwise resolve via attachment lookup — match the JSON's fileName to an
  // entry in AttachmentFiles and use that attachment's ServerRelativeUrl.
  const fileName = obj.fileName;
  if (typeof fileName === 'string' && fileName) {
    const url = findAttachmentUrl(attachmentFiles, fileName);
    if (url) {
      return url;
    }
  }

  return undefined;
}

// AttachmentFiles arrives as a plain array under minimal-metadata responses
// and as { results: [...] } under verbose. Returns the ServerRelativeUrl of
// the attachment whose FileName matches `fileName`, or undefined if no match.
function findAttachmentUrl(attachmentFiles: unknown, fileName: string): string | undefined {
  let list: unknown[] = [];
  if (Array.isArray(attachmentFiles)) {
    list = attachmentFiles;
  } else if (attachmentFiles && typeof attachmentFiles === 'object') {
    const results = (attachmentFiles as { results?: unknown }).results;
    if (Array.isArray(results)) {
      list = results;
    }
  }
  for (const entry of list) {
    if (entry && typeof entry === 'object') {
      const att = entry as { FileName?: unknown; ServerRelativeUrl?: unknown };
      if (att.FileName === fileName && typeof att.ServerRelativeUrl === 'string') {
        return att.ServerRelativeUrl;
      }
    }
  }
  return undefined;
}

// Picks the first non-title, non-image field in the view's order and returns
// its value as a display string. Returns '' when no usable value exists.
function getSecondaryLine(
  row: IListRow,
  viewFields: string[],
  imageFieldName: string | undefined
): string {
  for (const field of viewFields) {
    if (field === 'Title' || field === imageFieldName || field === 'ID' || field === 'LinkTitle') {
      continue;
    }
    const value = coerceToString(row[field]);
    if (value) {
      return value;
    }
  }
  return '';
}

// Coerces any field value to a display string.
// - Strings, numbers, booleans pass through.
// - Lookup / User columns arrive as objects with a Title (OData $expand) or
//   arrays of similar entries; both shapes are unwrapped to display name(s).
// - Returns '' for empty / null / unrecognized shapes.
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
  // OData $expand=Lookup&$select=Lookup/Title shape:
  const title = obj.Title ?? obj.title;
  if (typeof title === 'string' && title) {
    return title;
  }
  // RenderListDataAsStream lookup shape (kept for safety; no longer the
  // primary path after 1.0.1.3's switch to GetItems).
  const lookupValue = obj.lookupValue ?? obj.LookupValue;
  if (typeof lookupValue === 'string' && lookupValue) {
    return lookupValue;
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
