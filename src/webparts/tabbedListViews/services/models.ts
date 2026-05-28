// Domain models for the Tabbed List Views web part.

// A single tab as configured in the property pane: a label and the GUID of a
// list view that drives the items shown when the tab is active.
export interface ITabConfig {
  label: string;
  viewId: string;
}

export type Layout = 'gallery' | 'table';

export type OverlayPosition = 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right';

// Minimal list metadata for the property pane's list dropdown.
export interface IListInfo {
  id: string;
  title: string;
}

// Minimal view metadata for the per-tab view dropdown.
export interface IViewInfo {
  id: string;
  title: string;
}

// Field metadata used for the overlay-source dropdown and for table column
// headers (the latter pull display names from this).
export interface IFieldInfo {
  internalName: string;
  displayName: string;
  // SharePoint field type identifiers. typeAsString is the more reliable
  // discriminator for newer field shapes (Image columns report TypeAsString =
  // "Thumbnail" but FieldTypeKind = 0).
  typeKind: number;
  typeAsString: string;
}

// A row from RenderListDataAsStream. Values arrive as strings for most types,
// arrays for Lookup/Person, and JSON strings for Thumbnail/Image columns. Keep
// the values as `unknown` and let extractors at the call site resolve them.
export interface IListRow {
  [field: string]: unknown;
}

// Bundled response for a single tab's data load: the view's field order,
// display-name lookup for those fields (used by table headers), and the rows
// returned by the view's CAML query.
export interface ITabData {
  viewFields: string[];
  fieldDisplayNames: { [internalName: string]: string };
  rows: IListRow[];
}
