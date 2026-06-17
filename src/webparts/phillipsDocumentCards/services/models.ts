// Domain models for the Document Cards web part.

// Which library columns supply each card field. Internal names, resolved from
// the property-pane mapping (convention with override — defaults FileLeafRef /
// CardDescription / CardIcon / DocSection, per the I16 Stage 1 schema). The data
// service reads by these names, never hardcoded ones. Set once; applies to every
// configured column.
export interface IFieldMapping {
  // Display name. Defaults to the built-in FileLeafRef (file name, extension
  // stripped at map time); remappable to e.g. Title.
  titleField: string;
  // Multi-line text. Default internal name CardDescription (NOT "Description" —
  // see D044 / I16 Stage 1, where the deterministic internal name was chosen).
  descriptionField: string;
  // Image/Thumbnail column. Default CardIcon. Value is a serialized JSON blob;
  // resolved via extractImageColumnUrl.
  iconField: string;
  // Single-value Choice column the per-column $filter runs against. Default
  // DocSection.
  sectionField: string;
}

// A single document card's data, normalized away from the raw SharePoint REST
// shape (Image columns return serialized JSON; the file URL comes from FileRef).
export interface IDocCardItem {
  id: number;
  // Display name — FileLeafRef minus extension by default, or the mapped title
  // column value when remapped.
  title: string;
  // Plain-text description from the mapped Note column. '' when unset.
  description: string;
  // Resolved URL of the icon image from the mapped Image column. undefined when
  // none — the card shows a placeholder. (`undefined`, not `null`, per the rig's
  // @rushstack/no-new-null.)
  iconUrl?: string;
  // Server-relative URL of the file (FileRef), opened on click.
  docUrl: string;
  // The DocSection value this item carries (the column it belongs to).
  section: string;
}

// Per-column configuration, one row per on-screen column (max 4), authored via
// the PropertyFieldCollectionData control. Stage B renders header/icon/color;
// this stage uses filterValue (the DocSection value the column shows) and header.
export interface IDocColumnConfig {
  // The DocSection value this column server-side $filters on (e.g. "Guidelines").
  filterValue: string;
  // Column heading text.
  header: string;
  // Optional free-form header icon name (Office Fabric icon). Rendered in Stage B.
  iconName?: string;
  // Column accent color (hex from CustomCollectionFieldType.color). Used in Stage B.
  color: string;
}
