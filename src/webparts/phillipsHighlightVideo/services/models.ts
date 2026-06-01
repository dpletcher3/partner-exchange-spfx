// Which list columns supply each field of the highlight block. Internal names,
// resolved from the property-pane mapping (convention with override — defaults
// Title / Video / HighlightInfo). The data layer (Turn 2) reads by these names.
export interface IFieldMapping {
  titleField: string;
  videoField: string;
  infoField: string;
}

// The resolved featured item, normalized from the raw REST shape (the video
// field is a Hyperlink column → { Url, Description }; info is plain Note text).
export interface IHighlightItem {
  id: number;
  title: string;
  videoUrl: string;
  info: string;
}
