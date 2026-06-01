// Which list columns supply each field of the highlight block. Internal names,
// resolved from the property-pane mapping (convention with override — defaults
// Title / Video / HighlightInfo). The data layer (Turn 2) reads by these names.
export interface IFieldMapping {
  titleField: string;
  videoField: string;
  infoField: string;
}
