// Domain models for the Media Card Gallery.

// Which list columns supply each card field. Internal names, resolved from the
// property-pane mapping (convention with override — defaults Title/Video/Image0,
// main image unset). The data service reads by these names, never hardcoded ones.
export interface IFieldMapping {
  titleField: string;
  videoField: string;
  labelImageField: string;
  // '' when unset — the optional main-image override is not mapped on the
  // 15 Practices instance, so the main visual auto-derives from the video.
  mainImageField: string;
}

// A single card's data, normalized away from the raw SharePoint REST shape
// (Image columns return serialized JSON; URL columns return { Url, Description }).
export interface IMediaCardItem {
  id: number;
  title: string;
  // Video link (YouTube/Vimeo/other). '' when the item has no video — the card
  // is then non-clickable and shows a placeholder main visual.
  videoUrl: string;
  // Resolved server-relative (or absolute) URL of the label graphic from the
  // mapped Image column. undefined when none — the card falls back to styled
  // title text. (`undefined`, not `null`, per the rig's @rushstack/no-new-null.)
  labelImageUrl?: string;
  // Resolved URL of the mapped main-image override column, if mapped + populated.
  // undefined → main visual derives from the video (§3 precedence).
  mainImageOverrideUrl?: string;
}
