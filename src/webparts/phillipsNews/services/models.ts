// Domain models for the Phillips News web part.

// A single news card's data, normalized away from the raw SharePoint REST shape
// (Image columns return serialized JSON, URL columns return { Url, Description }).
export interface INewsItem {
  id: number;
  title: string;
  // Category is a MultiChoice column in the News Repository list, so an item can
  // carry more than one category. Stored as an array; the card renders the first.
  categories: string[];
  itemType: string;
  linkUrl: string;
  // Server-relative URL of the item's thumbnail, resolved from the
  // ThumbnailImage field + AttachmentFiles. Absent when the item has no usable
  // thumbnail — the card renders the red fallback. Uses `undefined` rather than
  // `null` per the rig's @rushstack/no-new-null rule (semantics identical).
  thumbnailImageUrl?: string;
  shortDescription: string;
  publishedDate: string; // ISO 8601 string as returned by SharePoint
}

// Filters applied to a news query. An empty `categories` array means "no
// category filter"; an `itemType` of undefined or the ANY_ITEM_TYPE sentinel
// means "no item-type filter".
export interface INewsFilters {
  categories: string[];
  itemType?: string;
}
