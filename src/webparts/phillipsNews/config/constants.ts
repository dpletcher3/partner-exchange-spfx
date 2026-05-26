// Configuration constants for the Phillips News web part.
//
// HUB_SITE_URL is the default source site for the News Repository list. It
// matches the hub root used by the BrandedHeader customizer (D023: sandbox
// specifics live in code; the production deploy updates the constant). Editors
// can override it per-instance via the Advanced property-pane group, so the
// web part can be repointed without a code change.

export const HUB_SITE_URL =
  'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox';

export const DEFAULT_LIST_TITLE = 'News Repository';

// Sentinel for "no item-type filter" in the item-type dropdown.
export const ANY_ITEM_TYPE = '(any)';

// Flip to true for local dev to render from MockNewsRepositoryService instead
// of hitting SharePoint REST. Ships false so deployed instances read the real
// seeded list.
export const USE_MOCK_SERVICE = false;

// Default card count when maxItems is unset; also the slider's lower/upper bounds.
export const DEFAULT_MAX_ITEMS = 6;
export const MIN_MAX_ITEMS = 1;
export const MAX_MAX_ITEMS = 24;
