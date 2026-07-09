// Configuration constants for the Phillips News web part.
//
// HUB_SITE_URL is the default source site for the News Repository list. It
// matches the hub root used by the BrandedHeader customizer (D023: sandbox
// specifics live in code; the production deploy updates the constant). Editors
// can override it per-instance via the Advanced property-pane group, so the
// web part can be repointed without a code change.

export const HUB_SITE_URL =
  'https://phillipscorp.sharepoint.com/sites/PartnerExchange';

export const DEFAULT_LIST_TITLE = 'News Repository';

// Sentinel for "no item-type filter" in the item-type dropdown.
export const ANY_ITEM_TYPE = '(any)';

// --- Data source toggle (I14 / D039) ---------------------------------------
//
// 'list'     — read the News Repository custom list (original behavior).
// 'pipeline' — read SharePoint news pages: Site Pages where PromotedState=2,
//              covering both News posts and News links.
//
// DEFAULT is 'list' so existing deployed instances are untouched on upgrade —
// the toggle only changes behavior when an editor explicitly flips it.
export type DataSource = 'list' | 'pipeline';
export const DEFAULT_DATA_SOURCE: DataSource = 'list';

// The SharePoint pages library title. The pipeline service always reads this
// library; the per-instance listTitle property is meaningless in pipeline mode
// (and hidden in the pane).
export const SITE_PAGES_LIBRARY_TITLE = 'Site Pages';

// Single-Choice column provisioned on the Site Pages libraries of the hub and
// Our Culture (I14 scope item 1). Server-side filterable, unlike the list's
// MultiChoice Category.
export const NEWS_CATEGORY_FIELD = 'PhillipsNewsCategory';

// PromotedState value that marks a page as news (both News posts and News
// links). 0 = ordinary page, 1 = promote-in-progress, 2 = published news.
export const PROMOTED_STATE_NEWS = 2;

// Derived item-type labels in pipeline mode. These are not stored on a column —
// they're derived per page from the news-link marker (see pipelineExtractors).
export const NEWS_POST_ITEM_TYPE = 'News post';
export const NEWS_LINK_ITEM_TYPE = 'News link';

// Flip to true for local dev to render from MockNewsRepositoryService instead
// of hitting SharePoint REST. Ships false so deployed instances read the real
// seeded list.
export const USE_MOCK_SERVICE = false;

// Default card count when maxItems is unset; also the slider's lower/upper bounds.
export const DEFAULT_MAX_ITEMS = 6;
export const MIN_MAX_ITEMS = 1;
export const MAX_MAX_ITEMS = 24;
