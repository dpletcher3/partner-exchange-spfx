import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { INewsRepositoryService } from './INewsRepositoryService';
import { INewsItem, INewsFilters } from './models';
import { extractChoices } from './extractors';
import {
  extractBannerImageUrl,
  derivePostOrLink,
  extractOrigin
} from './pipelineExtractors';
import {
  ANY_ITEM_TYPE,
  SITE_PAGES_LIBRARY_TITLE,
  NEWS_CATEGORY_FIELD,
  PROMOTED_STATE_NEWS,
  NEWS_POST_ITEM_TYPE,
  NEWS_LINK_ITEM_TYPE
} from '../config/constants';

// Raw REST shape for a Site Pages news row (minimal-metadata). Declared so the
// mapping avoids `any`. The OData__-prefixed properties are the REST encoding
// of the underscore-prefixed internal columns (_SPSitePageFlags,
// _OriginalSourceUrl) — these carry the News-link signal.
interface IRawPage {
  Id: number;
  Title?: string;
  Description?: string;
  FirstPublishedDate?: string;
  PromotedState?: number;
  FileRef?: string;
  // URL/picture column — { Url, Description } or a bare string. Handled by
  // extractBannerImageUrl.
  BannerImageUrl?: unknown;
  // Single-Choice column (string), but parsed defensively via extractChoices.
  PhillipsNewsCategory?: unknown;
  OData__SPSitePageFlags?: unknown;
  OData__OriginalSourceUrl?: unknown;
}

interface IItemsResponse {
  value: IRawPage[];
}

interface IChoicesResponse {
  Choices?: string[] | { results?: string[] };
}

// Core fields present on every modern Site Pages library. The link-detection
// fields are requested on top of these; if that request fails (e.g. a tenant
// where the OData__ columns are absent), the query retries with the core set
// and every page degrades to a News post — a safe, non-fatal fallback.
const SELECT_CORE =
  'Id,Title,Description,FirstPublishedDate,PromotedState,FileRef,BannerImageUrl,' +
  NEWS_CATEGORY_FIELD;
const SELECT_LINK = 'OData__SPSitePageFlags,OData__OriginalSourceUrl';

// Item type (News post vs News link) is DERIVED per page, not a queryable
// column, so an item-type filter must be applied client-side. When one is
// active we over-fetch up to this ceiling, derive, filter, then slice. The
// category filter, by contrast, is single-Choice and filtered server-side with
// no ceiling (the whole point of the PhillipsNewsCategory column).
const ITEM_TYPE_FETCH_CEILING = 100;

export class NewsPipelineRepositoryService implements INewsRepositoryService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getCategories(siteUrl: string, _listTitle: string): Promise<string[]> {
    // Reads the PhillipsNewsCategory choices from the Site Pages library —
    // same fields/getbytitle(...)/Choices pattern as list mode, different
    // library and field. _listTitle is ignored: pipeline mode always targets
    // the Site Pages library.
    const url = `${trimTrailingSlash(siteUrl)}/_api/web/lists/getbytitle('${encodeListTitle(
      SITE_PAGES_LIBRARY_TITLE
    )}')/fields/getbytitle('${NEWS_CATEGORY_FIELD}')?$select=Choices`;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(
        `Failed to load '${NEWS_CATEGORY_FIELD}' choices: ${response.status} ${response.statusText}`
      );
    }
    const json: IChoicesResponse = await response.json();
    return extractChoices(json.Choices);
  }

  public async getItemTypes(_siteUrl: string, _listTitle: string): Promise<string[]> {
    // The only item types in pipeline mode are the two derived kinds. No REST
    // call — the values are intrinsic to the news-pipeline model.
    return [NEWS_POST_ITEM_TYPE, NEWS_LINK_ITEM_TYPE];
  }

  public async getNewsItems(
    siteUrl: string,
    _listTitle: string,
    filters: INewsFilters,
    maxItems: number
  ): Promise<INewsItem[]> {
    const filterByCategory = !!(filters.categories && filters.categories.length > 0);
    const filterByItemType = !!(filters.itemType && filters.itemType !== ANY_ITEM_TYPE);

    // PromotedState=2 is the news gate. PhillipsNewsCategory is single-Choice,
    // so a category filter is server-side (OR across the selected choices).
    const filterClauses: string[] = [`PromotedState eq ${PROMOTED_STATE_NEWS}`];
    if (filterByCategory) {
      const catClause = filters.categories
        .map((c) => `${NEWS_CATEGORY_FIELD} eq '${escapeODataLiteral(c)}'`)
        .join(' or ');
      filterClauses.push(`(${catClause})`);
    }

    const serverTop = filterByItemType ? ITEM_TYPE_FETCH_CEILING : maxItems;
    const baseQuery = [
      `$orderby=${encodeURIComponent('FirstPublishedDate desc')}`,
      `$top=${serverTop}`,
      `$filter=${encodeURIComponent(filterClauses.join(' and '))}`
    ];

    const itemsBase = `${trimTrailingSlash(siteUrl)}/_api/web/lists/getbytitle('${encodeListTitle(
      SITE_PAGES_LIBRARY_TITLE
    )}')/items`;

    const rows = await this._fetchRows(itemsBase, baseQuery);
    const origin = extractOrigin(siteUrl);

    let items = rows.map((row) => mapPage(row, origin));

    if (filterByItemType) {
      items = items.filter((item) => item.itemType === filters.itemType);
    }

    return items.slice(0, maxItems);
  }

  // Requests rows with the link-detection fields; on failure, warns and retries
  // with the core field set (degraded: every page maps as a News post). Throws
  // only if the core query also fails.
  private async _fetchRows(itemsBase: string, baseQuery: string[]): Promise<IRawPage[]> {
    const fullUrl = `${itemsBase}?${[`$select=${SELECT_CORE},${SELECT_LINK}`, ...baseQuery].join('&')}`;
    const full = await this._spHttpClient.get(fullUrl, SPHttpClient.configurations.v1);
    if (full.ok) {
      const json: IItemsResponse = await full.json();
      return json.value || [];
    }

    console.warn(
      `[PhillipsNews] Pipeline query with link fields failed (${full.status}); retrying with core fields`
    );
    const coreUrl = `${itemsBase}?${[`$select=${SELECT_CORE}`, ...baseQuery].join('&')}`;
    const core: SPHttpClientResponse = await this._spHttpClient.get(
      coreUrl,
      SPHttpClient.configurations.v1
    );
    if (!core.ok) {
      throw new Error(`News pipeline query failed: ${core.status} ${core.statusText}`);
    }
    const json: IItemsResponse = await core.json();
    return json.value || [];
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mapPage(row: IRawPage, origin: string): INewsItem {
  const link = derivePostOrLink({
    flags: row.OData__SPSitePageFlags,
    redirectUrlRaw: row.OData__OriginalSourceUrl,
    fileRef: row.FileRef || '',
    origin
  });

  return {
    id: row.Id,
    title: row.Title || '',
    // Single-Choice column: extractChoices yields [value] or [] when unset.
    categories: extractChoices(row.PhillipsNewsCategory),
    itemType: link.itemType,
    linkUrl: link.linkUrl,
    thumbnailImageUrl: extractBannerImageUrl(row.BannerImageUrl),
    shortDescription: row.Description || '',
    publishedDate: row.FirstPublishedDate || ''
  };
}

// Single quotes inside an OData string literal are escaped by doubling them.
function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}

function encodeListTitle(listTitle: string): string {
  return encodeURIComponent(escapeODataLiteral(listTitle));
}

function trimTrailingSlash(url: string): string {
  return url.replace(/\/+$/, '');
}
