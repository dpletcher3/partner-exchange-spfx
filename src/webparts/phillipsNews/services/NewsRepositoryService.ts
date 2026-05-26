import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { INewsRepositoryService } from './INewsRepositoryService';
import { INewsItem, INewsFilters, INewsThumbnail } from './models';
import { ANY_ITEM_TYPE } from '../config/constants';

// Raw REST shapes (minimal-metadata). Declared so the mapping code avoids `any`.
interface IRawUrlField {
  Url?: string;
  Description?: string;
}

interface IRawListItem {
  Id: number;
  Title?: string;
  // MultiChoice column: minimal-metadata returns a plain array, verbose returns
  // { results: [...] }; a single string is handled defensively too.
  Category?: string[] | { results?: string[] } | string;
  ItemType?: string;
  // SharePoint may return these as null at runtime; the mapping guards handle
  // falsy values. Typed optional (not `| null`) per the rig's no-new-null rule.
  // URL fields usually arrive as { Url, Description }, but can surface as a bare
  // string depending on the column/metadata — extractUrl handles both.
  LinkUrl?: IRawUrlField | string;
  ThumbnailImage?: string;
  ShortDescription?: string;
  PublishedDate?: string;
}

interface IItemsResponse {
  value: IRawListItem[];
}

interface IChoicesResponse {
  // SharePoint REST (OData v3) returns a field's collection-valued Choices as a
  // plain array under nometadata, but wrapped as { results: [...] } under the
  // minimalmetadata level SPHttpClient.configurations.v1 negotiates.
  Choices?: string[] | { results?: string[] };
}

interface IRawThumbnail {
  serverRelativeUrl?: string;
  alt?: string;
}

const SELECT_FIELDS =
  'Id,Title,Category,ItemType,LinkUrl,ThumbnailImage,ShortDescription,PublishedDate';

// When a category filter is active we must over-fetch and filter client-side
// (Category is MultiChoice and can't be used in an OData $filter), then slice to
// maxItems. This ceiling caps that over-fetch.
const CLIENT_FILTER_FETCH_CEILING = 200;

export class NewsRepositoryService implements INewsRepositoryService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getCategories(siteUrl: string, listTitle: string): Promise<string[]> {
    return this._getFieldChoices(siteUrl, listTitle, 'Category');
  }

  public async getItemTypes(siteUrl: string, listTitle: string): Promise<string[]> {
    return this._getFieldChoices(siteUrl, listTitle, 'ItemType');
  }

  public async getNewsItems(
    siteUrl: string,
    listTitle: string,
    filters: INewsFilters,
    maxItems: number
  ): Promise<INewsItem[]> {
    const filterByCategory = !!(filters.categories && filters.categories.length > 0);

    // ItemType is a single-value Choice column, so it can be filtered server-side.
    // Category is MultiChoice — filtered client-side below. When a category filter
    // is active we over-fetch so the post-filter slice can still reach maxItems.
    const itemTypeClause =
      filters.itemType && filters.itemType !== ANY_ITEM_TYPE
        ? `ItemType eq '${escapeODataLiteral(filters.itemType)}'`
        : '';
    const serverTop = filterByCategory ? CLIENT_FILTER_FETCH_CEILING : maxItems;

    const query: string[] = [
      `$select=${SELECT_FIELDS}`,
      `$orderby=${encodeURIComponent('PublishedDate desc')}`,
      `$top=${serverTop}`
    ];
    if (itemTypeClause) {
      query.push(`$filter=${encodeURIComponent(itemTypeClause)}`);
    }

    const url = `${trimTrailingSlash(siteUrl)}/_api/web/lists/getbytitle('${encodeListTitle(
      listTitle
    )}')/items?${query.join('&')}`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`News query failed: ${response.status} ${response.statusText}`);
    }

    const json: IItemsResponse = await response.json();
    let items = (json.value || []).map(mapListItem);

    if (filterByCategory) {
      const wanted = filters.categories;
      items = items
        .filter((item) => item.categories.some((c) => wanted.indexOf(c) >= 0))
        .slice(0, maxItems);
    }

    return items;
  }

  private async _getFieldChoices(
    siteUrl: string,
    listTitle: string,
    fieldName: string
  ): Promise<string[]> {
    const url = `${trimTrailingSlash(siteUrl)}/_api/web/lists/getbytitle('${encodeListTitle(
      listTitle
    )}')/fields/getbytitle('${fieldName}')?$select=Choices`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load '${fieldName}' choices: ${response.status} ${response.statusText}`
      );
    }

    const json: IChoicesResponse = await response.json();
    return extractChoices(json.Choices);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mapListItem(row: IRawListItem): INewsItem {
  return {
    id: row.Id,
    title: row.Title || '',
    categories: toCategoryArray(row.Category),
    itemType: row.ItemType || '',
    linkUrl: extractUrl(row.LinkUrl),
    thumbnail: parseThumbnail(row.ThumbnailImage),
    shortDescription: row.ShortDescription || '',
    publishedDate: row.PublishedDate || ''
  };
}

// Normalize a URL column value. SharePoint usually returns { Url, Description },
// but some metadata/column configurations surface a bare string. Logs only when
// it receives a non-empty object with no Url, so an unexpected shape is visible
// while the common cases stay silent (console clean on the happy path).
function extractUrl(raw: IRawUrlField | string | undefined): string {
  if (!raw) {
    return '';
  }
  if (typeof raw === 'string') {
    return raw;
  }
  if (raw.Url) {
    return raw.Url;
  }
  console.warn('[PhillipsNews] Unexpected LinkUrl shape', raw);
  return '';
}

// Normalize a field's Choices across metadata modes: a plain array under
// nometadata, or { results: [...] } under the minimalmetadata level that
// SPHttpClient negotiates. Same dual-shape problem as toCategoryArray.
function extractChoices(raw: string[] | { results?: string[] } | undefined): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw;
  }
  return Array.isArray(raw.results) ? raw.results : [];
}

// Normalize the MultiChoice Category value across metadata modes: minimal
// metadata returns a plain array, verbose returns { results: [...] }, and a bare
// string is handled defensively.
function toCategoryArray(
  raw: string[] | { results?: string[] } | string | undefined
): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw;
  }
  if (typeof raw === 'string') {
    return [raw];
  }
  return Array.isArray(raw.results) ? raw.results : [];
}

function parseThumbnail(raw: string | undefined): INewsThumbnail | undefined {
  if (!raw) {
    return undefined;
  }
  try {
    const parsed = JSON.parse(raw) as IRawThumbnail;
    if (parsed && parsed.serverRelativeUrl) {
      return { serverRelativeUrl: parsed.serverRelativeUrl, alt: parsed.alt || '' };
    }
  } catch {
    // Malformed Image-column JSON: treat as no thumbnail.
  }
  return undefined;
}

// Single quotes inside an OData string literal are escaped by doubling them.
function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}

// List title goes inside getbytitle('...'); URL-encode so spaces and other
// characters are safe in the request URL. Apostrophes are also doubled so the
// OData literal stays valid after decoding.
function encodeListTitle(listTitle: string): string {
  return encodeURIComponent(escapeODataLiteral(listTitle));
}

function trimTrailingSlash(url: string): string {
  return url.replace(/\/+$/, '');
}
