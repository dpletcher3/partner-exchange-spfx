import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IColumnInfo {
  internalName: string;
  displayName: string;
}

export interface IListItemRef {
  id: number;
  title: string;
}

interface IRawField {
  InternalName: string;
  Title: string;
}
interface IFieldsResponse {
  value: IRawField[];
}

interface IRawItem {
  Id: number;
  Title?: string;
}
interface IItemsResponse {
  value: IRawItem[];
}

const MAX_ITEMS = 500;

// Property-pane data for Highlight Video. Mirrors the Media Card Gallery's
// proven pattern (spHttpClient + the correct /_api/web/lists(guid'…') endpoint,
// not the /_api/lists endpoint the @pnp pickers hit and 404 on silently).
//   - getColumns → the field-mapping dropdowns (Title/Video/Info)
//   - getListItems → the item picker (one featured item, listed by Title)
export class HighlightVideoService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getColumns(siteUrl: string, listId: string): Promise<IColumnInfo[]> {
    const base = siteUrl.replace(/\/+$/, '');
    const filter = encodeURIComponent('Hidden eq false and ReadOnlyField eq false');
    const url =
      `${base}/_api/web/lists(guid'${listId}')/fields` +
      `?$select=InternalName,Title&$filter=${filter}&$orderby=Title&$top=500`;
    console.log(`[HighlightVideo] getColumns URL: ${url}`);

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(
        `fields fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
      );
    }
    const json = (await response.json()) as IFieldsResponse;
    const rows = json && json.value ? json.value : [];
    const cols = rows
      .filter((f) => !!f.InternalName)
      .map((f) => ({ internalName: f.InternalName, displayName: f.Title || f.InternalName }));
    console.log(`[HighlightVideo] getColumns returned ${cols.length} columns`);
    return cols;
  }

  public async getListItems(siteUrl: string, listId: string): Promise<IListItemRef[]> {
    const base = siteUrl.replace(/\/+$/, '');
    const url =
      `${base}/_api/web/lists(guid'${listId}')/items` +
      `?$select=Id,Title&$orderby=${encodeURIComponent('Title asc')}&$top=${MAX_ITEMS}`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(
        `items fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
      );
    }
    const json = (await response.json()) as IItemsResponse;
    const rows = json && json.value ? json.value : [];
    return rows.map((r) => ({ id: r.Id, title: r.Title || '' }));
  }
}
