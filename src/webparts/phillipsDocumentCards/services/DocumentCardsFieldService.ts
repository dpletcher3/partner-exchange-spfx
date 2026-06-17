import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IColumnInfo {
  internalName: string;
  displayName: string;
}

interface IRawField {
  InternalName: string;
  Title: string;
}

interface IFieldsResponse {
  value: IRawField[];
}

// Fetches a library's columns for the field-mapping dropdowns. Mirrors the
// proven Media Gallery / PhillipsListView pattern (spHttpClient +
// /_api/web/lists(guid'…')/fields) rather than @pnp PropertyFieldColumnPicker,
// which fetches the wrong endpoint and fails silently (per D044 / lessons-learned).
export class DocumentCardsFieldService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getColumns(siteUrl: string, listId: string): Promise<IColumnInfo[]> {
    const base = siteUrl.replace(/\/+$/, '');
    // Hidden eq false keeps the dropdown short; ReadOnlyField eq false drops
    // computed/system columns that can't back a mapping.
    const filter = encodeURIComponent('Hidden eq false and ReadOnlyField eq false');
    const url =
      `${base}/_api/web/lists(guid'${listId}')/fields` +
      `?$select=InternalName,Title&$filter=${filter}&$orderby=Title&$top=500`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(`fields fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`);
    }

    const json = (await response.json()) as IFieldsResponse;
    const rows = json && json.value ? json.value : [];
    return rows
      // Exclude SharePoint system columns (internal names starting with '_').
      // The built-in document-library "Description" column is internal name
      // _ExtendedDescription — it shares the display name of our custom
      // CardDescription, so without this filter the mapping dropdown shows two
      // identical "Description" options and the wrong one (which 400s the items
      // $select with "field or property '_ExtendedDescription' does not exist")
      // can be selected. No mapping target legitimately starts with '_'.
      .filter((f) => !!f.InternalName && f.InternalName.charAt(0) !== '_')
      .map((f) => ({ internalName: f.InternalName, displayName: f.Title || f.InternalName }));
  }
}
