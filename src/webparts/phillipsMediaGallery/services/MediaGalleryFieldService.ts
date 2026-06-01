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

// Fetches a list's columns for the field-mapping dropdowns. Mirrors the proven
// PhillipsListView service pattern (spHttpClient + /_api/web/lists(guid'…')) so
// we don't depend on @pnp PropertyFieldColumnPicker, which fails silently
// (SPColumnPickerService returns undefined when orderBy is null, and its host
// swallows every error with `.catch(() => {})`).
export class MediaGalleryFieldService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getColumns(siteUrl: string, listId: string): Promise<IColumnInfo[]> {
    const base = siteUrl.replace(/\/+$/, '');
    // Hidden eq false keeps the dropdown short; ReadOnlyField eq false drops
    // computed/system columns (e.g. ContentType) that can't back a mapping.
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
      .filter((f) => !!f.InternalName)
      .map((f) => ({ internalName: f.InternalName, displayName: f.Title || f.InternalName }));
  }
}
