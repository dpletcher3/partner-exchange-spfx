import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { IFieldMapping, IMediaCardItem } from './models';
import { extractUrl, extractImageColumnUrl } from './extractors';

const LOG = '[MediaGallery]';
const MAX_ITEMS = 500;

interface IItemsResponse {
  value: Array<Record<string, unknown>>;
}

// Reads card items from the mapped list. All field access is by the mapped
// internal names (never hardcoded), so the same service serves any list the web
// part is pointed at. AttachmentFiles is expanded so the Image-column extractor
// can resolve "reserved attachment" image values to a real URL.
export class PhillipsMediaGalleryService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getItems(
    siteUrl: string,
    listId: string,
    mapping: IFieldMapping
  ): Promise<IMediaCardItem[]> {
    const base = siteUrl.replace(/\/+$/, '');
    const { titleField, videoField, labelImageField, mainImageField } = mapping;

    const selectFields = ['Id', titleField, videoField, labelImageField];
    if (mainImageField) {
      selectFields.push(mainImageField);
    }
    selectFields.push('AttachmentFiles/FileName', 'AttachmentFiles/ServerRelativeUrl');

    const query = [
      `$select=${selectFields.join(',')}`,
      `$expand=AttachmentFiles`,
      `$orderby=${encodeURIComponent(`${titleField} asc`)}`, // alphabetical by mapped title
      `$top=${MAX_ITEMS}`
    ];
    const url = `${base}/_api/web/lists(guid'${listId}')/items?${query.join('&')}`;

    console.log(
      `${LOG} fetching items: list=${listId}, orderby="${titleField} asc", select=[${selectFields.join(', ')}]`
    );

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
    const items = rows.map((row) => this._mapItem(row, mapping));
    console.log(`${LOG} mapped ${items.length} items from list ${listId}`);
    return items;
  }

  private _mapItem(row: Record<string, unknown>, mapping: IFieldMapping): IMediaCardItem {
    const id = typeof row.Id === 'number' ? row.Id : Number(row.Id) || 0;
    const titleVal = row[mapping.titleField];
    const title = typeof titleVal === 'string' ? titleVal : '';
    const videoUrl = extractUrl(row[mapping.videoField]);
    const labelImageUrl = extractImageColumnUrl(row[mapping.labelImageField], row.AttachmentFiles);
    const mainImageOverrideUrl = mapping.mainImageField
      ? extractImageColumnUrl(row[mapping.mainImageField], row.AttachmentFiles)
      : undefined;

    return { id, title, videoUrl, labelImageUrl, mainImageOverrideUrl };
  }
}
