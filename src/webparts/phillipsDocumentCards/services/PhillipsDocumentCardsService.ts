import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { IFieldMapping, IDocCardItem } from './models';
import { extractImageColumnUrl, asString } from './extractors';

const LOG = '[DocumentCards]';
const MAX_ITEMS = 500;

interface IItemsResponse {
  value: Array<Record<string, unknown>>;
}

// Reads document-card items from the mapped document library, ONE column at a
// time: each on-screen column server-side $filters the library on its configured
// DocSection value (single-value Choice filters cleanly in OData — confirmed in
// the I16 Stage 1 probe). All field access is by the mapped internal names
// (never hardcoded), so the same service serves any library the web part is
// pointed at. The CardIcon Image column stores an inline serverRelativeUrl, so
// NO AttachmentFiles expand is used — that expand is invalid on a document
// library (items carry no list attachments) and 400s the whole fetch with
// "Column 'Attachments' does not exist". The inline-URL extractor branch needs
// no attachment lookup.
export class PhillipsDocumentCardsService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getItemsForColumn(
    siteUrl: string,
    listId: string,
    mapping: IFieldMapping,
    filterValue: string
  ): Promise<IDocCardItem[]> {
    const base = siteUrl.replace(/\/+$/, '');
    const titleField = mapping.titleField || 'FileLeafRef';

    // FileLeafRef (name) + FileRef (server-relative URL, the click target) are
    // built-in. De-dup so a title remapped onto a built-in doesn't double up.
    const selectFields = ['Id', 'FileLeafRef', 'FileRef', titleField, mapping.descriptionField, mapping.iconField, mapping.sectionField]
      .filter((f, i, arr) => !!f && arr.indexOf(f) === i);

    // OData eq on a single-value Choice; double any single quotes in the value.
    const filterClause = `${mapping.sectionField} eq '${filterValue.replace(/'/g, "''")}'`;

    const query = [
      `$select=${selectFields.join(',')}`,
      `$filter=${encodeURIComponent(filterClause)}`,
      `$orderby=${encodeURIComponent('FileLeafRef asc')}`,
      `$top=${MAX_ITEMS}`
    ];
    const url = `${base}/_api/web/lists(guid'${listId}')/items?${query.join('&')}`;

    console.log(`${LOG} fetching column items: list=${listId}, filter="${filterClause}", select=[${selectFields.join(', ')}]`);

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
    console.log(`${LOG} mapped ${items.length} items for filter "${filterClause}"`);
    return items;
  }

  private _mapItem(row: Record<string, unknown>, mapping: IFieldMapping): IDocCardItem {
    const id = typeof row.Id === 'number' ? row.Id : Number(row.Id) || 0;
    const leaf = asString(row.FileLeafRef);

    const titleField = mapping.titleField || 'FileLeafRef';
    let title: string;
    if (titleField === 'FileLeafRef') {
      title = stripExtension(leaf);
    } else {
      const mapped = asString(row[titleField]);
      title = mapped || stripExtension(leaf);
    }

    const description = asString(row[mapping.descriptionField]);
    // Inline-URL Image column: the blob carries serverRelativeUrl, so no
    // attachment-files lookup is needed (pass undefined for that argument).
    const iconUrl = extractImageColumnUrl(row[mapping.iconField], undefined);
    const docUrl = asString(row.FileRef);
    const section = asString(row[mapping.sectionField]);

    return { id, title, description, iconUrl, docUrl, section };
  }
}

// "Onboarding Guide.txt" -> "Onboarding Guide". Leaves dot-less names untouched.
function stripExtension(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.substring(0, dot) : name;
}
