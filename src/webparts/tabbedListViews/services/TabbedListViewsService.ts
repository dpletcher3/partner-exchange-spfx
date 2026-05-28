import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { ITabbedListViewsService } from './ITabbedListViewsService';
import { IListInfo, IViewInfo, IFieldInfo, ITabData, IListRow } from './models';

// Field types whose values benefit from $expand=<field>&$select=<field>/Title
// so OData returns the display name instead of just the Id. Without these,
// Recipient (Lookup) and Author (User) come back as raw integers.
const EXPANDABLE_FIELD_TYPES = new Set<string>([
  'Lookup',
  'LookupMulti',
  'User',
  'UserMulti'
]);

// Raw OData shapes for the requests below — declared locally so the mapping
// code avoids `any`. Minimal-metadata serialization is assumed (SPHttpClient
// default); collection-valued properties that arrive as { results: [...] }
// under verbose metadata are handled defensively where they appear.

interface IRawList {
  Id: string;
  Title: string;
  BaseTemplate: number;
}

interface IRawView {
  Id: string;
  Title: string;
}

interface IRawField {
  InternalName: string;
  Title: string;
  FieldTypeKind: number;
  TypeAsString: string;
}

interface IListValueResponse<T> {
  value: T[];
}

interface IRawViewDefinition {
  ViewQuery?: string;
  RowLimit?: number;
  ViewFields?: unknown;
}

interface IGetItemsResponse {
  value?: IListRow[];
}

export class TabbedListViewsService implements ITabbedListViewsService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getLists(siteUrl: string): Promise<IListInfo[]> {
    // Show all visible lists; the editor reads the title and picks the right
    // one. We deliberately don't filter by BaseTemplate — different
    // provisioning paths (PnP, m365 CLI) produce lists with different
    // templates, and an over-tight filter is what hid Awards / Partner
    // Profiles from the dropdown in 1.0.1.0.
    // OData filter values are URL-encoded so the embedded spaces don't get
    // mangled by SPHttpClient (PhillipsNews uses the same pattern).
    const filter = encodeURIComponent('Hidden eq false');
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists` +
      `?$filter=${filter}` +
      `&$select=Id,Title,BaseTemplate` +
      `&$orderby=Title`;

    const json = await this._getJson<IListValueResponse<IRawList>>(url, 'lists');
    // Drop the worst noise: document libraries (101), wiki pages (119), site
    // pages (850/851), workflow history (140), form templates (117). Keep
    // everything else so a list with an unexpected template still shows up.
    const NOISY_TEMPLATES = new Set<number>([101, 117, 119, 140, 850, 851]);
    return (json.value || [])
      .filter((row) => !NOISY_TEMPLATES.has(row.BaseTemplate))
      .map((row) => ({ id: row.Id, title: row.Title }));
  }

  public async getViews(siteUrl: string, listId: string): Promise<IViewInfo[]> {
    const filter = encodeURIComponent('Hidden eq false and PersonalView eq false');
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/views` +
      `?$filter=${filter}` +
      `&$select=Id,Title` +
      `&$orderby=Title`;

    const json = await this._getJson<IListValueResponse<IRawView>>(url, 'views');
    return (json.value || []).map((row) => ({ id: row.Id, title: row.Title }));
  }

  public async getFields(siteUrl: string, listId: string): Promise<IFieldInfo[]> {
    // Hidden eq false keeps the dropdown short; calculated fields stay in
    // because they're often the overlay source (e.g. TenuredChampionMilestone).
    const filter = encodeURIComponent('Hidden eq false');
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/fields` +
      `?$filter=${filter}` +
      `&$select=InternalName,Title,FieldTypeKind,TypeAsString` +
      `&$orderby=Title`;

    const json = await this._getJson<IListValueResponse<IRawField>>(url, 'fields');
    return (json.value || []).map((row) => ({
      internalName: row.InternalName,
      displayName: row.Title,
      typeKind: row.FieldTypeKind,
      typeAsString: row.TypeAsString
    }));
  }

  public async getTabData(
    siteUrl: string,
    listId: string,
    viewId: string,
    extraFields: string[]
  ): Promise<ITabData> {
    // Pull view + field metadata in parallel — they're independent, and the
    // component blocks on both before rendering.
    const [view, fields] = await Promise.all([
      this._getViewDefinition(siteUrl, listId, viewId),
      this.getFields(siteUrl, listId)
    ]);

    const allViewFields = uniqueStrings([...view.viewFields, ...extraFields]);
    const fieldsByName = new Map<string, IFieldInfo>();
    for (const f of fields) {
      fieldsByName.set(f.internalName, f);
    }

    // Build $select + $expand. Lookups/Users get expanded so the display name
    // arrives instead of the raw Id. AttachmentFiles is always expanded so
    // SharePoint Image columns (which store only a fileName in the column
    // value) can be resolved to a real URL on the client.
    const selectParts: string[] = ['Id'];
    const expandParts: string[] = [];
    for (const fieldName of allViewFields) {
      const meta = fieldsByName.get(fieldName);
      if (meta && EXPANDABLE_FIELD_TYPES.has(meta.typeAsString)) {
        selectParts.push(`${fieldName}/Title`);
        expandParts.push(fieldName);
      } else {
        selectParts.push(fieldName);
      }
    }
    selectParts.push('AttachmentFiles/FileName', 'AttachmentFiles/ServerRelativeUrl');
    expandParts.push('AttachmentFiles');

    const viewXml = buildViewXml(view.viewQuery, allViewFields, view.rowLimit);
    const rows = await this._getItems(
      siteUrl,
      listId,
      viewXml,
      selectParts,
      expandParts
    );

    const fieldDisplayNames: { [internalName: string]: string } = {};
    for (const f of fields) {
      fieldDisplayNames[f.internalName] = f.displayName;
    }

    return {
      viewFields: view.viewFields,
      fields,
      fieldDisplayNames,
      rows
    };
  }

  // -------------------------------------------------------------------------
  // Internals
  // -------------------------------------------------------------------------

  private async _getViewDefinition(
    siteUrl: string,
    listId: string,
    viewId: string
  ): Promise<{ viewFields: string[]; viewQuery: string; rowLimit: number }> {
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/views(guid'${viewId}')` +
      `?$select=ViewQuery,RowLimit,ViewFields&$expand=ViewFields`;

    const json = await this._getJson<IRawViewDefinition>(url, 'view');
    return {
      viewFields: extractViewFieldItems(json.ViewFields),
      viewQuery: json.ViewQuery || '',
      // SharePoint defaults views to 30 rows; matching that ceiling here keeps
      // a misconfigured view from accidentally pulling every row.
      rowLimit: json.RowLimit && json.RowLimit > 0 ? json.RowLimit : 30
    };
  }

  private async _getItems(
    siteUrl: string,
    listId: string,
    viewXml: string,
    selectParts: string[],
    expandParts: string[]
  ): Promise<IListRow[]> {
    // GetItems honors the view's CAML query (filter, sort, row limit) AND
    // supports OData $select/$expand in the URL. RenderListDataAsStream (used
    // up through 1.0.1.2) honored CAML but couldn't expand AttachmentFiles,
    // which is exactly what SharePoint Image columns need to resolve their
    // fileName to a URL — see PhillipsNews for the same pattern.
    const select = encodeURIComponent(uniqueStrings(selectParts).join(','));
    const expand = encodeURIComponent(uniqueStrings(expandParts).join(','));
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/GetItems` +
      `?$select=${select}&$expand=${expand}`;

    // __metadata.type is required by SharePoint for the CamlQuery payload
    // shape under minimal-metadata; omitting it produces a 400.
    const body = JSON.stringify({
      query: {
        __metadata: { type: 'SP.CamlQuery' },
        ViewXml: viewXml
      }
    });

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json' },
        body
      }
    );

    if (!response.ok) {
      throw new Error(`Items query failed: ${response.status} ${response.statusText}`);
    }

    const json: IGetItemsResponse = await response.json();
    return json.value || [];
  }

  private async _getJson<T>(url: string, what: string): Promise<T> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to load ${what}: ${response.status} ${response.statusText}`);
    }

    return response.json() as Promise<T>;
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function trimSlash(url: string): string {
  return url.replace(/\/+$/, '');
}

function uniqueStrings(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of values) {
    if (v && !seen.has(v)) {
      seen.add(v);
      out.push(v);
    }
  }
  return out;
}

function escapeXmlAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildViewXml(viewQuery: string, fields: string[], rowLimit: number): string {
  const fieldRefs = fields
    .map((f) => `<FieldRef Name="${escapeXmlAttribute(f)}" />`)
    .join('');
  // RenderListDataAsStream expects a complete <View> XML. ViewQuery from the
  // view definition is the *inner* CAML (the contents of <Query>), so we wrap
  // it here. RowLimit goes after ViewFields per SharePoint's schema.
  return (
    `<View>` +
    `<Query>${viewQuery}</Query>` +
    `<ViewFields>${fieldRefs}</ViewFields>` +
    `<RowLimit>${rowLimit}</RowLimit>` +
    `</View>`
  );
}

// The ViewFields collection on a view object arrives in one of a few shapes
// depending on metadata level: { Items: [...] }, { Items: { results: [...] } },
// or { results: [...] }. Normalize to a flat string array.
function extractViewFieldItems(raw: unknown): string[] {
  if (!raw || typeof raw !== 'object') {
    return [];
  }
  const obj = raw as { [key: string]: unknown };

  const items = obj.Items;
  if (Array.isArray(items)) {
    return items.filter(isNonEmptyString);
  }
  if (items && typeof items === 'object') {
    const inner = (items as { [key: string]: unknown }).results;
    if (Array.isArray(inner)) {
      return inner.filter(isNonEmptyString);
    }
  }
  const results = obj.results;
  if (Array.isArray(results)) {
    return results.filter(isNonEmptyString);
  }
  return [];
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0;
}
