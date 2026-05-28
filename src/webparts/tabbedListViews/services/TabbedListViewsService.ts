import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { ITabbedListViewsService } from './ITabbedListViewsService';
import { IListInfo, IViewInfo, IFieldInfo, ITabData, IListRow } from './models';

// Raw OData shapes for the requests below — declared locally so the mapping
// code avoids `any`. Minimal-metadata serialization is assumed (SPHttpClient
// default); collection-valued properties that arrive as { results: [...] }
// under verbose metadata are handled defensively where they appear.

interface IRawList {
  Id: string;
  Title: string;
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

interface IRenderListDataResponse {
  Row?: IListRow[];
}

export class TabbedListViewsService implements ITabbedListViewsService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getLists(siteUrl: string): Promise<IListInfo[]> {
    // BaseTemplate 100 = Generic List. Filters out libraries (101), tasks
    // (107), and other system list types. Hidden lists are also excluded so
    // the editor doesn't see SP system catalogs.
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists` +
      `?$filter=Hidden eq false and BaseTemplate eq 100` +
      `&$select=Id,Title` +
      `&$orderby=Title`;

    const json = await this._getJson<IListValueResponse<IRawList>>(url, 'lists');
    return (json.value || []).map((row) => ({ id: row.Id, title: row.Title }));
  }

  public async getViews(siteUrl: string, listId: string): Promise<IViewInfo[]> {
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/views` +
      `?$filter=Hidden eq false and PersonalView eq false` +
      `&$select=Id,Title` +
      `&$orderby=Title`;

    const json = await this._getJson<IListValueResponse<IRawView>>(url, 'views');
    return (json.value || []).map((row) => ({ id: row.Id, title: row.Title }));
  }

  public async getFields(siteUrl: string, listId: string): Promise<IFieldInfo[]> {
    // Hidden eq false keeps the dropdown short; calculated fields stay in
    // because they're often the overlay source (e.g. TenuredChampionMilestone).
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/fields` +
      `?$filter=Hidden eq false` +
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

    const allFields = uniqueStrings([...view.viewFields, ...extraFields]);
    const viewXml = buildViewXml(view.viewQuery, allFields, view.rowLimit);
    const rows = await this._renderListData(siteUrl, listId, viewXml);

    const fieldDisplayNames: { [internalName: string]: string } = {};
    for (const f of fields) {
      fieldDisplayNames[f.internalName] = f.displayName;
    }

    return {
      viewFields: view.viewFields,
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

  private async _renderListData(
    siteUrl: string,
    listId: string,
    viewXml: string
  ): Promise<IListRow[]> {
    const url = `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/RenderListDataAsStream`;
    const body = JSON.stringify({ parameters: { ViewXml: viewXml } });

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

    const json: IRenderListDataResponse = await response.json();
    return json.Row || [];
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
