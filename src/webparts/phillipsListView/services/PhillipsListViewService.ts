import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { IPhillipsListViewService } from './IPhillipsListViewService';
import { IListInfo, IViewInfo, IFieldInfo, ITabData, IListRow } from './models';

// Field types whose values are stored as SharePoint Image / Thumbnail columns.
// Drives the conditional sidecar AttachmentFiles fetch (only needed when an
// image field appears in the view; skipping the sidecar avoids the unrelated
// 500s observed for Partner Profiles in 1.0.1.3).
const IMAGE_FIELD_TYPES = new Set<string>(['Thumbnail', 'Image']);

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

interface IRenderListDataResponse {
  Row?: IListRow[];
}

interface IAttachmentLookupResponse {
  value?: Array<{
    Id?: number | string;
    ID?: number | string;
    AttachmentFiles?: unknown;
  }>;
}

export class PhillipsListViewService implements IPhillipsListViewService {
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
    const viewXml = buildViewXml(view.viewQuery, allViewFields, view.rowLimit);

    // RenderListDataAsStream is the primary items query — honors the view's
    // CAML (filter / sort / row limit) and resolves Lookup + User columns to
    // [{lookupId, lookupValue}] arrays so display names show without a second
    // round-trip. 1.0.1.3 briefly used GetItems + OData $expand, which
    // returned 500s for Partner Profiles in production despite working for
    // Awards — root cause unconfirmed without DevTools, but reverting here
    // restores the path that already proved compatible with both lists.
    const rows = await this._renderListData(siteUrl, listId, viewXml);

    // Conditional sidecar: SharePoint Image columns store only a fileName in
    // their column value, with the real URL living in the item's
    // AttachmentFiles. Only fetch that when the view actually includes an
    // image-typed field, so lists without image columns (Partner Profiles,
    // Celebrations views) skip the second call entirely.
    const fieldsByName = new Map<string, IFieldInfo>();
    for (const f of fields) {
      fieldsByName.set(f.internalName, f);
    }
    const hasImageField = allViewFields.some((name) => {
      const meta = fieldsByName.get(name);
      return !!meta && IMAGE_FIELD_TYPES.has(meta.typeAsString);
    });
    if (hasImageField && rows.length > 0) {
      await this._enrichRowsWithAttachments(siteUrl, listId, rows);
    }

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

  // Sidecar OData call: fetches AttachmentFiles for the items already
  // returned by RenderListDataAsStream and writes them into each row by Id.
  // Only called when an image-typed field is in the view's selection (see
  // `hasImageField` in getTabData), so non-image lists skip the round-trip.
  private async _enrichRowsWithAttachments(
    siteUrl: string,
    listId: string,
    rows: IListRow[]
  ): Promise<void> {
    const ids: number[] = [];
    for (const row of rows) {
      const id = rowId(row);
      if (id !== undefined) {
        ids.push(id);
      }
    }
    if (ids.length === 0) {
      return;
    }

    // Build $filter=Id eq 1 or Id eq 2 or ... — SP REST OData doesn't
    // support the `in` operator, so we OR the IDs together. URL length is
    // bounded by the view's RowLimit (default 30 ⇒ ~250 chars), well under
    // SharePoint's request-line cap.
    const filterClause = ids.map((id) => `Id eq ${id}`).join(' or ');
    const url =
      `${trimSlash(siteUrl)}/_api/web/lists(guid'${listId}')/items` +
      `?$select=Id,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl` +
      `&$expand=AttachmentFiles` +
      `&$filter=${encodeURIComponent(filterClause)}` +
      `&$top=${ids.length}`;

    const json = await this._getJson<IAttachmentLookupResponse>(url, 'attachments');
    const byId = new Map<number, unknown>();
    for (const item of json.value || []) {
      const id = item.Id ?? item.ID;
      if (typeof id === 'number') {
        byId.set(id, item.AttachmentFiles);
      } else if (typeof id === 'string') {
        const parsed = Number(id);
        if (!isNaN(parsed)) {
          byId.set(parsed, item.AttachmentFiles);
        }
      }
    }

    for (const row of rows) {
      const id = rowId(row);
      if (id !== undefined && byId.has(id)) {
        (row as { AttachmentFiles?: unknown }).AttachmentFiles = byId.get(id);
      }
    }
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

// RenderListDataAsStream returns Id under both `ID` (uppercase) and `Id`
// depending on the field; normalize so the sidecar can match item IDs.
function rowId(row: IListRow): number | undefined {
  const candidate = row.ID ?? row.Id;
  if (typeof candidate === 'number') {
    return candidate;
  }
  if (typeof candidate === 'string') {
    const n = Number(candidate);
    if (!isNaN(n)) {
      return n;
    }
  }
  return undefined;
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
