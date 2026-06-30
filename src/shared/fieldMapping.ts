import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';

// Shared field-mapping module — the union of what the Media Card Gallery,
// Highlight Video, and Celebrations web parts each need from a property-pane
// "map a list column to a role" group. Extracted (D036 / lessons-learned) so a
// fourth web part reuses rather than re-cloning the pattern.
//
// Carries over exactly the proven behaviors:
//   - fetch columns from the CORRECT /_api/web/lists(guid'…')/fields endpoint
//     (never @pnp PropertyFieldColumnPicker, which hits /_api/lists and 404s
//     silently — see lessons-learned);
//   - dependent-pane refresh on list change (clear → fetch → propertyPane.refresh());
//   - first-list-selection preserves slot defaults; only clears when switching
//     between two real lists;
//   - the mapping group renders EXPANDED (a collapsed group reads as empty);
//   - loud [shared] console diagnostics (fetch URL / column count / branch).
//
// NOT wired into the gallery/highlight parts here — those migrate later.

const LOG = '[shared]';

// Column-type restriction for a slot (e.g. only Date columns, only Person columns).
export type ColumnTypeFilter = 'any' | 'date' | 'person' | 'text' | 'url' | 'choice';

export interface IFieldSlot {
  // The web-part property that stores the chosen internal name.
  property: string;
  label: string;
  // Documented default internal name (also set in the manifest preconfigured
  // properties; the data layer falls back to it when the property is unset).
  defaultInternalName: string;
  // Restrict the dropdown to columns of this kind. Omitted / 'any' = no filter.
  typeFilter?: ColumnTypeFilter;
}

export interface IColumnInfo {
  internalName: string;
  displayName: string;
  typeAsString: string;
}

interface IRawField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
}
interface IFieldsResponse {
  value: IRawField[];
}

// Fetches a list's columns (with type) from the correct web-scoped endpoint.
export async function fetchListColumns(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listId: string
): Promise<IColumnInfo[]> {
  const base = siteUrl.replace(/\/+$/, '');
  const filter = encodeURIComponent('Hidden eq false and ReadOnlyField eq false');
  const url =
    `${base}/_api/web/lists(guid'${listId}')/fields` +
    `?$select=InternalName,Title,TypeAsString&$filter=${filter}&$orderby=Title&$top=500`;
  console.log(`${LOG} fetchListColumns URL: ${url}`);

  const response: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1);
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
    .map((f) => ({
      internalName: f.InternalName,
      displayName: f.Title || f.InternalName,
      typeAsString: f.TypeAsString
    }));
  console.log(`${LOG} fetchListColumns returned ${cols.length} columns`);
  return cols;
}

function matchesType(col: IColumnInfo, filter?: ColumnTypeFilter): boolean {
  if (!filter || filter === 'any') {
    return true;
  }
  switch (filter) {
    case 'date':
      return col.typeAsString === 'DateTime';
    case 'person':
      return col.typeAsString === 'User' || col.typeAsString === 'UserMulti';
    case 'text':
      return col.typeAsString === 'Text' || col.typeAsString === 'Note';
    case 'url':
      return col.typeAsString === 'URL';
    case 'choice':
      return col.typeAsString === 'Choice' || col.typeAsString === 'MultiChoice';
    default:
      return true;
  }
}

// Resolves a list's GUID from its title on a (possibly cross-site) web. Throws on
// failure so the caller can fall back / fail-closed. Used when a web part targets
// a list by title on another site (no list picker available cross-site).
export async function resolveListIdByTitle(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listTitle: string
): Promise<string> {
  const base = siteUrl.replace(/\/+$/, '');
  const url = `${base}/_api/web/lists/getByTitle('${listTitle.replace(/'/g, "''")}')?$select=Id`;
  console.log(`${LOG} resolveListIdByTitle URL: ${url}`);
  const response: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(
      `list resolve failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
    );
  }
  const json = (await response.json()) as { Id?: string };
  return json && typeof json.Id === 'string' ? json.Id : '';
}

interface IChoiceFieldResponse {
  Choices?: string[] | { results?: string[] };
}

// Fetches the choice values of a single Choice/MultiChoice column, so a web part
// can drive a LIVE option list (new choices appear automatically) instead of a
// hardcoded array. Returns [] for a non-choice column.
export async function fetchChoiceFieldValues(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listId: string,
  internalName: string
): Promise<string[]> {
  const base = siteUrl.replace(/\/+$/, '');
  const url =
    `${base}/_api/web/lists(guid'${listId}')/fields/getByInternalNameOrTitle('${internalName.replace(/'/g, "''")}')` +
    `?$select=Choices`;
  console.log(`${LOG} fetchChoiceFieldValues URL: ${url}`);
  const response: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(
      `choices fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
    );
  }
  const json = (await response.json()) as IChoiceFieldResponse;
  const raw = json ? json.Choices : undefined;
  const list = Array.isArray(raw) ? raw : raw && Array.isArray(raw.results) ? raw.results : [];
  return list.filter((c): c is string => typeof c === 'string');
}

export interface IFieldMappingControllerOptions {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  slots: IFieldSlot[];
  // Current list id (closure so the controller always reads live state).
  getListId: () => string;
  // The web part's property bag (mutated to clear slot values on list switch).
  properties: Record<string, unknown>;
  // this.context.propertyPane.refresh
  refresh: () => void;
  // Label shown in place of the dropdowns when no list is selected.
  hintLabel?: string;
}

// Owns the available-columns state and the dependent-pane lifecycle for a set of
// mapping slots. A web part composes one of these and delegates to it.
export class FieldMappingController {
  private _available: IColumnInfo[] = [];
  private _loadedFor: string | undefined = undefined;
  private _loading = false;
  private readonly _o: IFieldMappingControllerOptions;

  public constructor(options: IFieldMappingControllerOptions) {
    this._o = options;
  }

  public get columnsReady(): boolean {
    const id = this._o.getListId();
    return !!id && this._loadedFor === id;
  }

  // Fetch the current list's columns (idempotent; in-flight guarded). Call from
  // onPropertyPaneConfigurationStart and after a list change.
  public load(): void {
    const listId = this._o.getListId();
    if (!listId) {
      this._available = [];
      this._loadedFor = undefined;
      return;
    }
    if (this._loadedFor === listId || this._loading) {
      return;
    }
    this._loading = true;
    const target = listId;
    fetchListColumns(this._o.spHttpClient, this._o.siteUrl, target)
      .then((cols) => {
        if (this._o.getListId() !== target) {
          return;
        }
        this._available = cols;
        this._loadedFor = target;
        console.log(`${LOG} loaded ${cols.length} columns for listId=${target}`);
        this._o.refresh();
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} column fetch FAILED for listId=${target}`, err);
        if (this._o.getListId() !== target) {
          return;
        }
        this._available = [];
        this._loadedFor = target;
        this._o.refresh();
      })
      .then(() => {
        this._loading = false;
      })
      .catch(() => {
        /* keep the promise non-floating */
      });
  }

  // Call from onPropertyPaneFieldChanged when the bound list property changes.
  // `oldListId` distinguishes a first selection (preserve defaults) from a switch
  // (clear stale slot mappings).
  public onListChanged(oldListId: unknown): void {
    if (oldListId) {
      for (const slot of this._o.slots) {
        this._o.properties[slot.property] = '';
      }
      console.log(`${LOG} list switched — cleared ${this._o.slots.length} slot mappings`);
    } else {
      console.log(`${LOG} first list selection — preserving slot defaults`);
    }
    this._available = [];
    this._loadedFor = undefined;
    this.load();
    this._o.refresh();
  }

  // The mapping group's fields (render in an EXPANDED group). When no list is
  // selected, a single hint label stands in for the dropdowns.
  public buildFields(): IPropertyPaneField<unknown>[] {
    const listId = this._o.getListId();
    const branch = !listId ? 'no-list (hint)' : this.columnsReady ? 'populated' : 'loading';
    console.log(`${LOG} field-mapping branch=${branch}, columns=${this._available.length}`);

    if (!listId) {
      return [
        PropertyPaneLabel('sharedFieldMappingHint', {
          text: this._o.hintLabel || 'Select a list above to map its columns.'
        })
      ];
    }
    return this._o.slots.map((slot) => this._dropdown(slot));
  }

  private _dropdown(slot: IFieldSlot): IPropertyPaneField<unknown> {
    const ready = this.columnsReady;
    const filtered = this._available.filter((c) => matchesType(c, slot.typeFilter));

    let options: IPropertyPaneDropdownOption[];
    if (!ready) {
      options = [{ key: '', text: 'Loading columns…' }];
    } else if (filtered.length === 0) {
      const kind = slot.typeFilter && slot.typeFilter !== 'any' ? `${slot.typeFilter} ` : '';
      options = [{ key: '', text: `(no ${kind}columns found on this list)` }];
    } else {
      options = filtered.map((c) => ({ key: c.internalName, text: c.displayName }));
    }

    const stored = (this._o.properties[slot.property] as string) || '';
    return PropertyPaneDropdown(slot.property, {
      label: slot.label,
      options,
      selectedKey: stored || undefined,
      disabled: !ready || filtered.length === 0
    });
  }
}
