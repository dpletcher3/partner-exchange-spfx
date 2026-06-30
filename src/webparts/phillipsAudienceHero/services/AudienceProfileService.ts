import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { IProfileLookupOptions } from './models';

const LOG = '[AudienceHero]';
const MAX_ITEMS = 500;

interface IRawUser {
  EMail?: string;
}
interface IRawRow {
  [key: string]: unknown;
}
interface IItemsResponse {
  value?: IRawRow[];
}

function coerceString(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

// Resolves the viewer's single Partner Profiles Division WITHOUT Microsoft Graph:
// identity comes from pageContext (the caller passes the email), and the division
// comes from a SharePoint REST query against Partner Profiles on its own site
// (cross-site from The Hub). Person-field $filter on EMail is tried first; if the
// tenant rejects that filter shape, it falls back to fetching rows and matching
// client-side. Returns undefined when no row matches (caller fail-closes); throws
// only when BOTH paths fail at the network/REST level.
export class AudienceProfileService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getViewerDivision(o: IProfileLookupOptions): Promise<string | undefined> {
    const email = (o.viewerEmail || '').trim();
    if (!email) {
      console.warn(`${LOG} no viewer email on pageContext — fail-closed (no division)`);
      return undefined;
    }
    const base = o.siteUrl.replace(/\/+$/, '');

    // Primary: server-side person-field EMail filter.
    try {
      const division = await this._filterByEmail(base, o, email);
      console.log(`${LOG} lookup path=filter → division=${division || '(no matching row)'}`);
      return division;
    } catch (err) {
      console.warn(`${LOG} server-side EMail $filter failed — falling back to client-side match`, err);
    }

    // Fallback: fetch rows (capped) and match the viewer's email in-memory.
    const division = await this._clientMatch(base, o, email);
    console.log(`${LOG} lookup path=clientmatch → division=${division || '(no matching row)'}`);
    return division;
  }

  private async _filterByEmail(
    base: string,
    o: IProfileLookupOptions,
    email: string
  ): Promise<string | undefined> {
    const select = encodeURIComponent(`${o.divisionField},${o.personField}/EMail`);
    const filter = encodeURIComponent(`${o.personField}/EMail eq '${email.replace(/'/g, "''")}'`);
    const list = `getByTitle('${o.listTitle.replace(/'/g, "''")}')`;
    const url =
      `${base}/_api/web/lists/${list}/items` +
      `?$expand=${o.personField}&$select=${select}&$filter=${filter}&$top=1`;
    const rows = await this._getRows(url, 'filter');
    if (!rows.length) {
      return undefined;
    }
    return coerceString(rows[0][o.divisionField]) || undefined;
  }

  private async _clientMatch(
    base: string,
    o: IProfileLookupOptions,
    email: string
  ): Promise<string | undefined> {
    const select = encodeURIComponent(`${o.divisionField},${o.personField}/EMail`);
    const list = `getByTitle('${o.listTitle.replace(/'/g, "''")}')`;
    const url =
      `${base}/_api/web/lists/${list}/items` +
      `?$expand=${o.personField}&$select=${select}&$top=${MAX_ITEMS}`;
    const rows = await this._getRows(url, 'clientmatch');
    const lower = email.toLowerCase();
    for (const row of rows) {
      const user = row[o.personField] as IRawUser | undefined;
      const rowEmail = user && typeof user.EMail === 'string' ? user.EMail.toLowerCase() : '';
      if (rowEmail && rowEmail === lower) {
        return coerceString(row[o.divisionField]) || undefined;
      }
    }
    return undefined;
  }

  private async _getRows(url: string, path: string): Promise<IRawRow[]> {
    console.log(`${LOG} (${path}) GET ${url}`);
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(
        `${path} fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
      );
    }
    const json = (await response.json()) as IItemsResponse;
    return json && json.value ? json.value : [];
  }
}
