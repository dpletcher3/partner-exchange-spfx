import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LOG = '[Celebrations]';
const MAX_ITEMS = 500;

export interface ICelebrationsMapping {
  personField: string;
  birthdayField: string;
  hireField: string;
}

// A person normalized for the calendar: the M365 photo URL (from the Person
// field, D030 pattern) + the raw ISO birth/hire dates (the pure date module
// does the recurrence math).
export interface ICelebrationPerson {
  id: number;
  name: string;
  photoUrl?: string;
  birthDate?: string;
  hireDate?: string;
}

interface IRawUser {
  Title?: string;
  EMail?: string;
}
interface IRawRow {
  Id: number;
  Title?: string;
  [key: string]: unknown;
}
interface IItemsResponse {
  value: IRawRow[];
}

export class CelebrationsDataService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  // Reads people flagged for celebrations. Photo derives from the Person field
  // (LinkedUser → userphoto.aspx), NOT an image column.
  public async getPeople(
    siteUrl: string,
    listId: string,
    mapping: ICelebrationsMapping
  ): Promise<ICelebrationPerson[]> {
    const base = siteUrl.replace(/\/+$/, '');
    const { personField, birthdayField, hireField } = mapping;

    const select = [
      'Id',
      'Title',
      birthdayField,
      hireField,
      `${personField}/Id`,
      `${personField}/Title`,
      `${personField}/EMail`
    ].join(',');
    // ShowInCelebrations + IsActive are Partner Profiles Yes/No fields (spec §2).
    const filter = encodeURIComponent('ShowInCelebrations eq 1 and IsActive eq 1');
    const url =
      `${base}/_api/web/lists(guid'${listId}')/items` +
      `?$select=${select}&$expand=${personField}&$filter=${filter}&$top=${MAX_ITEMS}`;
    console.log(`${LOG} getPeople URL: ${url}`);

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(
        `celebrations fetch failed (${response.status} ${response.statusText}) — ${body.slice(0, 200)}`
      );
    }

    const json = (await response.json()) as IItemsResponse;
    const rows = json && json.value ? json.value : [];
    const people = rows.map((row) => this._map(row, mapping, base));
    console.log(`${LOG} filtered ${people.length} people (ShowInCelebrations + IsActive)`);
    return people;
  }

  private _map(row: IRawRow, mapping: ICelebrationsMapping, base: string): ICelebrationPerson {
    const user = row[mapping.personField] as IRawUser | undefined;
    const name = (user && user.Title) || row.Title || '';
    const email = user && typeof user.EMail === 'string' ? user.EMail : '';
    const photoUrl = email
      ? `${base}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`
      : undefined;

    const rawBirth = row[mapping.birthdayField];
    const rawHire = row[mapping.hireField];
    return {
      id: row.Id,
      name,
      photoUrl,
      birthDate: typeof rawBirth === 'string' ? rawBirth : undefined,
      hireDate: typeof rawHire === 'string' ? rawHire : undefined
    };
  }
}
