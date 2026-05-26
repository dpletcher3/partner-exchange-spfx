import { INewsRepositoryService } from './INewsRepositoryService';
import { INewsItem, INewsFilters } from './models';
import { ANY_ITEM_TYPE } from '../config/constants';

// Deterministic fixtures for local dev and future tests. Mirrors the shape the
// REST service produces so the component can't tell them apart.
const FIXTURES: INewsItem[] = [
  {
    id: 1,
    title: 'Phillips Opens New Advanced Manufacturing Center',
    categories: ['Company News'],
    itemType: 'Article',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/amc.aspx',
    shortDescription:
      'The new facility expands our precision machining capacity and brings additional capability online for partner programs.',
    publishedDate: '2026-05-22T13:00:00Z'
  },
  {
    id: 2,
    title: 'Partner Spotlight: DMG MORI Collaboration',
    categories: ['Partners'],
    itemType: 'Article',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/dmgmori.aspx',
    shortDescription:
      'A look at the joint engineering work driving faster turnaround on five-axis tooling for our shared customers.',
    publishedDate: '2026-05-20T15:30:00Z'
  },
  {
    id: 3,
    title: 'Federal & Defense Quarterly Briefing Now Available',
    categories: ['Federal & Defense'],
    itemType: 'Announcement',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/fed-q2.aspx',
    shortDescription:
      'Watch the recorded briefing covering program updates, compliance changes, and the FY26 roadmap.',
    publishedDate: '2026-05-18T09:00:00Z'
  },
  {
    id: 4,
    title: 'MyCDT Training Sessions Open for Registration',
    categories: ['Training'],
    itemType: 'Announcement',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/mycdt.aspx',
    shortDescription:
      'New competency development tracks are open. Reserve a seat for the June cohort before slots fill.',
    publishedDate: '2026-05-15T11:45:00Z'
  },
  {
    id: 5,
    title: 'In the News: Phillips Featured in Modern Machine Shop',
    categories: ['Company News', 'Phillips In The News'],
    itemType: 'Press',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/mms.aspx',
    shortDescription:
      'Our automation strategy earned a feature in this month’s issue, highlighting throughput gains across the shop floor.',
    publishedDate: '2026-05-12T08:15:00Z'
  },
  {
    id: 6,
    title: 'Culture Corner: Volunteer Day Recap',
    categories: ['Our Culture'],
    itemType: 'Article',
    linkUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/volunteer.aspx',
    shortDescription:
      'Teams across three sites came together for the spring day of service. Here are the highlights and photos.',
    publishedDate: '2026-05-09T16:20:00Z'
  }
];

export class MockNewsRepositoryService implements INewsRepositoryService {
  public async getCategories(): Promise<string[]> {
    const all: string[] = [];
    for (const item of FIXTURES) {
      all.push(...item.categories);
    }
    return distinct(all);
  }

  public async getItemTypes(): Promise<string[]> {
    return distinct(FIXTURES.map((i) => i.itemType));
  }

  public async getNewsItems(
    _siteUrl: string,
    _listTitle: string,
    filters: INewsFilters,
    maxItems: number
  ): Promise<INewsItem[]> {
    const filtered = FIXTURES.filter((item) => {
      const categoryOk =
        !filters.categories ||
        filters.categories.length === 0 ||
        item.categories.some((c) => filters.categories.indexOf(c) >= 0);
      const itemTypeOk =
        !filters.itemType ||
        filters.itemType === ANY_ITEM_TYPE ||
        filters.itemType === item.itemType;
      return categoryOk && itemTypeOk;
    });

    return filtered
      .sort((a, b) => Date.parse(b.publishedDate) - Date.parse(a.publishedDate))
      .slice(0, maxItems);
  }
}

function distinct(values: string[]): string[] {
  const seen: Record<string, boolean> = {};
  const result: string[] = [];
  for (const v of values) {
    if (v && !seen[v]) {
      seen[v] = true;
      result.push(v);
    }
  }
  return result;
}
