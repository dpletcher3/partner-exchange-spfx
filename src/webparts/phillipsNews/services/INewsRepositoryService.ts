import { INewsItem, INewsFilters } from './models';

// Data-source contract for the Phillips News web part. The component depends on
// this interface, not on a concrete implementation — NewsRepositoryService hits
// SharePoint REST, MockNewsRepositoryService returns fixtures. The web part
// chooses which to inject in onInit.
export interface INewsRepositoryService {
  // Distinct choices from the list's Category column, for the property pane.
  getCategories(siteUrl: string, listTitle: string): Promise<string[]>;

  // Distinct choices from the list's ItemType column, for the property pane.
  getItemTypes(siteUrl: string, listTitle: string): Promise<string[]>;

  // News items matching the filters, newest first, capped at maxItems.
  getNewsItems(
    siteUrl: string,
    listTitle: string,
    filters: INewsFilters,
    maxItems: number
  ): Promise<INewsItem[]>;
}
