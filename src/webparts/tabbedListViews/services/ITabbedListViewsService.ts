import { IListInfo, IViewInfo, IFieldInfo, ITabData } from './models';

// Data-source contract for the Tabbed List Views web part. The React component
// depends on this interface, not on a concrete implementation, so the property
// pane / runtime can swap in a mock for local development if needed later.
export interface ITabbedListViewsService {
  // Visible, user-created lists on the current site (filtered to BaseTemplate
  // 100 — Generic List). Used by the property pane's list dropdown.
  getLists(siteUrl: string): Promise<IListInfo[]>;

  // Visible, non-personal views on the selected list. Used by the per-tab view
  // dropdown.
  getViews(siteUrl: string, listId: string): Promise<IViewInfo[]>;

  // Visible fields on the selected list, used for the overlay-source dropdown
  // and to resolve table-column display names.
  getFields(siteUrl: string, listId: string): Promise<IFieldInfo[]>;

  // Items for a tab: runs the view's CAML query via RenderListDataAsStream so
  // filter, sort, and row-limit are honored. `extraFields` adds any fields the
  // caller needs that aren't already in the view (e.g. the overlay source).
  getTabData(
    siteUrl: string,
    listId: string,
    viewId: string,
    extraFields: string[]
  ): Promise<ITabData>;
}
