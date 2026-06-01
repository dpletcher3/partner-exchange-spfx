import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'PhillipsHighlightVideoWebPartStrings';
import { PhillipsHighlightVideo, IPhillipsHighlightVideoProps } from './components/PhillipsHighlightVideo';
import { HighlightVideoService, IColumnInfo, IListItemRef } from './services/HighlightVideoService';

// @pnp/spfx-property-controls ships its own nested copy of
// @microsoft/sp-component-base, so PropertyFieldListPicker's `context` prop type
// isn't structurally assignable from WebPartContext. Cast to the exact type the
// function expects (avoids `any`).
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

const LOG = '[HighlightVideo]';

export interface IPhillipsHighlightVideoWebPartProps {
  listId: string;
  // The featured item's ID (0 = none selected).
  itemId: number;
  // Field mapping (convention with override; defaults applied in render()).
  titleField: string;
  videoField: string;
  infoField: string;
}

type MappingProperty = 'titleField' | 'videoField' | 'infoField';

export default class PhillipsHighlightVideoWebPart extends BaseClientSideWebPart<IPhillipsHighlightVideoWebPartProps> {
  private _service!: HighlightVideoService;

  // Both column metadata (for field mapping) and item refs (for the item
  // picker) depend on the selected list; both re-fetch on a list change.
  private _availableColumns: IColumnInfo[] = [];
  private _columnsLoadedFor: string | undefined = undefined;
  private _columnsLoading = false;

  private _availableItems: IListItemRef[] = [];
  private _itemsLoadedFor: string | undefined = undefined;
  private _itemsLoading = false;

  protected onInit(): Promise<void> {
    this._service = new HighlightVideoService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    // Turn 1 is scaffold only — no item read or Vimeo embed yet (Turn 2). The
    // component shows the configure prompt until a list AND item are chosen,
    // then a placeholder shell.
    const props: IPhillipsHighlightVideoProps = {
      hasList: !!this.properties.listId,
      hasItem: !!this.properties.itemId
    };
    ReactDom.render(React.createElement(PhillipsHighlightVideo, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // -------------------------------------------------------------------------
  // Dependent property-pane loading (columns + items, both keyed to the list)
  // -------------------------------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    // Initial-open path: a list may already be persisted, so load both its
    // columns and items now (no list re-pick required).
    this._loadColumnsForCurrentList();
    this._loadItemsForCurrentList();
  }

  private _loadColumnsForCurrentList(): void {
    const listId = this.properties.listId;
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    if (!listId) {
      this._availableColumns = [];
      this._columnsLoadedFor = undefined;
      return;
    }
    if (this._columnsLoadedFor === listId || this._columnsLoading) {
      return;
    }
    this._columnsLoading = true;
    console.log(`${LOG} loading columns for listId=${listId}`);
    const target = listId;
    this._service
      .getColumns(siteUrl, target)
      .then((cols) => {
        if (this.properties.listId !== target) {
          return;
        }
        this._availableColumns = cols;
        this._columnsLoadedFor = target;
        console.log(`${LOG} loaded ${cols.length} columns:`, cols.map((c) => c.internalName));
        this.context.propertyPane.refresh();
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} column fetch FAILED for listId=${target}`, err);
        if (this.properties.listId !== target) {
          return;
        }
        this._availableColumns = [];
        this._columnsLoadedFor = target;
        this.context.propertyPane.refresh();
      })
      .then(() => {
        this._columnsLoading = false;
      })
      .catch(() => {
        /* non-floating */
      });
  }

  private _loadItemsForCurrentList(): void {
    const listId = this.properties.listId;
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    if (!listId) {
      this._availableItems = [];
      this._itemsLoadedFor = undefined;
      return;
    }
    if (this._itemsLoadedFor === listId || this._itemsLoading) {
      return;
    }
    this._itemsLoading = true;
    console.log(`${LOG} loading items for listId=${listId}`);
    const target = listId;
    this._service
      .getListItems(siteUrl, target)
      .then((items) => {
        if (this.properties.listId !== target) {
          return;
        }
        this._availableItems = items;
        this._itemsLoadedFor = target;
        console.log(`${LOG} loaded ${items.length} items for listId=${target}`);
        this.context.propertyPane.refresh();
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} item fetch FAILED for listId=${target}`, err);
        if (this.properties.listId !== target) {
          return;
        }
        this._availableItems = [];
        this._itemsLoadedFor = target;
        this.context.propertyPane.refresh();
      })
      .then(() => {
        this._itemsLoading = false;
      })
      .catch(() => {
        /* non-floating */
      });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      console.log(`${LOG} listId changed: ${String(oldValue) || '(none)'} → ${String(newValue) || '(none)'}`);
      // The previously-featured item belongs to the old list — always clear it.
      this.properties.itemId = 0;
      // Convention-with-override: only clear mappings when switching BETWEEN two
      // real lists; the first selection keeps the defaults (Title/Video/HighlightInfo).
      if (oldValue) {
        console.log(`${LOG} switching lists — clearing item + stale column mappings`);
        this.properties.titleField = '';
        this.properties.videoField = '';
        this.properties.infoField = '';
      } else {
        console.log(`${LOG} first list selection — preserving default mappings`);
      }
      // Re-fetch BOTH the item list and the columns, then refresh so both the
      // item dropdown and the mapping dropdowns repopulate (the stale-dropdown risk).
      this._availableColumns = [];
      this._columnsLoadedFor = undefined;
      this._availableItems = [];
      this._itemsLoadedFor = undefined;
      this._loadColumnsForCurrentList();
      this._loadItemsForCurrentList();
      this.context.propertyPane.refresh();
    }

    this.render();
  }

  // -------------------------------------------------------------------------
  // Property-pane configuration
  // -------------------------------------------------------------------------

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listId = this.properties.listId || '';
    const columnsReady = !!listId && this._columnsLoadedFor === listId;
    const itemsReady = !!listId && this._itemsLoadedFor === listId;
    const mappingBranch = !listId ? 'no-list (hint)' : columnsReady ? 'populated' : 'loading';
    console.log(
      `${LOG} pane config: listId=${listId || '(none)'}, itemsReady=${itemsReady} (${this._availableItems.length}), ` +
        `columnsReady=${columnsReady} (${this._availableColumns.length}), field-mapping branch=${mappingBranch}, ` +
        `itemId=${this.properties.itemId || 0}`
    );

    const mappingFields: IPropertyPaneField<unknown>[] = listId
      ? [
          this._mappingDropdown('titleField', strings.TitleFieldLabel, columnsReady),
          this._mappingDropdown('videoField', strings.VideoFieldLabel, columnsReady),
          this._mappingDropdown('infoField', strings.InfoFieldLabel, columnsReady)
        ]
      : [PropertyPaneLabel('fieldMappingHint', { text: strings.FieldMappingEmptyLabel })];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: strings.ListFieldLabel,
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as unknown as PnpContext,
                  deferredValidationTime: 0,
                  key: 'highlightVideoListPicker'
                }),
                this._itemDropdown(listId, itemsReady)
              ]
            },
            {
              groupName: strings.FieldMappingGroupName,
              // Expanded (not collapsed): a collapsed group renders header-only,
              // which reads as "no dropdowns." Matches the Media Card Gallery.
              isCollapsed: false,
              groupFields: mappingFields
            }
          ]
        }
      ]
    };
  }

  // Item picker: lists the selected list's items by Title, stores the item ID.
  private _itemDropdown(listId: string, itemsReady: boolean): IPropertyPaneField<unknown> {
    let options: IPropertyPaneDropdownOption[];
    if (!listId) {
      options = [{ key: 0, text: 'Select a list first' }];
    } else if (!itemsReady) {
      options = [{ key: 0, text: 'Loading items…' }];
    } else if (this._availableItems.length === 0) {
      options = [{ key: 0, text: '(no items in this list)' }];
    } else {
      options = this._availableItems.map((i) => ({
        key: i.id,
        text: i.title || `(untitled — item ${i.id})`
      }));
    }

    return PropertyPaneDropdown('itemId', {
      label: strings.ItemFieldLabel,
      options,
      selectedKey: this.properties.itemId || undefined,
      disabled: !itemsReady || this._availableItems.length === 0
    });
  }

  // A field-mapping dropdown populated from the selected list's columns.
  private _mappingDropdown(
    targetProperty: MappingProperty,
    label: string,
    columnsReady: boolean
  ): IPropertyPaneField<unknown> {
    const columnOptions: IPropertyPaneDropdownOption[] = this._availableColumns.map((c) => ({
      key: c.internalName,
      text: c.displayName
    }));

    let options: IPropertyPaneDropdownOption[];
    if (!columnsReady) {
      options = [{ key: '', text: 'Loading columns…' }];
    } else if (columnOptions.length === 0) {
      options = [{ key: '', text: '(no columns found on this list)' }];
    } else {
      options = columnOptions;
    }

    return PropertyPaneDropdown(targetProperty, {
      label,
      options,
      selectedKey: (this.properties[targetProperty] as string) || undefined,
      disabled: !columnsReady || columnOptions.length === 0
    });
  }
}
