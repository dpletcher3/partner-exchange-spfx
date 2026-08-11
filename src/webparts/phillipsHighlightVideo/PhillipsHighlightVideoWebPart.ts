import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'PhillipsHighlightVideoWebPartStrings';
import { PhillipsHighlightVideo, IPhillipsHighlightVideoProps } from './components/PhillipsHighlightVideo';
import { HighlightVideoService, IListItemRef } from './services/HighlightVideoService';
import { FieldMappingController, IFieldSlot } from '../../shared/fieldMapping';
import { IFieldMapping } from './services/models';

// @pnp/spfx-property-controls ships its own nested copy of
// @microsoft/sp-component-base, so PropertyFieldListPicker's `context` prop type
// isn't structurally assignable from WebPartContext. Cast to the exact type the
// function expects (avoids `any`).
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

const LOG = '[HighlightVideo]';

export interface IPhillipsHighlightVideoWebPartProps {
  // Optional section header above the title (default "Practice of the Month",
  // seeded via the manifest's preconfiguredEntries). Empty = no header.
  sectionHeader: string;
  listId: string;
  // The featured item's ID (0 = none selected).
  itemId: number;
  // Field mapping (convention with override; defaults applied in render()).
  titleField: string;
  videoField: string;
  infoField: string;
}

export default class PhillipsHighlightVideoWebPart extends BaseClientSideWebPart<IPhillipsHighlightVideoWebPartProps> {
  private _service!: HighlightVideoService;

  // Column metadata (field mapping) is owned by the shared FieldMappingController.
  // Item refs (the featured-item picker) are NOT part of that module's scope, so
  // they keep their local loader — both still re-fetch on a list change.
  private _mapping!: FieldMappingController;

  private _availableItems: IListItemRef[] = [];
  private _itemsLoadedFor: string | undefined = undefined;
  private _itemsLoading = false;

  // No typeFilter: the mapping dropdowns have always offered every column, and
  // this migration is behavior-neutral. defaultInternalName documents the §2
  // convention-with-override defaults that render() applies.
  private get _slots(): IFieldSlot[] {
    return [
      { property: 'titleField', label: strings.TitleFieldLabel, defaultInternalName: 'Title' },
      { property: 'videoField', label: strings.VideoFieldLabel, defaultInternalName: 'Video' },
      { property: 'infoField', label: strings.InfoFieldLabel, defaultInternalName: 'HighlightInfo' }
    ];
  }

  protected onInit(): Promise<void> {
    this._service = new HighlightVideoService(this.context.spHttpClient);
    this._mapping = new FieldMappingController({
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      slots: this._slots,
      getListId: () => this.properties.listId || '',
      properties: this.properties as unknown as Record<string, unknown>,
      refresh: () => this.context.propertyPane.refresh(),
      hintLabel: strings.FieldMappingEmptyLabel
    });
    return super.onInit();
  }

  public render(): void {
    // Convention-with-override: apply the §2 defaults so the data layer reads by
    // real internal names even when a mapping property is unset.
    const mapping: IFieldMapping = {
      titleField: this.properties.titleField || 'Title',
      videoField: this.properties.videoField || 'Video',
      infoField: this.properties.infoField || 'HighlightInfo'
    };

    const props: IPhillipsHighlightVideoProps = {
      service: this._service,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listId: this.properties.listId || '',
      itemId: this.properties.itemId || 0,
      mapping,
      // Empty string when unset/cleared so the header is hidden (reusable
      // without one); new instances are seeded from preconfiguredEntries.
      sectionHeader: this.properties.sectionHeader || ''
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
    // columns and items now (no list re-pick required). This is a best-effort
    // first attempt; getPropertyPaneConfiguration self-heals if it races
    // property hydration (listId not yet set when this fires).
    this._mapping.load();
    this._loadItemsForCurrentList();
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
      // Shared module owns the mapping half: it clears slots only when switching
      // between two real lists (a first selection keeps the Title/Video/
      // HighlightInfo defaults), re-fetches the new list's columns, and refreshes.
      this._mapping.onListChanged(oldValue);
      // The item half stays local — re-fetch so the item dropdown repopulates
      // rather than showing the old list's items (the stale-dropdown risk).
      this._availableItems = [];
      this._itemsLoadedFor = undefined;
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

    // Self-heal: getPropertyPaneConfiguration is the one pane method guaranteed to
    // run on every pane render, so kick the loaders here if a persisted list isn't
    // loaded yet. Covers the race where onPropertyPaneConfigurationStart fired
    // before this.properties.listId was hydrated (its !listId early-return would
    // otherwise leave the pane stuck 'loading' with nothing to re-trigger it). Both
    // loaders self-guard (_loadedFor === listId || _loading — the controller does
    // this internally too), so this fires each fetch at most once per listId; each
    // calls propertyPane.refresh() on completion to re-render with populated
    // dropdowns. Retained deliberately: the shared controller is only load()ed from
    // onPropertyPaneConfigurationStart in Celebrations, which has no item picker
    // and so never exercised this race.
    if (listId) {
      this._mapping.load();
      if (this._itemsLoadedFor !== listId && !this._itemsLoading) {
        this._loadItemsForCurrentList();
      }
    }

    const itemsReady = !!listId && this._itemsLoadedFor === listId;
    console.log(
      `${LOG} pane config: listId=${listId || '(none)'}, itemsReady=${itemsReady} (${this._availableItems.length}), ` +
        `columnsReady=${this._mapping.columnsReady}, itemId=${this.properties.itemId || 0}`
    );

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('sectionHeader', {
                  label: strings.SectionHeaderFieldLabel
                }),
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
              groupFields: this._mapping.buildFields()
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

}
