import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'TabbedListViewsWebPartStrings';
import {
  TabbedListViews,
  ITabbedListViewsProps
} from './components/TabbedListViews';
import { ITabbedListViewsService } from './services/ITabbedListViewsService';
import { TabbedListViewsService } from './services/TabbedListViewsService';
import {
  ITabConfig,
  IListInfo,
  IViewInfo,
  IFieldInfo,
  Layout,
  OverlayPosition
} from './services/models';

export interface ITabbedListViewsWebPartProps {
  sectionTitle: string;
  listId: string;
  layout: Layout;
  tabCount: number;
  tabs: ITabConfig[];
  seeAllUrl: string;
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
}

export default class TabbedListViewsWebPart extends BaseClientSideWebPart<ITabbedListViewsWebPartProps> {
  private _service!: ITabbedListViewsService;
  private _availableLists: IListInfo[] = [];
  private _availableViews: IViewInfo[] = [];
  private _availableFields: IFieldInfo[] = [];
  private _listChoicesLoaded = false;
  private _listChoicesLoading = false;
  private _viewsAndFieldsLoadedFor: string | undefined = undefined;

  protected onInit(): Promise<void> {
    this._service = new TabbedListViewsService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    const props: ITabbedListViewsProps = {
      service: this._service,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      sectionTitle: this.properties.sectionTitle || '',
      listId: this.properties.listId || '',
      layout: this.properties.layout || 'gallery',
      tabCount: this.properties.tabCount || 2,
      tabs: this.properties.tabs || [],
      seeAllUrl: this.properties.seeAllUrl || '',
      showOverlay: !!this.properties.showOverlay,
      overlaySourceField: this.properties.overlaySourceField || '',
      overlayLabelTemplate: this.properties.overlayLabelTemplate || '{value}',
      overlayPosition: this.properties.overlayPosition || 'bottom-left'
    };

    ReactDom.render(React.createElement(TabbedListViews, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // -------------------------------------------------------------------------
  // Property pane data loading
  // -------------------------------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    if (this._listChoicesLoaded || this._listChoicesLoading) {
      // Lists already loaded — kick a views/fields refresh if a list is set
      // and we haven't loaded its dependents yet (e.g. pane reopened).
      this._loadViewsAndFieldsForCurrentList();
      return;
    }
    this._listChoicesLoading = true;

    this._service
      .getLists(this.context.pageContext.web.absoluteUrl)
      .then((lists) => {
        this._availableLists = lists;
      })
      .catch((err: unknown) => {
        // Loud log so a failed lists fetch is visible in DevTools rather than
        // silently leaving the dropdown empty (the 1.0.1.0 failure mode).
        console.error('[TabbedListViews] Failed to load lists', err);
        this._availableLists = [];
      })
      .then(() => {
        this._listChoicesLoaded = true;
        this._listChoicesLoading = false;
        this._loadViewsAndFieldsForCurrentList();
        this.context.propertyPane.refresh();
      })
      .catch(() => {
        // Trailing catch keeps the promise non-floating; pane refresh failure
        // is non-fatal.
      });
  }

  private _loadViewsAndFieldsForCurrentList(): void {
    const listId = this.properties.listId;
    if (!listId) {
      this._availableViews = [];
      this._availableFields = [];
      this._viewsAndFieldsLoadedFor = undefined;
      return;
    }
    if (this._viewsAndFieldsLoadedFor === listId) {
      return;
    }

    // Capture the in-flight listId in the closure so a later list change
    // discards this load's result instead of clobbering newer data.
    const targetListId = listId;
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    Promise.all([
      this._service.getViews(siteUrl, targetListId),
      this._service.getFields(siteUrl, targetListId)
    ])
      .then(([views, fields]) => {
        if (this.properties.listId !== targetListId) {
          return;
        }
        this._availableViews = views;
        this._availableFields = fields;
        this._viewsAndFieldsLoadedFor = targetListId;
        this.context.propertyPane.refresh();
      })
      .catch((err: unknown) => {
        console.warn('[TabbedListViews] Failed to load views/fields', err);
        if (this.properties.listId !== targetListId) {
          return;
        }
        this._availableViews = [];
        this._availableFields = [];
        this._viewsAndFieldsLoadedFor = targetListId;
        this.context.propertyPane.refresh();
      });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      // List changed: clear dependent properties so the editor isn't left with
      // tab views or overlay source from the prior list.
      this.properties.tabs = [];
      this.properties.overlaySourceField = '';
      this._availableViews = [];
      this._availableFields = [];
      this._viewsAndFieldsLoadedFor = undefined;
      this._loadViewsAndFieldsForCurrentList();
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    if (propertyPath === 'layout' && newValue !== 'gallery') {
      // Table layout has no overlay — turn it off so stale overlay config
      // doesn't surface when the user switches back to gallery later.
      this.properties.showOverlay = false;
    }

    this.render();
  }

  // -------------------------------------------------------------------------
  // Property pane configuration
  // -------------------------------------------------------------------------

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listOptions: IPropertyPaneDropdownOption[] = this._availableLists.map((l) => ({
      key: l.id,
      text: l.title
    }));

    const viewOptions: { key: string; text: string }[] = this._availableViews.map((v) => ({
      key: v.id,
      text: v.title
    }));

    const fieldOptions: IPropertyPaneDropdownOption[] = this._availableFields.map((f) => ({
      key: f.internalName,
      text: f.displayName
    }));

    const tabCountOptions: IPropertyPaneDropdownOption[] = [2, 3, 4, 5].map((n) => ({
      key: n,
      text: String(n)
    }));

    const layoutOptions: IPropertyPaneDropdownOption[] = [
      { key: 'gallery', text: strings.LayoutGalleryOptionLabel },
      { key: 'table', text: strings.LayoutTableOptionLabel }
    ];

    const overlayPositionOptions: IPropertyPaneDropdownOption[] = [
      { key: 'top-left', text: strings.OverlayPositionTopLeftLabel },
      { key: 'top-right', text: strings.OverlayPositionTopRightLabel },
      { key: 'bottom-left', text: strings.OverlayPositionBottomLeftLabel },
      { key: 'bottom-right', text: strings.OverlayPositionBottomRightLabel }
    ];

    // Loading indicator while the lists fetch is in flight, and a not-found
    // placeholder option after a successful-but-empty load. Always rendering
    // the dropdown (instead of swapping it with a label) avoids the 1.0.1.0
    // failure mode where a silently-failed fetch left the user with no list
    // picker at all.
    const listDropdownOptions: IPropertyPaneDropdownOption[] = !this._listChoicesLoaded
      ? [{ key: '', text: 'Loading lists…' }]
      : listOptions.length > 0
        ? listOptions
        : [{ key: '', text: '(no lists found on this site)' }];

    const contentFields = [
      PropertyPaneTextField('sectionTitle', {
        label: strings.SectionTitleFieldLabel
      }),
      PropertyPaneDropdown('listId', {
        label: strings.ListFieldLabel,
        options: listDropdownOptions,
        selectedKey: this.properties.listId || undefined,
        disabled: !this._listChoicesLoaded || listOptions.length === 0
      }),
      PropertyPaneDropdown('layout', {
        label: strings.LayoutFieldLabel,
        options: layoutOptions,
        selectedKey: this.properties.layout || 'gallery'
      }),
      PropertyPaneDropdown('tabCount', {
        label: strings.TabCountFieldLabel,
        options: tabCountOptions,
        selectedKey: this.properties.tabCount || 2
      }),
      PropertyFieldCollectionData('tabs', {
        key: 'tabsField',
        label: strings.TabsFieldLabel,
        panelHeader: strings.TabsPanelHeader,
        manageBtnLabel: strings.TabsManageButtonLabel,
        value: this.properties.tabs || [],
        enableSorting: true,
        // Item creation is always allowed so the Add button is visible even
        // before a list is picked — the view dropdown stays empty until then,
        // but the editor isn't trapped on a "No data" screen with no controls.
        fields: [
          {
            id: 'label',
            title: strings.TabLabelColumnLabel,
            type: CustomCollectionFieldType.string,
            required: true
          },
          {
            id: 'viewId',
            title: strings.TabViewColumnLabel,
            type: CustomCollectionFieldType.dropdown,
            required: true,
            options: viewOptions,
            placeholder: strings.TabViewEmptyOptionLabel
          }
        ]
      }),
      PropertyPaneTextField('seeAllUrl', {
        label: strings.SeeAllUrlFieldLabel
      })
    ];

    const overlayFields = [
      PropertyPaneToggle('showOverlay', {
        label: strings.ShowOverlayFieldLabel,
        disabled: this.properties.layout !== 'gallery'
      })
    ];

    if (this.properties.showOverlay && this.properties.layout === 'gallery') {
      overlayFields.push(
        PropertyPaneDropdown('overlaySourceField', {
          label: strings.OverlaySourceFieldLabel,
          options: fieldOptions,
          selectedKey: this.properties.overlaySourceField || undefined
        }),
        PropertyPaneTextField('overlayLabelTemplate', {
          label: strings.OverlayLabelTemplateFieldLabel
        }),
        PropertyPaneDropdown('overlayPosition', {
          label: strings.OverlayPositionFieldLabel,
          options: overlayPositionOptions,
          selectedKey: this.properties.overlayPosition || 'bottom-left'
        })
      );
    }

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ContentGroupName,
              groupFields: contentFields
            },
            {
              groupName: strings.OverlayGroupName,
              groupFields: overlayFields
            }
          ]
        }
      ]
    };
  }
}
