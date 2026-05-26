import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PhillipsNewsWebPartStrings';
import { PhillipsNews, IPhillipsNewsProps } from './components/PhillipsNews';
import { INewsRepositoryService } from './services/INewsRepositoryService';
import { NewsRepositoryService } from './services/NewsRepositoryService';
import { MockNewsRepositoryService } from './services/MockNewsRepositoryService';
import {
  HUB_SITE_URL,
  DEFAULT_LIST_TITLE,
  ANY_ITEM_TYPE,
  USE_MOCK_SERVICE,
  DEFAULT_MAX_ITEMS,
  MIN_MAX_ITEMS,
  MAX_MAX_ITEMS
} from './config/constants';

export interface IPhillipsNewsWebPartProps {
  sectionTitle: string;
  categoryFilter: string[];
  itemTypeFilter: string;
  maxItems: number;
  showViewAllLink: boolean;
  sourceSiteUrl: string;
  listTitle: string;
}

// Synthetic property-path prefix for the dynamic category checkboxes. These are
// not persisted properties — onPropertyPaneFieldChanged intercepts them and
// folds the toggle into the categoryFilter array.
const CATEGORY_FIELD_PREFIX = 'categoryCheckbox_';

export default class PhillipsNewsWebPart extends BaseClientSideWebPart<IPhillipsNewsWebPartProps> {
  private _service!: INewsRepositoryService;
  private _availableCategories: string[] = [];
  private _availableItemTypes: string[] = [];
  private _choicesLoaded = false;
  private _choicesLoading = false;

  protected onInit(): Promise<void> {
    this._service = USE_MOCK_SERVICE
      ? new MockNewsRepositoryService()
      : new NewsRepositoryService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IPhillipsNewsProps> = React.createElement(PhillipsNews, {
      service: this._service,
      sectionTitle: this.properties.sectionTitle || '',
      categoryFilter: this.properties.categoryFilter || [],
      itemTypeFilter: this.properties.itemTypeFilter || ANY_ITEM_TYPE,
      maxItems: this.properties.maxItems || DEFAULT_MAX_ITEMS,
      showViewAllLink: this.properties.showViewAllLink !== false,
      sourceSiteUrl: this._resolvedSiteUrl,
      listTitle: this._resolvedListTitle
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Load category + item-type choices from the list when the property pane opens.
  protected onPropertyPaneConfigurationStart(): void {
    if (this._choicesLoaded || this._choicesLoading) {
      return;
    }
    this._choicesLoading = true;

    Promise.all([
      this._service.getCategories(this._resolvedSiteUrl, this._resolvedListTitle),
      this._service.getItemTypes(this._resolvedSiteUrl, this._resolvedListTitle)
    ])
      .then(
        ([categories, itemTypes]) => {
          this._availableCategories = categories;
          this._availableItemTypes = itemTypes;
        },
        (err: unknown) => {
          // Choices unavailable (e.g. list not found): leave empty, pane still
          // works. Logged so an unexpected failure isn't silent in the console.
          console.warn('[PhillipsNews] Failed to load property-pane choices', err);
          this._availableCategories = [];
          this._availableItemTypes = [];
        }
      )
      .then(() => {
        this._choicesLoaded = true;
        this._choicesLoading = false;
        this.context.propertyPane.refresh();
      })
      .catch(() => {
        // Pane refresh failure is non-fatal; swallow so the promise isn't floating.
      });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (propertyPath.indexOf(CATEGORY_FIELD_PREFIX) === 0) {
      const index = parseInt(propertyPath.substring(CATEGORY_FIELD_PREFIX.length), 10);
      const category = this._availableCategories[index];
      if (category) {
        const selected = new Set<string>(this.properties.categoryFilter || []);
        if (newValue === true) {
          selected.add(category);
        } else {
          selected.delete(category);
        }
        this.properties.categoryFilter = Array.from(selected);
      }
      // Drop the synthetic checkbox property so it doesn't persist in the bag.
      delete (this.properties as unknown as Record<string, unknown>)[propertyPath];
      this.render();
      return;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const itemTypeOptions: IPropertyPaneDropdownOption[] = [
      { key: ANY_ITEM_TYPE, text: ANY_ITEM_TYPE },
      ...this._availableItemTypes.map((t) => ({ key: t, text: t }))
    ];

    const selectedCategories = this.properties.categoryFilter || [];
    const categoryFields =
      this._availableCategories.length > 0
        ? this._availableCategories.map((category, index) =>
            PropertyPaneCheckbox(`${CATEGORY_FIELD_PREFIX}${index}`, {
              text: category,
              checked: selectedCategories.indexOf(category) >= 0
            })
          )
        : [PropertyPaneLabel('categoryFilterEmpty', { text: strings.CategoryFilterEmptyLabel })];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('sectionTitle', {
                  label: strings.SectionTitleFieldLabel
                }),
                PropertyPaneLabel('categoryFilterLabel', {
                  text: strings.CategoryFilterFieldLabel
                }),
                ...categoryFields,
                PropertyPaneDropdown('itemTypeFilter', {
                  label: strings.ItemTypeFilterFieldLabel,
                  options: itemTypeOptions,
                  selectedKey: this.properties.itemTypeFilter || ANY_ITEM_TYPE
                }),
                PropertyPaneSlider('maxItems', {
                  label: strings.MaxItemsFieldLabel,
                  min: MIN_MAX_ITEMS,
                  max: MAX_MAX_ITEMS,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle('showViewAllLink', {
                  label: strings.ShowViewAllLinkFieldLabel
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('sourceSiteUrl', {
                  label: strings.SourceSiteUrlFieldLabel
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private get _resolvedSiteUrl(): string {
    const configured = this.properties.sourceSiteUrl;
    return configured && configured.trim() ? configured.trim() : HUB_SITE_URL;
  }

  private get _resolvedListTitle(): string {
    const configured = this.properties.listTitle;
    return configured && configured.trim() ? configured.trim() : DEFAULT_LIST_TITLE;
  }
}
