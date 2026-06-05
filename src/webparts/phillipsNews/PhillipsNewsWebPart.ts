import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PhillipsNewsWebPartStrings';
import { PhillipsNews, IPhillipsNewsProps } from './components/PhillipsNews';
import { INewsRepositoryService } from './services/INewsRepositoryService';
import { NewsRepositoryService } from './services/NewsRepositoryService';
import { NewsPipelineRepositoryService } from './services/NewsPipelineRepositoryService';
import { MockNewsRepositoryService } from './services/MockNewsRepositoryService';
import {
  HUB_SITE_URL,
  DEFAULT_LIST_TITLE,
  ANY_ITEM_TYPE,
  USE_MOCK_SERVICE,
  DEFAULT_MAX_ITEMS,
  MIN_MAX_ITEMS,
  MAX_MAX_ITEMS,
  DataSource,
  DEFAULT_DATA_SOURCE
} from './config/constants';

export interface IPhillipsNewsWebPartProps {
  dataSource: DataSource;
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
    this._service = this._createService(this._resolvedDataSource);
    return super.onInit();
  }

  // Chooses the data-source implementation behind the shared
  // INewsRepositoryService interface — same injection seam as the existing
  // USE_MOCK_SERVICE switch. The mock takes precedence for local dev; otherwise
  // the dataSource property selects list vs news-pipeline.
  private _createService(dataSource: DataSource): INewsRepositoryService {
    if (USE_MOCK_SERVICE) {
      return new MockNewsRepositoryService();
    }
    return dataSource === 'pipeline'
      ? new NewsPipelineRepositoryService(this.context.spHttpClient)
      : new NewsRepositoryService(this.context.spHttpClient);
  }

  public render(): void {
    const element: React.ReactElement<IPhillipsNewsProps> = React.createElement(PhillipsNews, {
      service: this._service,
      dataSource: this._resolvedDataSource,
      sectionTitle: this.properties.sectionTitle || '',
      categoryFilter: this.properties.categoryFilter || [],
      itemTypeFilter: this.properties.itemTypeFilter || ANY_ITEM_TYPE,
      maxItems: this.properties.maxItems || DEFAULT_MAX_ITEMS,
      showViewAllLink: this.properties.showViewAllLink !== false,
      sourceSiteUrl: this._resolvedSiteUrl,
      listTitle: this._resolvedListTitle,
      isEditMode: this.displayMode === DisplayMode.Edit
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Load category + item-type choices from the active source when the pane
  // opens. The choices come from whichever service is currently injected, so
  // the source of the data (News Repository list vs Site Pages news column)
  // follows the dataSource toggle.
  protected onPropertyPaneConfigurationStart(): void {
    this._loadChoices();
  }

  private _loadChoices(): void {
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
          // Choices unavailable (e.g. list/library not found): leave empty,
          // pane still works. Logged so an unexpected failure isn't silent.
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

    if (propertyPath === 'dataSource') {
      // Re-inject the service so subsequent reads hit the new source.
      this._service = this._createService(newValue as DataSource);

      // The two sources expose different category/item-type vocabularies, so
      // any prior selections would silently mis-filter. Clear them and reload
      // the choice lists from the new source, then refresh the pane.
      this.properties.categoryFilter = [];
      this.properties.itemTypeFilter = ANY_ITEM_TYPE;
      this._availableCategories = [];
      this._availableItemTypes = [];
      this._choicesLoaded = false;
      this._choicesLoading = false;
      this._loadChoices();

      this.render();
    }
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

    // The listTitle field is meaningless in pipeline mode (the pipeline always
    // reads the Site Pages library); hide it there. sourceSiteUrl stays in both
    // modes — the Loop instance will point pipeline mode at Our Culture.
    const isPipeline = this._resolvedDataSource === 'pipeline';
    const advancedFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('sourceSiteUrl', {
        label: strings.SourceSiteUrlFieldLabel
      })
    ];
    if (!isPipeline) {
      advancedFields.push(
        PropertyPaneTextField('listTitle', {
          label: strings.ListTitleFieldLabel
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
              groupFields: [
                // The data-source choice sits at the very top of the Content
                // group: it changes the meaning of every field below it.
                PropertyPaneChoiceGroup('dataSource', {
                  label: strings.DataSourceFieldLabel,
                  options: [
                    { key: 'list', text: strings.DataSourceListLabel },
                    { key: 'pipeline', text: strings.DataSourcePipelineLabel }
                  ]
                }),
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
              groupFields: advancedFields
            }
          ]
        }
      ]
    };
  }

  private get _resolvedDataSource(): DataSource {
    return this.properties.dataSource || DEFAULT_DATA_SOURCE;
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
