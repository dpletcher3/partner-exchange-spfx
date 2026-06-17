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
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'PhillipsDocumentCardsWebPartStrings';
import { PhillipsDocumentCards, IPhillipsDocumentCardsProps } from './components/PhillipsDocumentCards';
import { DocumentCardsFieldService, IColumnInfo } from './services/DocumentCardsFieldService';
import { PhillipsDocumentCardsService } from './services/PhillipsDocumentCardsService';
import { IFieldMapping, IDocColumnConfig } from './services/models';

// @pnp/spfx-property-controls ships its own nested copy of
// @microsoft/sp-component-base, so PropertyFieldListPicker's `context` prop type
// isn't structurally assignable from WebPartContext. Cast to the exact type the
// function expects (avoids `any`, stays correct across package bumps).
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

const LOG = '[DocumentCards]';

// Document library base template (101) — scopes the list picker to libraries.
const DOCUMENT_LIBRARY_TEMPLATE = 101;
// Spec caps the per-column config at 4 rows.
const MAX_COLUMNS = 4;

// Default field mappings (I16 Stage 1 schema). The data layer falls back to
// these when the property is unset, so a freshly-added instance pointed at a
// Phillips Documents library needs nothing remapped.
const DEFAULT_TITLE_FIELD = 'FileLeafRef';
const DEFAULT_DESCRIPTION_FIELD = 'CardDescription';
const DEFAULT_ICON_FIELD = 'CardIcon';
const DEFAULT_SECTION_FIELD = 'DocSection';

export interface IPhillipsDocumentCardsWebPartProps {
  listId: string;
  // Field mapping (convention with override) — set once, applies to all columns.
  titleField: string;
  descriptionField: string;
  iconField: string;
  sectionField: string;
  // Per-column config (one row per on-screen column, max 4).
  columns: IDocColumnConfig[];
}

type MappingProperty = 'titleField' | 'descriptionField' | 'iconField' | 'sectionField';

export default class PhillipsDocumentCardsWebPart extends BaseClientSideWebPart<IPhillipsDocumentCardsWebPartProps> {
  private _fieldService!: DocumentCardsFieldService;
  private _dataService!: PhillipsDocumentCardsService;
  private _availableColumns: IColumnInfo[] = [];
  private _columnsLoadedFor: string | undefined = undefined;
  private _columnsLoading = false;

  protected onInit(): Promise<void> {
    this._fieldService = new DocumentCardsFieldService(this.context.spHttpClient);
    this._dataService = new PhillipsDocumentCardsService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    // Convention-with-override: apply the §1 defaults here, AND drop any stored
    // value that resolves to a SharePoint system column, so the data layer reads
    // by real internal names even when a mapping property is unset or stale.
    const mapping: IFieldMapping = this._resolveMapping();

    const props: IPhillipsDocumentCardsProps = {
      service: this._dataService,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listId: this.properties.listId || '',
      mapping,
      columns: this.properties.columns || []
    };

    ReactDom.render(React.createElement(PhillipsDocumentCards, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Resolve the stored mapping to safe internal names. A stored value that is
  // empty OR a SharePoint system column (internal name starting with '_') falls
  // back to the role default. This guards the built-in document-library
  // "Description" column (_ExtendedDescription): it shares our CardDescription's
  // display name, so an instance that captured it from the (formerly ambiguous)
  // dropdown would 400 the items $select. With this guard such a stale instance
  // still fetches correctly by CardDescription, no property-pane re-pick needed.
  private _resolveMapping(): IFieldMapping {
    const pick = (stored: string | undefined, fallback: string): string => {
      const v = (stored || '').trim();
      return !v || v.charAt(0) === '_' ? fallback : v;
    };
    return {
      titleField: pick(this.properties.titleField, DEFAULT_TITLE_FIELD),
      descriptionField: pick(this.properties.descriptionField, DEFAULT_DESCRIPTION_FIELD),
      iconField: pick(this.properties.iconField, DEFAULT_ICON_FIELD),
      sectionField: pick(this.properties.sectionField, DEFAULT_SECTION_FIELD)
    };
  }

  // -------------------------------------------------------------------------
  // Property-pane column loading (dependent on the selected library)
  // -------------------------------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    // Initial-open path: a library may already be persisted. Load its columns
    // now so the mapping dropdowns populate without requiring a re-pick.
    this._loadColumnsForCurrentList();
  }

  private _loadColumnsForCurrentList(): void {
    const listId = this.properties.listId;
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    if (!listId) {
      console.log(`${LOG} no library selected — skipping column load`);
      this._availableColumns = [];
      this._columnsLoadedFor = undefined;
      return;
    }
    if (this._columnsLoadedFor === listId || this._columnsLoading) {
      console.log(`${LOG} columns already loaded/loading for listId=${listId}`);
      return;
    }

    this._columnsLoading = true;
    console.log(`${LOG} loading columns for listId=${listId} from ${siteUrl}`);

    const targetListId = listId;
    this._fieldService
      .getColumns(siteUrl, targetListId)
      .then((cols) => {
        if (this.properties.listId !== targetListId) {
          console.log(`${LOG} library changed while loading (${targetListId} → ${this.properties.listId}); discarding`);
          return;
        }
        this._availableColumns = cols;
        this._columnsLoadedFor = targetListId;
        console.log(`${LOG} loaded ${cols.length} columns for listId=${targetListId}:`, cols.map((c) => c.internalName));
        this.context.propertyPane.refresh();
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} column fetch FAILED for listId=${targetListId}`, err);
        if (this.properties.listId !== targetListId) {
          return;
        }
        this._availableColumns = [];
        this._columnsLoadedFor = targetListId; // mark resolved so we show "(no columns)" not a spinner
        this.context.propertyPane.refresh();
      })
      .then(() => {
        this._columnsLoading = false;
      })
      .catch(() => {
        /* keep the promise non-floating */
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      console.log(`${LOG} listId changed: ${String(oldValue) || '(none)'} → ${String(newValue) || '(none)'}`);
      // Reload the new library's columns so the mapping dropdowns repopulate.
      // The mapping defaults are identical across Phillips Documents libraries,
      // so stored mappings are PRESERVED across a library switch (no clearing).
      this._availableColumns = [];
      this._columnsLoadedFor = undefined;
      this._loadColumnsForCurrentList();
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
    const branch = !listId ? 'no-list' : columnsReady ? 'populated' : 'loading';
    console.log(
      `${LOG} pane config: listId=${listId || '(none)'}, columnsLoadedFor=${this._columnsLoadedFor || '(none)'}, ` +
        `availableColumns=${this._availableColumns.length}, field-mapping branch=${branch}`
    );

    const mappingFields: IPropertyPaneField<unknown>[] = listId
      ? [
          this._mappingDropdown('titleField', strings.TitleFieldLabel, columnsReady),
          this._mappingDropdown('descriptionField', strings.DescriptionFieldLabel, columnsReady),
          this._mappingDropdown('iconField', strings.IconFieldLabel, columnsReady),
          this._mappingDropdown('sectionField', strings.SectionFieldLabel, columnsReady)
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
                  label: strings.LibraryFieldLabel,
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  baseTemplate: DOCUMENT_LIBRARY_TEMPLATE,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as unknown as PnpContext,
                  deferredValidationTime: 0,
                  key: 'documentCardsLibraryPicker'
                })
              ]
            },
            {
              groupName: strings.FieldMappingGroupName,
              isCollapsed: false,
              groupFields: mappingFields
            },
            {
              groupName: strings.ColumnsGroupName,
              groupFields: [
                PropertyFieldCollectionData('columns', {
                  key: 'documentCardsColumns',
                  label: strings.ColumnsFieldLabel,
                  panelHeader: strings.ColumnsPanelHeader,
                  manageBtnLabel: strings.ColumnsManageButtonLabel,
                  value: this.properties.columns || [],
                  enableSorting: true,
                  // Cap at 4 columns — disable the Add row once at the limit.
                  disableItemCreation: (this.properties.columns || []).length >= MAX_COLUMNS,
                  fields: [
                    {
                      id: 'header',
                      title: strings.ColumnHeaderLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'filterValue',
                      title: strings.ColumnFilterValueLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'color',
                      title: strings.ColumnColorLabel,
                      type: CustomCollectionFieldType.color
                    },
                    {
                      id: 'iconName',
                      title: strings.ColumnIconNameLabel,
                      type: CustomCollectionFieldType.string
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // A field-mapping dropdown populated from the selected library's columns.
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
      options = [{ key: '', text: '(no columns found on this library)' }];
    } else {
      options = columnOptions;
    }

    const stored = (this.properties[targetProperty] as string) || '';
    return PropertyPaneDropdown(targetProperty, {
      label,
      options,
      selectedKey: stored || undefined,
      disabled: !columnsReady || columnOptions.length === 0
    });
  }
}
