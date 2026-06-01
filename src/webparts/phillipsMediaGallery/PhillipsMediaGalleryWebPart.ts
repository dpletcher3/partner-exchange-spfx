import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'PhillipsMediaGalleryWebPartStrings';
import { PhillipsMediaGallery, IPhillipsMediaGalleryProps } from './components/PhillipsMediaGallery';
import { MediaGalleryFieldService, IColumnInfo } from './services/MediaGalleryFieldService';
import { PhillipsMediaGalleryService } from './services/PhillipsMediaGalleryService';
import { IFieldMapping } from './services/models';

// @pnp/spfx-property-controls ships its own nested copy of
// @microsoft/sp-component-base, so PropertyFieldListPicker's `context` prop type
// isn't structurally assignable from WebPartContext. Cast to the exact type the
// function expects (avoids `any`, stays correct across package bumps).
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

const LOG = '[MediaGallery]';

export interface IPhillipsMediaGalleryWebPartProps {
  listId: string;
  columns: number;
  sectionTitle: string;
  openInNewTab: boolean;
  // Field mapping (convention with override). Each defaults to the §2 internal
  // name; the data layer (Turn 2) falls back to these defaults when a value is
  // unset, so the 15 Practices instance works with nothing remapped here.
  titleField: string;
  videoField: string;
  labelImageField: string;
  mainImageField: string;
}

const MIN_COLUMNS = 3;
const MAX_COLUMNS = 5;
const DEFAULT_COLUMNS = 4;
const NONE_KEY = '__none__';

type MappingProperty = 'titleField' | 'videoField' | 'labelImageField' | 'mainImageField';

export default class PhillipsMediaGalleryWebPart extends BaseClientSideWebPart<IPhillipsMediaGalleryWebPartProps> {
  private _fieldService!: MediaGalleryFieldService;
  private _dataService!: PhillipsMediaGalleryService;
  private _availableColumns: IColumnInfo[] = [];
  private _columnsLoadedFor: string | undefined = undefined;
  private _columnsLoading = false;

  protected onInit(): Promise<void> {
    this._fieldService = new MediaGalleryFieldService(this.context.spHttpClient);
    this._dataService = new PhillipsMediaGalleryService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    // Convention-with-override: apply the §2 defaults here so the data layer
    // reads by real internal names even when the property is unset.
    const mapping: IFieldMapping = {
      titleField: this.properties.titleField || 'Title',
      videoField: this.properties.videoField || 'Video',
      labelImageField: this.properties.labelImageField || 'Image0',
      mainImageField: this.properties.mainImageField || ''
    };

    const props: IPhillipsMediaGalleryProps = {
      service: this._dataService,
      httpClient: this.context.httpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listId: this.properties.listId || '',
      mapping,
      columns: this._resolvedColumns,
      sectionTitle: this.properties.sectionTitle || '',
      openInNewTab: this.properties.openInNewTab !== false
    };

    ReactDom.render(React.createElement(PhillipsMediaGallery, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private get _resolvedColumns(): number {
    const c = this.properties.columns || DEFAULT_COLUMNS;
    return Math.min(MAX_COLUMNS, Math.max(MIN_COLUMNS, c));
  }

  // -------------------------------------------------------------------------
  // Property-pane column loading (dependent on the selected list)
  // -------------------------------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    // Initial-open path: a list may already be persisted (the reported bug —
    // pane opens with 15 Practices selected). Load its columns now so the
    // mapping dropdowns populate without requiring a list re-pick.
    this._loadColumnsForCurrentList();
  }

  private _loadColumnsForCurrentList(): void {
    const listId = this.properties.listId;
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    if (!listId) {
      console.log(`${LOG} no list selected — skipping column load`);
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

    // Capture the in-flight listId so a later list change discards this result.
    const targetListId = listId;
    this._fieldService
      .getColumns(siteUrl, targetListId)
      .then((cols) => {
        if (this.properties.listId !== targetListId) {
          console.log(`${LOG} list changed while loading (${targetListId} → ${this.properties.listId}); discarding`);
          return;
        }
        this._availableColumns = cols;
        this._columnsLoadedFor = targetListId;
        console.log(
          `${LOG} loaded ${cols.length} columns for listId=${targetListId}:`,
          cols.map((c) => c.internalName)
        );
        this.context.propertyPane.refresh();
      })
      .catch((err: unknown) => {
        // Loud, unlike the @pnp control's silent .catch(() => {}).
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

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    // The optional main-image "(none)" choice maps to an unset value so the data
    // layer (Turn 2) treats it as "auto-derive from video", consistent with the
    // unset default.
    if (propertyPath === 'mainImageField' && newValue === NONE_KEY) {
      this.properties.mainImageField = '';
    }

    if (propertyPath === 'listId' && oldValue !== newValue) {
      console.log(`${LOG} listId changed: ${String(oldValue) || '(none)'} → ${String(newValue) || '(none)'}`);
      // Convention-with-override: only clear mappings when switching BETWEEN two
      // real lists. The first selection (oldValue empty) must keep the manifest
      // defaults (Title/Video/Image0) so the 15 Practices instance needs no
      // remapping.
      if (oldValue) {
        console.log(`${LOG} switching lists — clearing stale column mappings`);
        this.properties.titleField = '';
        this.properties.videoField = '';
        this.properties.labelImageField = '';
        this.properties.mainImageField = '';
      } else {
        console.log(`${LOG} first list selection — preserving default mappings (Title/Video/Image0)`);
      }
      // Re-fetch columns for the new list, then refresh so the dropdowns
      // repopulate. Refresh immediately too so they show "Loading columns…".
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
          this._mappingDropdown('titleField', strings.TitleFieldLabel, columnsReady, false),
          this._mappingDropdown('videoField', strings.VideoFieldLabel, columnsReady, false),
          this._mappingDropdown('labelImageField', strings.LabelImageFieldLabel, columnsReady, false),
          this._mappingDropdown('mainImageField', strings.MainImageFieldLabel, columnsReady, true)
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
                  key: 'mediaGalleryListPicker'
                }),
                PropertyPaneSlider('columns', {
                  label: strings.ColumnsFieldLabel,
                  min: MIN_COLUMNS,
                  max: MAX_COLUMNS,
                  step: 1,
                  showValue: true,
                  value: this._resolvedColumns
                }),
                PropertyPaneTextField('sectionTitle', {
                  label: strings.SectionTitleFieldLabel
                }),
                PropertyPaneToggle('openInNewTab', {
                  label: strings.OpenInNewTabFieldLabel,
                  checked: this.properties.openInNewTab !== false
                })
              ]
            },
            {
              groupName: strings.FieldMappingGroupName,
              isCollapsed: false,
              groupFields: mappingFields
            }
          ]
        }
      ]
    };
  }

  // A field-mapping dropdown populated from the selected list's columns.
  // `optional` adds a "(none)" choice (used for the optional main-image override).
  private _mappingDropdown(
    targetProperty: MappingProperty,
    label: string,
    columnsReady: boolean,
    optional: boolean
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
      options = optional
        ? [{ key: NONE_KEY, text: '(none — auto-derive)' }, ...columnOptions]
        : columnOptions;
    }

    const stored = (this.properties[targetProperty] as string) || '';
    const selectedKey = optional && stored === '' ? NONE_KEY : stored || undefined;

    return PropertyPaneDropdown(targetProperty, {
      label,
      options,
      selectedKey,
      disabled: !columnsReady || columnOptions.length === 0
    });
  }
}
