import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
  PropertyFieldColumnPicker,
  PropertyFieldColumnPickerOrderBy,
  IColumnReturnProperty
} from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';

import * as strings from 'PhillipsMediaGalleryWebPartStrings';
import { PhillipsMediaGallery, IPhillipsMediaGalleryProps } from './components/PhillipsMediaGallery';

// @pnp/spfx-property-controls ships its own nested copy of
// @microsoft/sp-component-base, so its `context` prop type isn't structurally
// assignable from WebPartContext (separate `_serviceScope` declarations). Cast
// to the exact type the picker functions expect — avoids `any` and stays
// correct if the package bumps its bundled sp-component-base.
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

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

export default class PhillipsMediaGalleryWebPart extends BaseClientSideWebPart<IPhillipsMediaGalleryWebPartProps> {
  public render(): void {
    // Turn 1 is scaffold only — no list reads, attachment/column image
    // resolution, or thumbnail derivation yet (those land in Turn 2). The
    // component renders placeholder cards in the "loaded" state and an
    // empty-state prompt until a list is selected.
    const props: IPhillipsMediaGalleryProps = {
      columns: this._resolvedColumns,
      sectionTitle: this.properties.sectionTitle || '',
      hasList: !!this.properties.listId
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

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      // Only clear mappings when switching between two real lists — preserve the
      // manifest defaults (Title/Video/Image0) on the first list selection so the
      // column pickers show them when the chosen list has those columns.
      if (oldValue) {
        this.properties.titleField = '';
        this.properties.videoField = '';
        this.properties.labelImageField = '';
        this.properties.mainImageField = '';
      }
      // Dependent property pane: refresh so the column pickers re-render against
      // the new list. Their keys include the listId (below), so they remount and
      // re-fetch the new list's columns — the known stale-dropdown fix.
      this.context.propertyPane.refresh();
    }

    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listId = this.properties.listId || '';

    const mappingFields: IPropertyPaneField<unknown>[] = listId
      ? [
          this._columnPicker('titleField', strings.TitleFieldLabel),
          this._columnPicker('videoField', strings.VideoFieldLabel),
          this._columnPicker('labelImageField', strings.LabelImageFieldLabel),
          this._columnPicker('mainImageField', strings.MainImageFieldLabel)
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
              isCollapsed: true,
              groupFields: mappingFields
            }
          ]
        }
      ]
    };
  }

  // Build a column picker bound to the currently selected list. The key embeds
  // the listId so a list change remounts the control and forces a fresh column
  // fetch (PropertyFieldColumnPicker self-fetches columns from `listId`).
  private _columnPicker(
    targetProperty: keyof IPhillipsMediaGalleryWebPartProps,
    label: string
  ): IPropertyPaneField<unknown> {
    return PropertyFieldColumnPicker(targetProperty, {
      label,
      context: this.context as unknown as PnpContext,
      selectedColumn: this.properties[targetProperty] as string,
      listId: this.properties.listId,
      disabled: false,
      orderBy: PropertyFieldColumnPickerOrderBy.Title,
      columnReturnProperty: IColumnReturnProperty['Internal Name'],
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      deferredValidationTime: 0,
      key: `mediaGallery-${String(targetProperty)}-${this.properties.listId}`
    });
  }
}
