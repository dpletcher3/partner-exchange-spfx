import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'PhillipsCelebrationsWebPartStrings';
import { PhillipsCelebrations, IPhillipsCelebrationsProps } from './components/PhillipsCelebrations';
import { FieldMappingController, IFieldSlot } from '../../shared/fieldMapping';
import { CelebrationsDataService, ICelebrationsMapping } from './services/CelebrationsDataService';

// @pnp PropertyFieldListPicker's context type isn't structurally assignable from
// WebPartContext (nested sp-component-base). Cast to the exact expected type.
type PnpContext = Parameters<typeof PropertyFieldListPicker>[1]['context'];

const LOG = '[Celebrations]';

export interface IPhillipsCelebrationsWebPartProps {
  listId: string;
  // Field mapping (defaults applied in render / via the shared module slots).
  personField: string;
  birthdayField: string;
  hireField: string;
  // Send-a-Wish target (Turn 3); week-window + default tab.
  communityUrl: string;
  weekStart: string; // 'sunday' | 'monday'
  defaultTab: string; // 'birthdays' | 'anniversaries'
  // Send-a-Wish button visibility (default false / hidden). Undefined on existing
  // placed instances → treated as false in render(), so they need no reconfiguration.
  showSendAWish: boolean;
}

export default class PhillipsCelebrationsWebPart extends BaseClientSideWebPart<IPhillipsCelebrationsWebPartProps> {
  private _mapping!: FieldMappingController;
  private _data!: CelebrationsDataService;

  private get _slots(): IFieldSlot[] {
    return [
      { property: 'personField', label: strings.PersonFieldLabel, defaultInternalName: 'LinkedUser', typeFilter: 'person' },
      { property: 'birthdayField', label: strings.BirthdayFieldLabel, defaultInternalName: 'BirthDate', typeFilter: 'date' },
      { property: 'hireField', label: strings.HireFieldLabel, defaultInternalName: 'HireDate', typeFilter: 'date' }
    ];
  }

  protected onInit(): Promise<void> {
    this._mapping = new FieldMappingController({
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      slots: this._slots,
      getListId: () => this.properties.listId || '',
      properties: this.properties as unknown as Record<string, unknown>,
      refresh: () => this.context.propertyPane.refresh(),
      hintLabel: strings.FieldMappingEmptyLabel
    });
    this._data = new CelebrationsDataService(this.context.spHttpClient);
    return super.onInit();
  }

  public render(): void {
    // Convention-with-override: apply the §2 defaults so the data layer reads by
    // real internal names even when a mapping property is unset.
    const mapping: ICelebrationsMapping = {
      personField: this.properties.personField || 'LinkedUser',
      birthdayField: this.properties.birthdayField || 'BirthDate',
      hireField: this.properties.hireField || 'HireDate'
    };

    const props: IPhillipsCelebrationsProps = {
      service: this._data,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listId: this.properties.listId || '',
      mapping,
      weekStart: this.properties.weekStart === 'monday' ? 'monday' : 'sunday',
      defaultTab: this.properties.defaultTab === 'anniversaries' ? 'anniversaries' : 'birthdays',
      // undefined on existing placed instances → false (button hidden); no reconfig needed.
      showSendAWish: this.properties.showSendAWish === true
    };
    console.log(`${LOG} render: listId=${props.listId || '(none)'}, weekStart=${props.weekStart}, defaultTab=${props.defaultTab}`);
    ReactDom.render(React.createElement(PhillipsCelebrations, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    // Load the current list's columns on pane open (handles an already-selected list).
    this._mapping.load();
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      console.log(`${LOG} listId changed: ${String(oldValue) || '(none)'} → ${String(newValue) || '(none)'}`);
      // Shared module: clears slots only when switching between real lists,
      // re-fetches the new list's columns, and refreshes the pane.
      this._mapping.onListChanged(oldValue);
    }

    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const weekStartOptions: IPropertyPaneDropdownOption[] = [
      { key: 'sunday', text: strings.WeekStartSundayLabel },
      { key: 'monday', text: strings.WeekStartMondayLabel }
    ];
    const defaultTabOptions: IPropertyPaneDropdownOption[] = [
      { key: 'birthdays', text: strings.TabBirthdaysLabel },
      { key: 'anniversaries', text: strings.TabAnniversariesLabel }
    ];

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
                  key: 'celebrationsListPicker'
                }),
                PropertyPaneDropdown('weekStart', {
                  label: strings.WeekStartFieldLabel,
                  options: weekStartOptions,
                  selectedKey: this.properties.weekStart || 'sunday'
                }),
                PropertyPaneDropdown('defaultTab', {
                  label: strings.DefaultTabFieldLabel,
                  options: defaultTabOptions,
                  selectedKey: this.properties.defaultTab || 'birthdays'
                }),
                PropertyPaneTextField('communityUrl', {
                  label: strings.CommunityUrlFieldLabel
                }),
                PropertyPaneToggle('showSendAWish', {
                  label: strings.ShowSendAWishFieldLabel
                })
              ]
            },
            {
              groupName: strings.FieldMappingGroupName,
              // Expanded — a collapsed group reads as "no dropdowns" (lessons-learned).
              isCollapsed: false,
              groupFields: this._mapping.buildFields()
            }
          ]
        }
      ]
    };
  }
}
