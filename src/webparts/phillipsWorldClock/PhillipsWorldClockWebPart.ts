import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'PhillipsWorldClockWebPartStrings';
import {
  PhillipsWorldClock,
  IPhillipsWorldClockProps,
  IClockConfig
} from './components/PhillipsWorldClock';
import { isValidTimezone } from './services/timeFormatter';

export interface IPhillipsWorldClockWebPartProps {
  sectionTitle: string;
  clocks: IClockConfig[];
}

// Maximum clock count called out in the spec. Enforced at the property
// pane level via PropertyFieldCollectionData's `maximumItems`.
const MAX_CLOCKS = 12;

const TZ_HELP_URL = 'https://en.wikipedia.org/wiki/List_of_tz_database_time_zones';

export default class PhillipsWorldClockWebPart extends BaseClientSideWebPart<IPhillipsWorldClockWebPartProps> {
  public render(): void {
    const props: IPhillipsWorldClockProps = {
      sectionTitle: this.properties.sectionTitle || '',
      clocks: this.properties.clocks || [],
      unconfiguredMessage: strings.UnconfiguredMessage
    };

    ReactDom.render(React.createElement(PhillipsWorldClock, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyFieldCollectionData('clocks', {
                  key: 'clocksField',
                  label: strings.ClocksFieldLabel,
                  panelHeader: strings.ClocksPanelHeader,
                  manageBtnLabel: strings.ClocksManageButtonLabel,
                  value: this.properties.clocks || [],
                  enableSorting: true,
                  // Spec caps at 12 entries — disable the Add row once we're
                  // at the limit so the editor can't push past it.
                  disableItemCreation: (this.properties.clocks || []).length >= MAX_CLOCKS,
                  fields: [
                    {
                      id: 'title',
                      title: strings.ClockTitleColumnLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'timezone',
                      title: strings.ClockTimezoneColumnLabel,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'America/New_York',
                      // Validates via Intl.DateTimeFormat construction: an
                      // unknown zone throws RangeError, which the helper
                      // catches and surfaces as a clean inline error.
                      onGetErrorMessage: (value: string): string => {
                        if (!value) {
                          return strings.ClockTimezoneRequiredError;
                        }
                        return isValidTimezone(value)
                          ? ''
                          : strings.ClockTimezoneInvalidError;
                      }
                    }
                  ]
                }),
                PropertyPaneLink('timezoneHelpLink', {
                  href: TZ_HELP_URL,
                  text: strings.TimezoneHelpLinkText,
                  target: '_blank'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
