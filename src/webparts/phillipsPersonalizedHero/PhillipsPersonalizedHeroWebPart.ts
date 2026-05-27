import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

// The file picker's `context` prop expects PnP's own BaseComponentContext, which
// resolves to PnP's nested @microsoft/sp-component-base@1.22.2 — a different type
// identity than the project's WebPartContext (sp-component-base@1.23.0). Deriving
// the prop's exact context type and casting to it bridges the skew without `any`;
// the runtime object IS a valid context, only the compile-time identity differs.
type FilePickerContext = Parameters<typeof PropertyFieldFilePicker>[1]['context'];

// PnP's PropertyFieldFilePicker reports the full IFilePickerResult to
// onPropertyChange, which the property-pane framework persists into the
// backgroundImage property — so the stored value may be that object, not a URL
// string. Coerce to a usable URL at read time, immune to which shape is stored.
function toImageUrl(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }
  if (value && typeof value === 'object') {
    const url = (value as { fileAbsoluteUrl?: unknown }).fileAbsoluteUrl;
    if (typeof url === 'string') {
      return url;
    }
  }
  return '';
}

import * as strings from 'PhillipsPersonalizedHeroWebPartStrings';
import {
  PhillipsPersonalizedHero,
  IPhillipsPersonalizedHeroProps,
  BackgroundType,
  GreetingAlignment
} from './components/PhillipsPersonalizedHero';

export interface IPhillipsPersonalizedHeroWebPartProps {
  backgroundType: string;
  backgroundColor: string;
  backgroundImage: string;
  bannerHeight: number;
  greetingColor: string;
  greetingSize: number;
  greetingWeight: number;
  greetingAlignment: string;
}

export default class PhillipsPersonalizedHeroWebPart extends BaseClientSideWebPart<IPhillipsPersonalizedHeroWebPartProps> {
  private _filePickerResult: IFilePickerResult | undefined;

  public render(): void {
    const element: React.ReactElement<IPhillipsPersonalizedHeroProps> = React.createElement(
      PhillipsPersonalizedHero,
      {
        backgroundType: (this.properties.backgroundType as BackgroundType) || 'Color',
        backgroundColor: this.properties.backgroundColor || '#C8102E',
        backgroundImage: toImageUrl(this.properties.backgroundImage),
        bannerHeight: this.properties.bannerHeight || 450,
        greetingColor: this.properties.greetingColor || '#FFFFFF',
        greetingSize: this.properties.greetingSize || 42,
        greetingWeight: this.properties.greetingWeight || 500,
        greetingAlignment: (this.properties.greetingAlignment as GreetingAlignment) || 'Left',
        displayName: this.context.pageContext.user.displayName,
        isEditMode: this.displayMode === DisplayMode.Edit
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isImage = this.properties.backgroundType === 'Image';

    // backgroundColor (full picker) and backgroundImage (file picker) are
    // mutually exclusive — the dropdown selects which one shows. The file picker
    // exposes Browse (site files), Upload, and From-a-link; the tabs that need
    // extra tenant config (stock/web search, org assets, OneDrive, recent) are
    // hidden. Selected file URL is stored in `backgroundImage`, which the
    // component already reads, so no component change is needed.
    const backgroundField = isImage
      ? PropertyFieldFilePicker('backgroundImage', {
          context: this.context as unknown as FilePickerContext,
          properties: this.properties,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          filePickerResult: this._filePickerResult as IFilePickerResult,
          onSave: (result: IFilePickerResult) => {
            this._filePickerResult = result;
            this.properties.backgroundImage = result.fileAbsoluteUrl || '';
            this.render();
          },
          onChanged: (result: IFilePickerResult) => {
            this._filePickerResult = result;
          },
          key: 'backgroundImageFieldId',
          buttonLabel: strings.BackgroundImageButtonLabel,
          label: strings.BackgroundImageFieldLabel,
          accepts: ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.svg'],
          hideRecentTab: true,
          hideWebSearchTab: true,
          hideStockImages: true,
          hideOrganisationalAssetTab: true,
          hideOneDriveTab: true,
          hideSiteFilesTab: false,
          hideLocalUploadTab: false,
          hideLinkUploadTab: false
        })
      : PropertyFieldColorPicker('backgroundColor', {
          label: strings.BackgroundColorFieldLabel,
          selectedColor: this.properties.backgroundColor,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          style: PropertyFieldColorPickerStyle.Full,
          key: 'backgroundColorFieldId'
        });

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BannerGroupName,
              groupFields: [
                PropertyPaneDropdown('backgroundType', {
                  label: strings.BackgroundTypeFieldLabel,
                  options: [
                    { key: 'Color', text: 'Color' },
                    { key: 'Image', text: 'Image' }
                  ],
                  selectedKey: this.properties.backgroundType || 'Color'
                }),
                backgroundField,
                PropertyPaneSlider('bannerHeight', {
                  label: strings.BannerHeightFieldLabel,
                  min: 200,
                  max: 600,
                  step: 10,
                  showValue: true
                })
              ]
            },
            {
              groupName: strings.GreetingGroupName,
              groupFields: [
                PropertyFieldColorPicker('greetingColor', {
                  label: strings.GreetingColorFieldLabel,
                  selectedColor: this.properties.greetingColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'greetingColorFieldId'
                }),
                PropertyPaneDropdown('greetingSize', {
                  label: strings.GreetingSizeFieldLabel,
                  options: [
                    { key: 28, text: '28' },
                    { key: 36, text: '36' },
                    { key: 42, text: '42' },
                    { key: 48, text: '48' },
                    { key: 56, text: '56' }
                  ],
                  selectedKey: this.properties.greetingSize || 42
                }),
                PropertyPaneDropdown('greetingWeight', {
                  label: strings.GreetingWeightFieldLabel,
                  options: [
                    { key: 400, text: 'Regular (400)' },
                    { key: 500, text: 'Medium (500)' },
                    { key: 700, text: 'Bold (700)' }
                  ],
                  selectedKey: this.properties.greetingWeight || 500
                }),
                PropertyPaneChoiceGroup('greetingAlignment', {
                  label: strings.GreetingAlignmentFieldLabel,
                  options: [
                    { key: 'Left', text: 'Left' },
                    { key: 'Center', text: 'Center' },
                    { key: 'Right', text: 'Right' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
