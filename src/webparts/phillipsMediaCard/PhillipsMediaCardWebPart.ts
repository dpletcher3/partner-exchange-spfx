import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldFilePicker,
  IFilePickerResult,
  IPropertyFieldFilePickerProps
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import * as strings from 'PhillipsMediaCardWebPartStrings';
import { PhillipsMediaCard, IPhillipsMediaCardProps } from './components/PhillipsMediaCard';
import { VideoSourceType } from './services/videoSource';
import { getVimeoThumbnail } from './services/videoThumbnail';

export interface IPhillipsMediaCardWebPartProps {
  eyebrow: string;
  header: string;
  description: string;
  // PropertyFieldFilePicker stores its selection as an IFilePickerResult-
  // shaped object. The downloadFileContent function field is lost on
  // serialization — we don't use it, so it's fine.
  imagePicker: IFilePickerResult | undefined;
  videoUrl: string;
  videoSourceType: VideoSourceType;
  // Vimeo thumbnail URL resolved at edit time via oEmbed and persisted so the
  // runtime render doesn't trigger a network round-trip. Cleared when source
  // type leaves Vimeo or when the fetch fails.
  vimeoThumbnailUrl: string;
}

// Max lengths called out in the spec — enforced at the property pane level so
// the editor sees the cap rather than hitting it after publish.
const EYEBROW_MAX = 30;
const HEADER_MAX = 60;
const DESCRIPTION_MAX = 300;

// Tracks the lifecycle of the in-flight Vimeo oEmbed call. 'idle' is the
// resting state for non-Vimeo sources; the others drive the Image-field
// label / required flag / inline error.
type VimeoFetchStatus = 'idle' | 'loading' | 'success' | 'error';

export default class PhillipsMediaCardWebPart extends BaseClientSideWebPart<IPhillipsMediaCardWebPartProps> {
  // Vimeo fetch state — not persisted; we re-derive from the URL when the
  // pane opens on a Vimeo-configured instance by trusting the stored
  // vimeoThumbnailUrl (success ⇒ has URL, otherwise idle).
  private _vimeoFetchStatus: VimeoFetchStatus = 'idle';
  // Incrementing token guards against stale fetch responses overwriting
  // newer ones when the editor edits the URL quickly.
  private _vimeoFetchToken = 0;

  protected onInit(): Promise<void> {
    // Restore fetch status from the persisted property so the property pane
    // shows the right Image-field label on first open. If the URL is set and
    // we have a thumbnail, the previous fetch succeeded. Otherwise idle; an
    // editor edit will retrigger the fetch.
    if (
      this.properties.videoSourceType === 'vimeo' &&
      this.properties.videoUrl &&
      this.properties.vimeoThumbnailUrl
    ) {
      this._vimeoFetchStatus = 'success';
    }
    return super.onInit();
  }

  public render(): void {
    const imageUrl = this.properties.imagePicker?.fileAbsoluteUrl || '';
    const imageAlt = this.properties.imagePicker?.fileNameWithoutExtension || '';

    const props: IPhillipsMediaCardProps = {
      eyebrow: this.properties.eyebrow || '',
      header: this.properties.header || '',
      description: this.properties.description || '',
      imageUrl,
      imageAlt,
      videoUrl: this.properties.videoUrl || '',
      videoSourceType: this.properties.videoSourceType || 'sharepoint',
      vimeoThumbnailUrl: this.properties.vimeoThumbnailUrl || '',
      unconfiguredMessage: strings.UnconfiguredMessage,
      playAriaLabel: strings.PlayButtonAriaLabel,
      closeAriaLabel: strings.CloseModalAriaLabel
    };

    ReactDom.render(React.createElement(PhillipsMediaCard, props), this.domElement);
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

    if (propertyPath === 'videoUrl') {
      // URL changed. If we're on Vimeo and the URL is non-empty, kick off
      // an oEmbed fetch; if the URL was cleared, drop any stored thumbnail
      // so a stale value can't leak into the rendered card.
      if (!newValue) {
        this.properties.vimeoThumbnailUrl = '';
        this._vimeoFetchStatus = 'idle';
        // Bump the token so any in-flight fetch from the old URL is
        // discarded when it resolves.
        this._vimeoFetchToken++;
      } else if (this.properties.videoSourceType === 'vimeo') {
        void this._fetchVimeoThumbnail(String(newValue));
      }
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'videoSourceType' && oldValue !== newValue) {
      // Source type flipped. Clear any Vimeo thumbnail so a YouTube URL
      // doesn't accidentally render through a leftover Vimeo poster, then
      // trigger a fresh oEmbed fetch if the new source is Vimeo.
      this.properties.vimeoThumbnailUrl = '';
      this._vimeoFetchStatus = 'idle';
      this._vimeoFetchToken++;
      if (newValue === 'vimeo' && this.properties.videoUrl) {
        void this._fetchVimeoThumbnail(this.properties.videoUrl);
      }
      this.context.propertyPane.refresh();
    }

    this.render();
  }

  // Fetches the Vimeo thumbnail at edit time. Captures an incrementing token
  // so a stale response (older URL) doesn't overwrite a newer one if the
  // editor types or pastes faster than the fetch resolves.
  private async _fetchVimeoThumbnail(url: string): Promise<void> {
    const myToken = ++this._vimeoFetchToken;
    this._vimeoFetchStatus = 'loading';
    this.context.propertyPane.refresh();

    const thumb = await getVimeoThumbnail(url);

    if (myToken !== this._vimeoFetchToken) {
      // Superseded by a newer fetch — discard this result.
      return;
    }

    if (thumb) {
      this.properties.vimeoThumbnailUrl = thumb;
      this._vimeoFetchStatus = 'success';
    } else {
      this.properties.vimeoThumbnailUrl = '';
      this._vimeoFetchStatus = 'error';
    }
    this.context.propertyPane.refresh();
    this.render();
  }

  // Decides which Image field label to show and whether the field is marked
  // required, based on the current source type and Vimeo fetch state. Editors
  // see at a glance whether they need to upload an image or can rely on the
  // auto-derived thumbnail.
  private _resolveImageFieldUx(): { label: string; required: boolean; error: string | undefined } {
    const sourceType = this.properties.videoSourceType;
    const hasVideoUrl = !!this.properties.videoUrl;

    if (hasVideoUrl && sourceType === 'youtube') {
      return { label: strings.ImageFieldLabelYouTube, required: false, error: undefined };
    }
    if (hasVideoUrl && sourceType === 'vimeo') {
      if (this._vimeoFetchStatus === 'loading') {
        return { label: strings.ImageFieldLabelVimeoLoading, required: false, error: undefined };
      }
      if (this._vimeoFetchStatus === 'success') {
        return { label: strings.ImageFieldLabelVimeoSuccess, required: false, error: undefined };
      }
      // Either 'error' or 'idle' with no thumbnail — require an uploaded
      // image and surface the inline error so the editor knows why.
      return {
        label: strings.ImageFieldLabelVimeoError,
        required: true,
        error: strings.VimeoFetchErrorMessage
      };
    }
    // SharePoint source OR no video URL — Image is required, standard label.
    return { label: strings.ImageFieldLabel, required: true, error: undefined };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const videoSourceOptions: IPropertyPaneDropdownOption[] = [
      { key: 'sharepoint', text: strings.VideoSourceSharePointLabel },
      { key: 'youtube', text: strings.VideoSourceYouTubeLabel },
      { key: 'vimeo', text: strings.VideoSourceVimeoLabel }
    ];

    const imageUx = this._resolveImageFieldUx();

    const contentFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('eyebrow', {
        label: strings.EyebrowFieldLabel,
        maxLength: EYEBROW_MAX
      }),
      PropertyPaneTextField('header', {
        label: strings.HeaderFieldLabel,
        maxLength: HEADER_MAX
      }),
      PropertyPaneTextField('description', {
        label: strings.DescriptionFieldLabel,
        multiline: true,
        rows: 5,
        maxLength: DESCRIPTION_MAX
      })
    ];

    const mediaFields: IPropertyPaneField<unknown>[] = [
      PropertyFieldFilePicker('imagePicker', {
        // @pnp/spfx-property-controls ships its own nested copy of
        // @microsoft/sp-component-base, so PnP's BaseComponentContext is a
        // nominally distinct type from the one our top-level SPFx packages
        // resolve to even though the runtime class is the same. Routing the
        // cast through PnP's own props interface picks up *its* view of the
        // type so the assignment type-checks without a blanket `any`.
        context: this.context as unknown as IPropertyFieldFilePickerProps['context'],
        filePickerResult: this.properties.imagePicker as IFilePickerResult,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        onSave: (result: IFilePickerResult) => {
          this.properties.imagePicker = result;
          this.render();
        },
        onChanged: (result: IFilePickerResult) => {
          this.properties.imagePicker = result;
        },
        key: 'imagePickerField',
        buttonLabel: strings.ImagePickerButtonLabel,
        label: imageUx.label,
        accepts: ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg'],
        required: imageUx.required,
        hideWebSearchTab: true,
        hideLinkUploadTab: true
      })
    ];

    if (imageUx.error) {
      // Surface the Vimeo-fetch failure right under the FilePicker. The
      // PropertyFieldFilePicker control doesn't expose an errorMessage prop
      // (unlike PropertyPaneTextField), so we render a separate label as the
      // inline error.
      mediaFields.push(
        PropertyPaneLabel('vimeoFetchError', {
          text: imageUx.error
        })
      );
    }

    mediaFields.push(
      PropertyPaneTextField('videoUrl', {
        label: strings.VideoUrlFieldLabel,
        placeholder: 'https://...'
      })
    );

    if (this.properties.videoUrl) {
      mediaFields.push(
        PropertyPaneDropdown('videoSourceType', {
          label: strings.VideoSourceTypeFieldLabel,
          options: videoSourceOptions,
          selectedKey: this.properties.videoSourceType || 'sharepoint'
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
              groupName: strings.MediaGroupName,
              groupFields: mediaFields
            }
          ]
        }
      ]
    };
  }
}
