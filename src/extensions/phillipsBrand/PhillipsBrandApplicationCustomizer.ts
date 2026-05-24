import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { PHIL_BRAND_CSS } from './generated/phillipsBrandCss';
import { BrandedHeader } from './components/BrandedHeader';

const LOG_SOURCE = 'PhillipsBrandApplicationCustomizer';
const STYLE_ELEMENT_ID = 'phil-brand';

type IPhillipsBrandApplicationCustomizerProperties = Record<string, never>;

export default class PhillipsBrandApplicationCustomizer
  extends BaseApplicationCustomizer<IPhillipsBrandApplicationCustomizerProperties> {

  private _topPlaceholder?: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PhillipsBrandApplicationCustomizer');

    this._injectBrandStyles();

    // Placeholders may not be ready when onInit fires — subscribe so we render
    // as soon as they become available, then attempt an immediate render too.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);
    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _injectBrandStyles(): void {
    // Idempotent: SharePoint's SPA shell can fire onInit again on cross-page
    // navigation within the same tab. If our <style> is already in <head>,
    // do nothing — re-appending would duplicate the brand CSS.
    if (document.getElementById(STYLE_ELEMENT_ID)) {
      return;
    }

    // Manual DOM injection rather than sp-css-loader's auto-injection — the
    // latter routes through window.__themeState__.loadStyles, which silently
    // drops :root declarations. See commit cfdce55 for the full diagnostic.
    const styleTag = document.createElement('style');
    styleTag.id = STYLE_ELEMENT_ID;
    styleTag.textContent = PHIL_BRAND_CSS;
    document.head.appendChild(styleTag);
  }

  private _renderPlaceholders = (): void => {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (!this._topPlaceholder) {
        Log.warn(LOG_SOURCE, 'Top placeholder is not available');
        return;
      }
    }

    if (this._topPlaceholder.domElement) {
      ReactDOM.render(
        React.createElement(BrandedHeader, { context: this.context }),
        this._topPlaceholder.domElement
      );
    }
  };

  private _onDispose = (): void => {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
  };
}
