import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { PHIL_BRAND_CSS } from './generated/phillipsBrandCss';

// The customizer takes no configurable properties — its only job is to inject
// the Phillips brand CSS into every page where it's registered.
type IPhillipsBrandApplicationCustomizerProperties = Record<string, never>;

const STYLE_ELEMENT_ID = 'phil-brand';

export default class PhillipsBrandApplicationCustomizer
  extends BaseApplicationCustomizer<IPhillipsBrandApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    this.injectBrandStyles();
    return Promise.resolve();
  }

  private injectBrandStyles(): void {
    // Idempotent: SharePoint's SPA shell can fire onInit again on cross-page
    // navigation within the same tab. If our <style> is already in <head>,
    // do nothing — re-appending would duplicate the brand CSS.
    if (document.getElementById(STYLE_ELEMENT_ID)) {
      return;
    }

    // Manual DOM injection rather than sp-css-loader's auto-injection — the
    // latter routes through window.__themeState__.loadStyles, which silently
    // drops :root declarations on modern SharePoint Online. See attempt 1
    // commit for the full diagnostic.
    const styleTag = document.createElement('style');
    styleTag.id = STYLE_ELEMENT_ID;
    styleTag.textContent = PHIL_BRAND_CSS;
    document.head.appendChild(styleTag);
  }
}
