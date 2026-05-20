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

    // TEMPORARY DIAGNOSTIC — remove after end-to-end verification.
    // The brand CSS itself injects --phil-* custom properties on :root, but
    // the sandbox home page has nothing yet that consumes them, so a working
    // injection looks identical to a non-working one. The banner below makes
    // onInit execution visible regardless of what's on the page.
    const banner = document.createElement('div');
    banner.id = 'phil-test-banner';
    banner.textContent = '🟥 PARTNER EXCHANGE CUSTOMIZER LOADED 🟥';
    banner.style.cssText = [
      'position: fixed',
      'top: 0',
      'left: 0',
      'right: 0',
      'z-index: 2147483647',
      'background: #F9423A',
      'color: white',
      'font: bold 16px/40px Arial, sans-serif',
      'text-align: center',
      'padding: 0 16px',
      'box-shadow: 0 2px 8px rgba(0,0,0,0.3)'
    ].join(';');
    document.body.appendChild(banner);

    console.log('[PhilCustomizer] Banner injected, onInit complete');

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
