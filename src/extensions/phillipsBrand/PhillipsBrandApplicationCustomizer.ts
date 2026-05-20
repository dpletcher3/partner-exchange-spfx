import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

// The customizer takes no configurable properties — its job is to inject the
// Phillips brand CSS into every page. Using Record<string, never> instead of
// an empty interface keeps ESLint quiet about no-empty-interface.
type IPhillipsBrandApplicationCustomizerProperties = Record<string, never>;

export default class PhillipsBrandApplicationCustomizer
  extends BaseApplicationCustomizer<IPhillipsBrandApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    // Prompt 3: prove the customizer is loading. This log gets replaced in
    // Prompt 4 by the actual CSS injection. The "[PhillipsBrand]" prefix is
    // a searchable pattern in the browser console.
    console.log('[PhillipsBrand] Customizer initialized');
    return Promise.resolve();
  }
}
