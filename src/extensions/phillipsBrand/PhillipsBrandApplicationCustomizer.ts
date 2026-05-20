import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

// Side-effect import: pulls the compiled Phillips brand CSS into the
// extension bundle. sp-css-loader auto-injects it into <head> at runtime
// and tracks the module so re-navigations across pages on the same SPA
// session don't duplicate the <style> tag. The `.global.scss` extension
// tells heft-sass-plugin to treat this as non-module CSS (no class hashing).
import '../../styles/index.global.scss';

// The customizer takes no configurable properties — its only job is to make
// the side-effect import above happen on every page where the customizer
// is registered. Record<string, never> keeps ESLint quiet about empty
// interfaces while still satisfying BaseApplicationCustomizer's generic.
type IPhillipsBrandApplicationCustomizerProperties = Record<string, never>;

export default class PhillipsBrandApplicationCustomizer
  extends BaseApplicationCustomizer<IPhillipsBrandApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    return Promise.resolve();
  }
}
