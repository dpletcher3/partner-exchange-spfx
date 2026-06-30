import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import type { ISPEventObserver } from '@microsoft/sp-core-library';
import styles from './BrandedHeader.module.scss';

// Base64-encoded Phillips wordmark. Sourced from src/assets/Logo.png and
// inlined to keep the header render synchronous — no extra network hop for
// the logo and nothing for CSP to block. Regenerate via:
//   base64 -w0 src/assets/Logo.png
// then paste below.
const LOGO_DATA_URI =
  'data:image/png;base64,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';

const SUBTITLE = 'PARTNER EXCHANGE';

interface IBrandedHeaderProps {
  context: ApplicationCustomizerContext;
}

interface INavItem {
  Title: string;
  SimpleUrl: string;
}

// Sandbox tenant URLs — acceptable per D023 (sandbox specifics live in code,
// production deploy updates these). Used both as the fallback when the hub
// nav fetch fails and as the canonical site root for the active-state check.
const FALLBACK_NAV: INavItem[] = [
  { Title: 'Home',         SimpleUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox' },
  { Title: 'Our Culture',  SimpleUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox-OurCulture' },
  { Title: 'Our Partners', SimpleUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox-OurPartners' },
  { Title: 'Dashboard',    SimpleUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox-Dashboard' },
  { Title: 'The Hub',      SimpleUrl: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox-TheHub' },
];

const HUB_ROOT_URL = 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox';
const CACHE_KEY = 'phil-hub-nav-cache';
const CACHE_TTL_MS = 5 * 60 * 1000;

interface ICachedNav {
  ts: number;
  items: INavItem[];
}

// Strip protocol + host so two URLs from different absolute forms can be
// compared by path alone. Returns lowercase path without trailing slash.
function urlPath(url: string): string {
  try {
    const u = new URL(url, window.location.href);
    let p = u.pathname.toLowerCase();
    if (p.length > 1 && p.endsWith('/')) {
      p = p.substring(0, p.length - 1);
    }
    return p;
  } catch {
    return url.toLowerCase();
  }
}

// Active iff the current path equals the link path, or continues past it with
// a `/`. The `+ '/'` guard is what prevents `/sites/PartnerExchange-Sandbox`
// from incorrectly matching `/sites/PartnerExchange-Sandbox-OurCulture`.
function isNavItemActive(itemUrl: string, currentUrl: string): boolean {
  const link = urlPath(itemUrl);
  const here = urlPath(currentUrl);
  if (here === link) {
    return true;
  }
  return here.indexOf(link + '/') === 0;
}

export const BrandedHeader: React.FC<IBrandedHeaderProps> = (props) => {
  const [navItems, setNavItems] = React.useState<INavItem[] | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    // 1) sessionStorage cache check
    try {
      const raw = window.sessionStorage.getItem(CACHE_KEY);
      if (raw) {
        const parsed: ICachedNav = JSON.parse(raw);
        if (parsed && Date.now() - parsed.ts < CACHE_TTL_MS && Array.isArray(parsed.items)) {
          setNavItems(parsed.items);
          setLoading(false);
          return;
        }
      }
    } catch {
      // Cache parse failure is non-fatal; fall through to fetch.
    }

    // 2) Fetch hub nav from REST
    const navUrl = `${HUB_ROOT_URL}/_api/web/Navigation/TopNavigationBar`;
    props.context.spHttpClient.get(navUrl, SPHttpClient.configurations.v1)
      .then((r) => r.json())
      .then((data: { value: { Title: string; Url: string }[] }) => {
        const items: INavItem[] = (data.value || []).map((node) => ({
          Title: node.Title,
          SimpleUrl: node.Url
        }));
        try {
          window.sessionStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), items } as ICachedNav));
        } catch {
          // sessionStorage write failure is non-fatal.
        }
        setNavItems(items);
        setLoading(false);
      })
      .catch((err: unknown) => {
        console.warn('Phillips Brand: hub nav fetch failed; using fallback', err);
        setNavItems(FALLBACK_NAV);
        setLoading(false);
      });
  }, [props.context]);

  // Current URL as reactive state so the active-nav underline follows SharePoint
  // client-side (SPA) navigation. The links are plain <a href>; SP intercepts
  // them and navigates without remounting this placeholder React tree, so a
  // render-time snapshot of window.location.href goes stale and the underline
  // would not move on a single click. We recompute it on every navigation.
  const [currentUrl, setCurrentUrl] = React.useState<string>(window.location.href);

  // Subscribe to navigation and refresh currentUrl. Kept separate from the
  // nav-fetch effect above. onNavigated reads window.location.href fresh on each
  // call (no stale closure). SPFx's SPEvent.add/remove require an ISPEventObserver
  // owner; a functional component has no BaseComponent `this`, so we build a
  // minimal stable observer from props.context. add() and remove() run in the
  // same effect closure, so the exact same observer/handler instances are used
  // for both (as the SDK requires).
  React.useEffect(() => {
    const onNavigated = (): void => setCurrentUrl(window.location.href);
    const observer: ISPEventObserver = {
      instanceId: props.context.instanceId,
      // componentId is the component GUID from the manifest. (The optional
      // `manifest` member of ISPEventObserver is @internal / stripped from the
      // public typings, so it is intentionally not set here.)
      componentId: props.context.manifest.id,
      isDisposed: false,
      dispose: () => { /* no-op: handler removal is handled in the cleanup below */ }
    };
    props.context.application.navigatedEvent.add(observer, onNavigated);
    window.addEventListener('popstate', onNavigated);
    return () => {
      props.context.application.navigatedEvent.remove(observer, onNavigated);
      window.removeEventListener('popstate', onNavigated);
    };
  }, [props.context]);

  // No FluentProvider wrapper: this component renders only plain HTML elements
  // styled by CSS modules — no Fluent v9 components used. Mounting a v9
  // FluentProvider here was load-bearing for nothing AND was actively breaking
  // SharePoint's own portaled v9 panels (New Item form, web part property
  // panes, Site Information, Highlighted Content config) because SP's panels
  // resolve their --colorNeutralBackground1 theme token via a `fui-FluentProvider#`
  // class that collides when a second FluentProvider mounts on the page (see
  // microsoft/fluentui#23821, SharePoint/sp-dev-docs#9847). Compounded by
  // applyStylesToPortals=true (the v9 default), which leaks our provider's
  // variables into SP's panel portals. Net effect: panels paint
  // `background-color: var(--colorNeutralBackground1)` with no value resolved
  // → transparent surface, page bleeds through. Removed in Iter 2c.3.
  return (
    <div className={styles.brandedHeader}>
      <div className={styles.logoRegion}>
        <img src={LOGO_DATA_URI} className={styles.logo} alt="Phillips Corporation" />
        <div className={styles.subtitle}>{SUBTITLE}</div>
      </div>
      <nav className={styles.navRegion} aria-label="Partner Exchange hub navigation">
        {loading || !navItems
          ? Array.from({ length: 5 }).map((_, i) => (
              <div key={`skeleton-${i}`} className={styles.skeleton} aria-hidden="true" />
            ))
          : navItems.map((item) => {
              const active = isNavItemActive(item.SimpleUrl, currentUrl);
              const className = active
                ? `${styles.navItem} ${styles.active}`
                : styles.navItem;
              return (
                <a
                  key={item.SimpleUrl}
                  href={item.SimpleUrl}
                  className={className}
                  aria-current={active ? 'page' : undefined}
                >
                  {item.Title}
                </a>
              );
            })}
      </nav>
    </div>
  );
};
