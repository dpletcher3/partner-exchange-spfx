import * as React from 'react';

import styles from './PhillipsAudienceHero.module.scss';
import { AudienceProfileService } from '../services/AudienceProfileService';
import { IAudienceTile } from '../services/models';
import { isTileVisible } from '../utils/audienceMatch';

const LOG = '[AudienceHero]';

export interface IPhillipsAudienceHeroProps {
  service: AudienceProfileService;
  partnerProfilesSiteUrl: string;
  listTitle: string;
  personField: string;
  divisionField: string;
  viewerEmail: string;
  tiles: IAudienceTile[];
  isEditMode: boolean;
}

type Status = 'loading' | 'ready' | 'error';

type LayoutMode = 'single' | 'stacked' | 'leftLarge' | 'evenGrid' | 'topLarge';

interface ILayout {
  mode: LayoutMode;
  // Column count for the even / top-large grids (consumed via --phil-ah-cols).
  columns: number;
  // Whether the first visible tile is the sized-up "large" tile.
  largeFirst: boolean;
}

// Balanced even-grid column count: two rows up to 10 tiles (4→2, 6→3, 8→4,
// 10→5), then capped at 5 and chosen for the fullest last row (e.g. 12→4 = 4×3).
function chooseColumns(evenCount: number): number {
  const half = Math.floor(evenCount / 2);
  if (half <= 5) {
    return half < 1 ? 1 : half;
  }
  const remOf5 = (5 - (evenCount % 5)) % 5;
  const remOf4 = (4 - (evenCount % 4)) % 4;
  return remOf4 < remOf5 ? 4 : 5;
}

// Choose an arrangement from the RENDERED (post division-filter) tile count:
//   1 → single full width · 2 → stacked · 3 → large-left + 2 stacked right
//   even N≥4 → balanced even grid · odd N≥5 → large-on-top + even grid below.
// (Only N=3 uses large-left; all other odd counts put the large tile on top.)
function computeLayout(n: number): ILayout {
  if (n <= 1) {
    return { mode: 'single', columns: 1, largeFirst: true };
  }
  if (n === 2) {
    return { mode: 'stacked', columns: 1, largeFirst: false };
  }
  if (n === 3) {
    return { mode: 'leftLarge', columns: 2, largeFirst: true };
  }
  if (n % 2 === 0) {
    return { mode: 'evenGrid', columns: chooseColumns(n), largeFirst: false };
  }
  return { mode: 'topLarge', columns: chooseColumns(n - 1), largeFirst: true };
}

function modeClassName(mode: LayoutMode): string {
  switch (mode) {
    case 'single':
      return styles.modeSingle;
    case 'stacked':
      return styles.modeStacked;
    case 'leftLarge':
      return styles.modeLeftLarge;
    case 'topLarge':
      return styles.modeTopLarge;
    case 'evenGrid':
    default:
      return styles.modeEvenGrid;
  }
}

export const PhillipsAudienceHero: React.FC<IPhillipsAudienceHeroProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');
  const [viewerDivision, setViewerDivision] = React.useState<string | undefined>(undefined);

  // Re-run the lookup when the source config or the viewer changes. Keyed on a
  // primitive so prop-object identity churn doesn't trigger spurious fetches.
  const depsKey =
    `${props.partnerProfilesSiteUrl}|${props.listTitle}|${props.personField}|` +
    `${props.divisionField}|${props.viewerEmail}`;

  const load = React.useCallback(() => {
    setStatus('loading');
    // Missing source/identity → resolved-with-no-division (NOT an error UI).
    if (!props.partnerProfilesSiteUrl || !props.listTitle || !props.viewerEmail) {
      console.warn(`${LOG} missing site/list/viewer email — fail-closed (no tiles)`);
      setViewerDivision(undefined);
      setStatus('ready');
      return;
    }
    props.service
      .getViewerDivision({
        siteUrl: props.partnerProfilesSiteUrl,
        listTitle: props.listTitle,
        personField: props.personField || 'LinkedUser',
        divisionField: props.divisionField || 'Division',
        viewerEmail: props.viewerEmail
      })
      .then((division) => {
        setViewerDivision(division);
        setStatus('ready');
      })
      .catch((err: unknown) => {
        console.warn(`${LOG} profile lookup failed — fail-closed (no tiles)`, err);
        setViewerDivision(undefined);
        setStatus('error');
      });
  }, [props.service, depsKey]);

  React.useEffect(() => {
    load();
  }, [load]);

  if (status === 'loading') {
    return (
      <div className={styles.audienceHero}>
        <div className={styles.loading} aria-hidden="true" />
      </div>
    );
  }

  const visible = (props.tiles || []).filter((t) => isTileVisible(viewerDivision, t.allowedDivisions));

  // Zero visible tiles is EXPECTED (non-partners, no matching division, or a
  // failed lookup) — fail-closed. Render nothing in display mode; in edit mode
  // show a quiet hint so authors understand why the surface is empty.
  if (visible.length === 0) {
    if (!props.isEditMode) {
      return <div className={styles.audienceHero} />;
    }
    let reason: string;
    if (status === 'error') {
      reason =
        'Could not resolve your Partner Profiles division (lookup failed). Fail-closed: no tiles are shown to anyone until the lookup succeeds.';
    } else if (!viewerDivision) {
      reason =
        'No Partner Profiles row matched your account. Fail-closed by design: viewers with no profile (or in an unlisted division) see no tiles.';
    } else {
      reason = `Your division "${viewerDivision}" is not in any tile's allowed divisions. Add it to a tile to show that tile to your division.`;
    }
    return (
      <div className={styles.audienceHero}>
        <p className={styles.editHint}>{reason}</p>
      </div>
    );
  }

  // Layout adapts to the RENDERED (post-filter) count — not the configured count.
  const layout = computeLayout(visible.length);
  const gridStyle = { ['--phil-ah-cols']: String(layout.columns) } as React.CSSProperties;

  return (
    <div className={styles.audienceHero}>
      <div className={`${styles.grid} ${modeClassName(layout.mode)}`} style={gridStyle}>
        {visible.map((t, i) => {
          const isLarge = layout.largeFirst && i === 0;
          // Background image is passed as a CSS custom property (consumed in SCSS),
          // matching the repo's inline-custom-property pattern. No image → the
          // SCSS neutral fill shows.
          const tileStyle = t.imageUrl
            ? ({ ['--phil-ah-image']: `url("${t.imageUrl.replace(/"/g, '%22')}")` } as React.CSSProperties)
            : undefined;
          return (
            <a
              key={i}
              className={isLarge ? `${styles.tile} ${styles.large}` : styles.tile}
              href={t.linkUrl || undefined}
              style={tileStyle}
            >
              <span className={styles.scrim} aria-hidden="true" />
              <span className={styles.header}>{t.header}</span>
            </a>
          );
        })}
      </div>
    </div>
  );
};
