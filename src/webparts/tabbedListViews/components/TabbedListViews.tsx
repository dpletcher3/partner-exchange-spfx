import * as React from 'react';

import styles from './TabbedListViews.module.scss';
import { ITabbedListViewsService } from '../services/ITabbedListViewsService';
import { ITabConfig, ITabData, Layout, OverlayPosition } from '../services/models';
import { GalleryLayout } from './GalleryLayout';
import { TableLayout } from './TableLayout';

export interface ITabbedListViewsProps {
  service: ITabbedListViewsService;
  siteUrl: string;
  sectionTitle: string;
  listId: string;
  layout: Layout;
  tabCount: number;
  tabs: ITabConfig[];
  seeAllUrl: string;
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
}

type Status = 'loading' | 'populated' | 'empty' | 'error' | 'unconfigured';

export const TabbedListViews: React.FC<ITabbedListViewsProps> = (props) => {
  // Slice tabs down to tabCount at render so reducing the count hides extras
  // without losing them from properties (they reappear when count rises).
  const visibleTabs = React.useMemo(
    () => (props.tabs || []).slice(0, props.tabCount),
    [props.tabs, props.tabCount]
  );

  const [activeIdx, setActiveIdx] = React.useState<number>(0);
  const [status, setStatus] = React.useState<Status>('loading');
  const [tabData, setTabData] = React.useState<ITabData | undefined>(undefined);

  // Clamp activeIdx if a tab was deleted from underneath us.
  React.useEffect(() => {
    if (activeIdx >= visibleTabs.length) {
      setActiveIdx(0);
    }
  }, [visibleTabs.length, activeIdx]);

  const activeTab: ITabConfig | undefined = visibleTabs[activeIdx];

  React.useEffect(() => {
    if (!props.listId || !activeTab || !activeTab.viewId) {
      setStatus('unconfigured');
      setTabData(undefined);
      return;
    }

    let cancelled = false;
    setStatus('loading');

    const extraFields = props.showOverlay && props.overlaySourceField
      ? [props.overlaySourceField]
      : [];

    props.service
      .getTabData(props.siteUrl, props.listId, activeTab.viewId, extraFields)
      .then((data) => {
        if (cancelled) {
          return;
        }
        setTabData(data);
        setStatus(data.rows.length > 0 ? 'populated' : 'empty');
      })
      .catch((err: unknown) => {
        if (cancelled) {
          return;
        }
        console.warn('[TabbedListViews] Items load failed', err);
        setTabData(undefined);
        setStatus('error');
      });

    return () => {
      cancelled = true;
    };
  }, [
    props.service,
    props.siteUrl,
    props.listId,
    activeTab,
    props.showOverlay,
    props.overlaySourceField
  ]);

  const hasHeader = !!props.sectionTitle || !!props.seeAllUrl;

  return (
    <section className={styles.container}>
      {hasHeader && (
        <div className={styles.header}>
          {props.sectionTitle && <h2 className={styles.title}>{props.sectionTitle}</h2>}
          {props.seeAllUrl && (
            <a
              className={styles.seeAll}
              href={props.seeAllUrl}
              target="_blank"
              rel="noopener noreferrer"
            >
              SEE ALL
            </a>
          )}
        </div>
      )}

      {visibleTabs.length > 0 && (
        <nav className={styles.tabStrip} role="tablist">
          {visibleTabs.map((tab, idx) => (
            <button
              key={`${tab.label}-${idx}`}
              role="tab"
              type="button"
              aria-selected={idx === activeIdx}
              className={
                idx === activeIdx
                  ? `${styles.tab} ${styles.tabActive}`
                  : styles.tab
              }
              onClick={() => setActiveIdx(idx)}
            >
              {tab.label || `Tab ${idx + 1}`}
            </button>
          ))}
        </nav>
      )}

      <div className={styles.content}>
        {status === 'unconfigured' && (
          <div className={styles.message}>
            {props.listId
              ? 'Open the property pane to finish configuring this tab.'
              : 'Open the property pane to pick a list and configure tabs.'}
          </div>
        )}
        {status === 'loading' && <div className={styles.message}>Loading…</div>}
        {status === 'error' && (
          <div className={styles.message}>Couldn&apos;t load items.</div>
        )}
        {status === 'empty' && (
          <div className={styles.message}>No items to display.</div>
        )}
        {status === 'populated' && tabData && (
          props.layout === 'table' ? (
            <TableLayout data={tabData} />
          ) : (
            <GalleryLayout
              data={tabData}
              showOverlay={props.showOverlay}
              overlaySourceField={props.overlaySourceField}
              overlayLabelTemplate={props.overlayLabelTemplate}
              overlayPosition={props.overlayPosition}
            />
          )
        )}
      </div>
    </section>
  );
};
