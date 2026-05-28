import * as React from 'react';

import styles from './PhillipsListView.module.scss';
import { ListViewContent } from './ListViewContent';
import { IPhillipsListViewService } from '../services/IPhillipsListViewService';
import { ITabConfig, Layout, OverlayPosition } from '../services/models';

export interface IPhillipsListViewProps {
  service: IPhillipsListViewService;
  siteUrl: string;
  sectionTitle: string;
  listId: string;
  layout: Layout;
  // Tab strip visibility. When false, the strip is hidden entirely and the
  // renderer is driven by the single configured viewId; when true, the strip
  // exposes tabCount tabs and the active tab's viewId drives the renderer.
  showTabs: boolean;
  // Used when showTabs is false.
  viewId: string;
  // Used when showTabs is true.
  tabCount: number;
  tabs: ITabConfig[];
  // Common across both modes.
  seeAllUrl: string;
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
}

export const PhillipsListView: React.FC<IPhillipsListViewProps> = (props) => {
  // When showTabs is false the tab list is irrelevant; keep it empty so the
  // active-index effect and tab strip render are no-ops in that mode.
  const visibleTabs = React.useMemo(
    () => (props.showTabs ? (props.tabs || []).slice(0, props.tabCount) : []),
    [props.showTabs, props.tabs, props.tabCount]
  );

  const [activeIdx, setActiveIdx] = React.useState<number>(0);

  // Clamp activeIdx if a tab was deleted from underneath us.
  React.useEffect(() => {
    if (visibleTabs.length > 0 && activeIdx >= visibleTabs.length) {
      setActiveIdx(0);
    }
  }, [visibleTabs.length, activeIdx]);

  const activeViewId: string = props.showTabs
    ? visibleTabs[activeIdx]?.viewId || ''
    : props.viewId;

  const hasHeader = !!props.sectionTitle || !!props.seeAllUrl;

  // Different unconfigured-message text per mode helps the editor see exactly
  // what's missing in the property pane.
  const unconfiguredMessage: string = !props.listId
    ? props.showTabs
      ? 'Open the property pane to pick a list and configure tabs.'
      : 'Open the property pane to pick a list and view.'
    : props.showTabs
      ? 'Open the property pane to finish configuring this tab.'
      : 'Open the property pane to pick a view.';

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

      {props.showTabs && visibleTabs.length > 0 && (
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
        <ListViewContent
          service={props.service}
          siteUrl={props.siteUrl}
          listId={props.listId}
          viewId={activeViewId}
          layout={props.layout}
          showOverlay={props.showOverlay}
          overlaySourceField={props.overlaySourceField}
          overlayLabelTemplate={props.overlayLabelTemplate}
          overlayPosition={props.overlayPosition}
          unconfiguredMessage={unconfiguredMessage}
        />
      </div>
    </section>
  );
};
