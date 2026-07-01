import * as React from 'react';

import styles from './PhillipsListView.module.scss';
import { GalleryLayout } from './GalleryLayout';
import { TableLayout } from './TableLayout';
import { IPhillipsListViewService } from '../services/IPhillipsListViewService';
import { ITabData, Layout, OverlayPosition } from '../services/models';

// Renderer: loads data for a single view and dispatches to the Gallery or
// Table layout. Decoupled from the tab strip so the orchestrator can drive it
// either from a tab selection (when "Show tabs" is on) or directly from a
// single configured viewId (when "Show tabs" is off).
type Status = 'loading' | 'populated' | 'empty' | 'error' | 'unconfigured';

export interface IListViewContentProps {
  service: IPhillipsListViewService;
  siteUrl: string;
  listId: string;
  viewId: string;
  layout: Layout;
  cardFieldCount: number;
  showOverlay: boolean;
  overlaySourceField: string;
  overlayLabelTemplate: string;
  overlayPosition: OverlayPosition;
  // Message shown when listId or viewId isn't configured yet. Caller picks the
  // wording (tabs-mode vs single-view mode) so the editor sees exactly what's
  // missing.
  unconfiguredMessage: string;
}

export const ListViewContent: React.FC<IListViewContentProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');
  const [data, setData] = React.useState<ITabData | undefined>(undefined);

  React.useEffect(() => {
    if (!props.listId || !props.viewId) {
      setStatus('unconfigured');
      setData(undefined);
      return;
    }

    let cancelled = false;
    setStatus('loading');

    const extraFields = props.showOverlay && props.overlaySourceField
      ? [props.overlaySourceField]
      : [];

    props.service
      .getTabData(props.siteUrl, props.listId, props.viewId, extraFields)
      .then((loaded) => {
        if (cancelled) {
          return;
        }
        setData(loaded);
        setStatus(loaded.rows.length > 0 ? 'populated' : 'empty');
      })
      .catch((err: unknown) => {
        if (cancelled) {
          return;
        }
        console.warn('[PhillipsListView] Items load failed', err);
        setData(undefined);
        setStatus('error');
      });

    return () => {
      cancelled = true;
    };
  }, [
    props.service,
    props.siteUrl,
    props.listId,
    props.viewId,
    props.showOverlay,
    props.overlaySourceField
  ]);

  if (status === 'unconfigured') {
    return <div className={styles.message}>{props.unconfiguredMessage}</div>;
  }
  if (status === 'loading') {
    return <div className={styles.message}>Loading…</div>;
  }
  if (status === 'error') {
    return <div className={styles.message}>Couldn&apos;t load items.</div>;
  }
  if (status === 'empty') {
    return <div className={styles.message}>No items to display.</div>;
  }
  if (status === 'populated' && data) {
    return props.layout === 'table' ? (
      <TableLayout data={data} />
    ) : (
      <GalleryLayout
        data={data}
        cardFieldCount={props.cardFieldCount}
        showOverlay={props.showOverlay}
        overlaySourceField={props.overlaySourceField}
        overlayLabelTemplate={props.overlayLabelTemplate}
        overlayPosition={props.overlayPosition}
      />
    );
  }
  return null;
};
