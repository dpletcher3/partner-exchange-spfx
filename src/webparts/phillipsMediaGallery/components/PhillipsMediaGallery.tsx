import * as React from 'react';
import { HttpClient } from '@microsoft/sp-http';

import styles from './PhillipsMediaGallery.module.scss';
import { PhillipsMediaGalleryService } from '../services/PhillipsMediaGalleryService';
import { IFieldMapping, IMediaCardItem } from '../services/models';
import { MediaGrid } from './MediaGrid';
import { LoadingState } from './LoadingState';
import { EmptyState } from './EmptyState';
import { ErrorState } from './ErrorState';

const LOG = '[MediaGallery]';

export interface IPhillipsMediaGalleryProps {
  service: PhillipsMediaGalleryService;
  httpClient: HttpClient;
  siteUrl: string;
  listId: string;
  mapping: IFieldMapping;
  columns: number;
  sectionTitle: string;
  openInNewTab: boolean;
}

type Status = 'loading' | 'empty' | 'error' | 'loaded';

const SKELETON_COUNT = 8;

export const PhillipsMediaGallery: React.FC<IPhillipsMediaGalleryProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');
  const [items, setItems] = React.useState<IMediaCardItem[]>([]);
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  const headerId = React.useMemo(
    () => `phil-mg-title-${Math.random().toString(36).slice(2)}`,
    []
  );

  // Re-run when the list or any mapped field changes.
  const m = props.mapping;
  const depsKey = `${props.listId}|${m.titleField}|${m.videoField}|${m.labelImageField}|${m.mainImageField}`;

  const load = React.useCallback(() => {
    if (!props.listId) {
      setItems([]);
      setStatus('empty');
      return;
    }
    setStatus('loading');
    props.service
      .getItems(props.siteUrl, props.listId, props.mapping)
      .then((result) => {
        console.log(`${LOG} fetched ${result.length} items`);
        if (!result.length) {
          setItems([]);
          setStatus('empty');
        } else {
          setItems(result);
          setStatus('loaded');
        }
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        console.warn(`${LOG} item fetch failed`, err);
        setErrorMessage(msg);
        setStatus('error');
      });
  }, [props.service, props.siteUrl, depsKey]);

  React.useEffect(() => {
    load();
  }, [load]);

  return (
    <section
      className={styles.mediaGallery}
      aria-labelledby={props.sectionTitle ? headerId : undefined}
    >
      {props.sectionTitle && (
        <div className={styles.header}>
          <h2 id={headerId} className={styles.title}>
            {props.sectionTitle}
          </h2>
        </div>
      )}

      {status === 'loading' && <LoadingState columns={props.columns} count={SKELETON_COUNT} />}
      {status === 'error' && <ErrorState message={errorMessage} onRetry={load} />}
      {status === 'empty' && (
        <EmptyState
          message={
            props.listId
              ? 'No items to show in this list yet.'
              : 'Select a list in the property pane to show media cards.'
          }
        />
      )}
      {status === 'loaded' && (
        <MediaGrid
          items={items}
          columns={props.columns}
          openInNewTab={props.openInNewTab}
          httpClient={props.httpClient}
        />
      )}
    </section>
  );
};
