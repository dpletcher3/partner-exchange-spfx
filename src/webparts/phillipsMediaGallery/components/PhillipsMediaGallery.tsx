import * as React from 'react';

import styles from './PhillipsMediaGallery.module.scss';
import { MediaGrid } from './MediaGrid';
import { LoadingState } from './LoadingState';
import { EmptyState } from './EmptyState';
import { ErrorState } from './ErrorState';

export interface IPhillipsMediaGalleryProps {
  columns: number;
  sectionTitle: string;
  // True once a list is selected in the property pane. Turn 1 has no data layer,
  // so this is the only signal that drives the placeholder/empty distinction.
  hasList: boolean;
}

type Status = 'loading' | 'empty' | 'error' | 'loaded';

// How many placeholder cards to render in the "loaded" state until Turn 2 wires
// the real list service.
const PLACEHOLDER_COUNT = 8;

export const PhillipsMediaGallery: React.FC<IPhillipsMediaGalleryProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');

  // Stable per-instance id so the <h2> can be referenced by aria-labelledby.
  const headerId = React.useMemo(
    () => `phil-mg-title-${Math.random().toString(36).slice(2)}`,
    []
  );

  // Turn 1 placeholder behavior: no real fetch yet. A list selection flips the
  // gallery from the empty prompt to placeholder cards. Turn 2 replaces this
  // effect with the list service that drives loading → loaded / empty / error.
  React.useEffect(() => {
    setStatus(props.hasList ? 'loaded' : 'empty');
  }, [props.hasList]);

  // Turn 2 will re-run the (not-yet-built) fetch. For Turn 1 this just re-derives
  // the placeholder/empty status.
  const retry = React.useCallback(() => {
    setStatus(props.hasList ? 'loaded' : 'empty');
  }, [props.hasList]);

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

      {status === 'loading' && <LoadingState columns={props.columns} count={PLACEHOLDER_COUNT} />}
      {status === 'error' && <ErrorState message="" onRetry={retry} />}
      {status === 'empty' && <EmptyState />}
      {status === 'loaded' && (
        <MediaGrid columns={props.columns} placeholderCount={PLACEHOLDER_COUNT} />
      )}
    </section>
  );
};
