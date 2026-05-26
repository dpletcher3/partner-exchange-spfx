import * as React from 'react';
import styles from './PhillipsNews.module.scss';

export interface ILoadingStateProps {
  count: number;
}

// Skeleton grid shown during the initial fetch and on property changes. The
// skeleton card footprint mirrors the populated layout so the section doesn't
// reflow when real items arrive.
export const LoadingState: React.FC<ILoadingStateProps> = ({ count }) => {
  return (
    <div className={styles.grid} aria-busy="true">
      {Array.from({ length: count }).map((_, i) => (
        <div key={`skeleton-${i}`} className={styles.skeletonCard} aria-hidden="true">
          <div className={styles.skeletonThumb} />
          <div className={styles.skeletonBody}>
            <div className={`${styles.skeletonLine} ${styles.skeletonLineShort}`} />
            <div className={styles.skeletonLine} />
            <div className={styles.skeletonLine} />
          </div>
        </div>
      ))}
    </div>
  );
};
