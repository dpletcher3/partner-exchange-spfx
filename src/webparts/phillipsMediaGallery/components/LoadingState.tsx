import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';

export interface ILoadingStateProps {
  columns: number;
  count: number;
}

// Skeleton grid shown during the initial fetch (Turn 2). The footprint mirrors
// the loaded grid so the section doesn't reflow when real cards arrive.
export const LoadingState: React.FC<ILoadingStateProps> = ({ columns, count }) => {
  const gridStyle = { ['--phil-mg-cols']: String(columns) } as React.CSSProperties;
  return (
    <div className={styles.grid} style={gridStyle} aria-busy="true">
      {Array.from({ length: count }).map((_, i) => (
        <div key={`skeleton-${i}`} className={styles.card} aria-hidden="true">
          <div className={`${styles.mainVisual} ${styles.placeholder}`} />
          <div className={styles.labelStrip}>
            <span className={`${styles.placeholderLine} ${styles.placeholder}`} />
          </div>
        </div>
      ))}
    </div>
  );
};
