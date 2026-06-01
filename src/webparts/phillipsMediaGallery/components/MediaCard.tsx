import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';

// Turn 1: placeholder-only card. The footprint (16:9 main visual + label strip
// below) matches the final card so the grid won't reflow when Turn 2 supplies
// real items, thumbnails, and label graphics.
export const MediaCard: React.FC = () => {
  return (
    <div className={styles.card} aria-hidden="true">
      <div className={`${styles.mainVisual} ${styles.placeholder}`} />
      <div className={styles.labelStrip}>
        <span className={`${styles.placeholderLine} ${styles.placeholder}`} />
      </div>
    </div>
  );
};
