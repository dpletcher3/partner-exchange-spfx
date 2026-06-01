import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';

// Shown until a list is selected (Turn 1) and, in Turn 2, when the selected list
// returns zero items.
export const EmptyState: React.FC = () => {
  return (
    <div className={styles.message}>
      Select a list in the property pane to show media cards.
    </div>
  );
};
