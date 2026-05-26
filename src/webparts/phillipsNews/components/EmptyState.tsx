import * as React from 'react';
import styles from './PhillipsNews.module.scss';

// Shown when the query returns zero items (typically because filters exclude
// everything).
export const EmptyState: React.FC = () => {
  return <div className={styles.message}>No news to show right now.</div>;
};
