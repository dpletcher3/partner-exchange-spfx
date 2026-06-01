import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';

export interface IEmptyStateProps {
  message: string;
}

// Shown when no list is selected (Turn 1 prompt) or the selected list returns
// zero items. The message distinguishes the two.
export const EmptyState: React.FC<IEmptyStateProps> = ({ message }) => {
  return <div className={styles.message}>{message}</div>;
};
