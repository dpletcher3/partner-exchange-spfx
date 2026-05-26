import * as React from 'react';
import styles from './PhillipsNews.module.scss';

export interface IErrorStateProps {
  message: string;
  onRetry: () => void;
}

// Shown on REST failure. Surfaces the error message and a retry button that
// re-runs the query.
export const ErrorState: React.FC<IErrorStateProps> = ({ message, onRetry }) => {
  return (
    <div className={styles.message} role="alert">
      <div>Couldn’t load news. {message}</div>
      <button type="button" className={styles.retryButton} onClick={onRetry}>
        Retry
      </button>
    </div>
  );
};
