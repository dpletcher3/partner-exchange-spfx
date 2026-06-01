import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';

export interface IErrorStateProps {
  message: string;
  onRetry: () => void;
}

// Shown on a list-read failure (Turn 2). Surfaces the error and a retry button.
export const ErrorState: React.FC<IErrorStateProps> = ({ message, onRetry }) => {
  return (
    <div className={styles.message} role="alert">
      <div>Couldn’t load this gallery.{message ? ` ${message}` : ''}</div>
      <button type="button" className={styles.retryButton} onClick={onRetry}>
        Retry
      </button>
    </div>
  );
};
