import * as React from 'react';
import * as ReactDom from 'react-dom';

import styles from './MediaCardModal.module.scss';
import { VideoPlayer } from './VideoPlayer';
import { IParsedVideo } from '../services/videoSource';

export interface IMediaCardModalProps {
  video: IParsedVideo;
  title: string;
  closeAriaLabel: string;
  onClose: () => void;
}

export const MediaCardModal: React.FC<IMediaCardModalProps> = ({
  video,
  title,
  closeAriaLabel,
  onClose
}) => {
  const closeButtonRef = React.useRef<HTMLButtonElement>(null);
  const previouslyFocusedRef = React.useRef<Element | null>(null);

  React.useEffect(() => {
    // Remember whatever was focused before the modal opened so we can put
    // focus back when it closes (standard dialog behavior — without this the
    // user's tab position vanishes into document.body).
    previouslyFocusedRef.current = document.activeElement;
    closeButtonRef.current?.focus();

    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
    };
    document.addEventListener('keydown', onKeyDown);

    return () => {
      document.removeEventListener('keydown', onKeyDown);
      const prior = previouslyFocusedRef.current;
      if (prior && prior instanceof HTMLElement) {
        prior.focus();
      }
    };
  }, [onClose]);

  // Backdrop click closes only when the click landed on the backdrop itself —
  // clicks bubbled up from the modal body (e.g. interacting with video
  // controls) shouldn't dismiss the player.
  const onBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) {
      onClose();
    }
  };

  // Render into document.body via createPortal so the modal isn't clipped by
  // any ancestor with overflow:hidden / transform / filter (SharePoint shells
  // routinely apply these to their canvas containers).
  const node = (
    <div
      className={styles.backdrop}
      role="dialog"
      aria-modal="true"
      aria-label={title}
      onClick={onBackdropClick}
    >
      <div className={styles.modal}>
        <button
          ref={closeButtonRef}
          type="button"
          className={styles.closeButton}
          aria-label={closeAriaLabel}
          onClick={onClose}
        >
          {/* Multiplication sign reads cleaner than X at this size */}
          ×
        </button>
        <div className={styles.playerSlot}>
          <VideoPlayer video={video} title={title} />
        </div>
      </div>
    </div>
  );

  return ReactDom.createPortal(node, document.body);
};
