import * as React from 'react';

import styles from './MediaCardImage.module.scss';

export interface IMediaCardImageProps {
  imageUrl: string;
  imageAlt: string;
  // When provided, the image becomes a click-to-play target and the play
  // overlay is rendered on top. When undefined, the image is purely
  // decorative.
  onPlayClick: (() => void) | undefined;
  playAriaLabel: string;
}

export const MediaCardImage: React.FC<IMediaCardImageProps> = ({
  imageUrl,
  imageAlt,
  onPlayClick,
  playAriaLabel
}) => {
  return (
    <div className={styles.imageWrap}>
      <img className={styles.image} src={imageUrl} alt={imageAlt} />
      {onPlayClick && (
        <button
          type="button"
          className={styles.playButton}
          aria-label={playAriaLabel}
          onClick={onPlayClick}
        >
          <div className={styles.playCircle}>
            {/* viewBox sized to the triangle's bounds so the SVG scales with
                the circle if we ever bump the play-icon size. */}
            <svg
              className={styles.playIcon}
              viewBox="0 0 22 22"
              aria-hidden="true"
              focusable="false"
            >
              <polygon points="5,3 19,11 5,19" />
            </svg>
          </div>
        </button>
      )}
    </div>
  );
};
