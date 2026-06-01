import * as React from 'react';
import styles from './PhillipsHighlightVideo.module.scss';

export interface IPhillipsHighlightVideoProps {
  hasList: boolean;
  hasItem: boolean;
}

type Status = 'configure' | 'loading' | 'error' | 'loaded';

export const PhillipsHighlightVideo: React.FC<IPhillipsHighlightVideoProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');

  // Turn 1: no data fetch yet. A list + item selection flips the block from the
  // configure prompt to the placeholder shell. Turn 2 replaces this with the
  // real item read (driving loading → loaded / error) and the inline Vimeo embed.
  React.useEffect(() => {
    setStatus(props.hasList && props.hasItem ? 'loaded' : 'configure');
  }, [props.hasList, props.hasItem]);

  return (
    <section className={styles.highlightVideo}>
      <div className={styles.block}>
        {status === 'configure' && (
          <div className={styles.message}>
            Select a list and an item in the property pane to feature a video.
          </div>
        )}
        {status === 'loading' && <div className={styles.message}>Loading…</div>}
        {status === 'error' && (
          <div className={styles.message} role="alert">
            Couldn’t load this highlight.
          </div>
        )}
        {status === 'loaded' && (
          <>
            <h2 className={styles.title}>Highlight title loads here</h2>
            <div className={styles.playerFrame}>
              <div className={styles.playerPlaceholder} aria-hidden="true">
                <span className={styles.playGlyph}>▶</span>
                <span>Video loads here</span>
              </div>
            </div>
            <div className={styles.info}>Practice info loads here.</div>
          </>
        )}
      </div>
    </section>
  );
};
