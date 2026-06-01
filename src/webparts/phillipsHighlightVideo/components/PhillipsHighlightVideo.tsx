import * as React from 'react';
import styles from './PhillipsHighlightVideo.module.scss';
import { HighlightVideoService } from '../services/HighlightVideoService';
import { IFieldMapping, IHighlightItem } from '../services/models';
// Reuse the Media Card Gallery's Vimeo parser — do not re-implement (D035 / spec §3).
import { detectVideo } from '../../phillipsMediaGallery/services/videoThumbnail';

const LOG = '[HighlightVideo]';

export interface IPhillipsHighlightVideoProps {
  service: HighlightVideoService;
  siteUrl: string;
  listId: string;
  itemId: number;
  mapping: IFieldMapping;
}

type Status = 'configure' | 'loading' | 'error' | 'loaded';

export const PhillipsHighlightVideo: React.FC<IPhillipsHighlightVideoProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('configure');
  const [item, setItem] = React.useState<IHighlightItem | undefined>(undefined);
  const [vimeoId, setVimeoId] = React.useState<string>('');
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  const m = props.mapping;
  const depsKey = `${props.listId}|${props.itemId}|${m.titleField}|${m.videoField}|${m.infoField}`;

  const load = React.useCallback(() => {
    if (!props.listId || !props.itemId) {
      setStatus('configure');
      return;
    }
    setStatus('loading');
    props.service
      .getItem(props.siteUrl, props.listId, props.itemId, props.mapping)
      .then((resolved) => {
        const v = detectVideo(resolved.videoUrl);
        console.log(`${LOG} resolved item ${resolved.id} "${resolved.title}"; video="${resolved.videoUrl}"; parsed=`, v);
        setItem(resolved);
        if (v.kind === 'vimeo' && v.id) {
          const src = `https://player.vimeo.com/video/${v.id}`;
          console.log(`${LOG} iframe src: ${src}`);
          setVimeoId(v.id);
          setStatus('loaded');
        } else {
          console.warn(`${LOG} no parseable Vimeo id from "${resolved.videoUrl}" — error state`);
          setVimeoId('');
          setErrorMessage(
            resolved.videoUrl
              ? 'The featured item’s video URL is not a recognized Vimeo link.'
              : 'The featured item has no video URL.'
          );
          setStatus('error');
        }
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        console.warn(`${LOG} item fetch failed`, err);
        setErrorMessage(msg);
        setStatus('error');
      });
  }, [props.service, props.siteUrl, depsKey]);

  React.useEffect(() => {
    load();
  }, [load]);

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
            {errorMessage || 'Couldn’t load this highlight.'}
            {item && item.videoUrl && (
              <div className={styles.fallbackWrap}>
                <a
                  className={styles.openLink}
                  href={item.videoUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Open the video in a new tab ↗
                </a>
              </div>
            )}
          </div>
        )}

        {status === 'loaded' && item && (
          <>
            {item.title && <h2 className={styles.title}>{item.title}</h2>}
            <div className={styles.playerFrame}>
              <iframe
                className={styles.playerIframe}
                src={`https://player.vimeo.com/video/${vimeoId}`}
                title={item.title || 'Featured video'}
                allow="autoplay; fullscreen; picture-in-picture"
                allowFullScreen
              />
            </div>
            {/* Fallback per spec §4: if the inline frame is CSP/privacy-blocked,
                this link still opens the video. */}
            <div className={styles.fallbackWrap}>
              <a
                className={styles.openLink}
                href={item.videoUrl}
                target="_blank"
                rel="noopener noreferrer"
              >
                Open in Vimeo ↗
              </a>
            </div>
            {item.info && <div className={styles.info}>{item.info}</div>}
          </>
        )}
      </div>
    </section>
  );
};
