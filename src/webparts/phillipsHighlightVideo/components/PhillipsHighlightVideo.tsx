import * as React from 'react';
import styles from './PhillipsHighlightVideo.module.scss';
import { HighlightVideoService } from '../services/HighlightVideoService';
import { IFieldMapping, IHighlightItem } from '../services/models';
// Provider dispatch (Vimeo or YouTube). Thin wrapper over the Media Card
// Gallery's parser — do not re-implement (D035 / spec §3).
import { resolveVideoEmbed, IVideoEmbed } from '../services/videoEmbed';

const LOG = '[HighlightVideo]';

export interface IPhillipsHighlightVideoProps {
  service: HighlightVideoService;
  siteUrl: string;
  listId: string;
  itemId: number;
  mapping: IFieldMapping;
  // Optional section header above the title (e.g. "Practice of the Month").
  // Empty string = no header (the part stays reusable without one).
  sectionHeader: string;
}

type Status = 'configure' | 'loading' | 'error' | 'loaded';

export const PhillipsHighlightVideo: React.FC<IPhillipsHighlightVideoProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('configure');
  const [item, setItem] = React.useState<IHighlightItem | undefined>(undefined);
  const [embed, setEmbed] = React.useState<IVideoEmbed | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  // Stable id linking the <section> to its header for aria-labelledby (mirrors
  // the Media Card Gallery's section-header pattern).
  const headerId = React.useMemo(
    () => `phil-hv-header-${Math.random().toString(36).slice(2)}`,
    []
  );

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
        const v = resolveVideoEmbed(resolved.videoUrl);
        console.log(`${LOG} resolved item ${resolved.id} "${resolved.title}"; video="${resolved.videoUrl}"; parsed=`, v);
        setItem(resolved);
        if (v) {
          console.log(`${LOG} ${v.provider} iframe src: ${v.embedSrc}`);
          setEmbed(v);
          setStatus('loaded');
        } else {
          console.warn(`${LOG} no parseable video id from "${resolved.videoUrl}" — error state`);
          setEmbed(undefined);
          setErrorMessage(
            resolved.videoUrl
              ? 'The featured item’s video URL is not a recognized Vimeo or YouTube link.'
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
    <section
      className={styles.highlightVideo}
      aria-labelledby={props.sectionHeader ? headerId : undefined}
    >
      <div className={styles.block}>
        {/* Section header above the title. Rendered only when non-empty so the
            part stays reusable without one. Mirrors the Media Card Gallery. */}
        {props.sectionHeader && (
          <div className={styles.sectionHeader}>
            <h2 id={headerId} className={styles.sectionTitle}>
              {props.sectionHeader}
            </h2>
          </div>
        )}

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

        {status === 'loaded' && item && embed && (
          <>
            {item.title && <h2 className={styles.title}>{item.title}</h2>}
            <div className={styles.playerFrame}>
              <iframe
                className={styles.playerIframe}
                src={embed.embedSrc}
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
                Open in {embed.providerLabel} ↗
              </a>
            </div>
            {item.info && <div className={styles.info}>{item.info}</div>}
          </>
        )}
      </div>
    </section>
  );
};
