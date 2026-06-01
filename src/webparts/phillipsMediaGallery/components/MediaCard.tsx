import * as React from 'react';
import { HttpClient } from '@microsoft/sp-http';
import styles from './PhillipsMediaGallery.module.scss';
import { IMediaCardItem } from '../services/models';
import { detectVideo, getYouTubeThumbnails } from '../services/videoThumbnail';
import { getVimeoThumbnail } from '../services/vimeoOEmbed';

const LOG = '[MediaGallery]';

export interface IMediaCardProps {
  item: IMediaCardItem;
  openInNewTab: boolean;
  httpClient: HttpClient;
}

type ThumbKind = 'override' | 'youtube' | 'vimeo' | 'placeholder';

interface IMainVisual {
  kind: ThumbKind;
  url?: string;
  hqFallback?: string; // youtube only — used by the <img onError> fallback
}

export const MediaCard: React.FC<IMediaCardProps> = ({ item, openInNewTab, httpClient }) => {
  const [visual, setVisual] = React.useState<IMainVisual>({ kind: 'placeholder' });
  const [triedHq, setTriedHq] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    setTriedHq(false);

    // §3 main-visual precedence.
    if (item.mainImageOverrideUrl) {
      // (a) explicit mapped override.
      console.log(`${LOG} card "${item.title}" thumbnail source: override`);
      setVisual({ kind: 'override', url: item.mainImageOverrideUrl });
    } else {
      const v = detectVideo(item.videoUrl);
      if (v.kind === 'youtube' && v.id) {
        // (b) YouTube — deterministic; maxres with hqdefault onError fallback.
        const t = getYouTubeThumbnails(v.id);
        console.log(`${LOG} card "${item.title}" thumbnail source: youtube`);
        setVisual({ kind: 'youtube', url: t.maxRes, hqFallback: t.hq });
      } else if (v.kind === 'vimeo') {
        // (b) Vimeo — async oEmbed (cached); placeholder until it resolves.
        setVisual({ kind: 'placeholder' });
        getVimeoThumbnail(item.videoUrl, httpClient)
          .then((url) => {
            if (cancelled) {
              return;
            }
            if (url) {
              console.log(`${LOG} card "${item.title}" thumbnail source: vimeo`);
              setVisual({ kind: 'vimeo', url });
            } else {
              console.log(
                `${LOG} card "${item.title}" thumbnail source: placeholder (vimeo oEmbed unavailable)`
              );
              setVisual({ kind: 'placeholder' });
            }
          })
          .catch(() => {
            if (!cancelled) {
              setVisual({ kind: 'placeholder' });
            }
          });
      } else {
        // (c) neutral placeholder (no video, or unrecognized URL).
        console.log(`${LOG} card "${item.title}" thumbnail source: placeholder`);
        setVisual({ kind: 'placeholder' });
      }
    }

    return () => {
      cancelled = true;
    };
  }, [item.id, item.title, item.videoUrl, item.mainImageOverrideUrl, httpClient]);

  const onImgError = (): void => {
    if (visual.kind === 'youtube' && visual.hqFallback && !triedHq) {
      setTriedHq(true);
    }
  };

  const imgSrc = visual.kind === 'youtube' && triedHq ? visual.hqFallback : visual.url;

  const inner = (
    <>
      <div className={styles.mainVisual}>
        {imgSrc ? (
          <img className={styles.mainVisualImg} src={imgSrc} alt="" onError={onImgError} />
        ) : (
          <div className={`${styles.placeholderFill} ${styles.placeholder}`} aria-hidden="true" />
        )}
      </div>
      <div className={styles.labelStrip}>
        {item.labelImageUrl ? (
          <img className={styles.labelImage} src={item.labelImageUrl} alt={item.title} />
        ) : (
          <span className={styles.labelText}>{item.title}</span>
        )}
      </div>
    </>
  );

  // Whole card is the click target opening the video; non-clickable with no video.
  if (item.videoUrl) {
    return (
      <a
        className={styles.card}
        href={item.videoUrl}
        target={openInNewTab ? '_blank' : '_self'}
        rel="noopener noreferrer"
        aria-label={item.title}
      >
        {inner}
      </a>
    );
  }
  return <div className={styles.card}>{inner}</div>;
};
