import * as React from 'react';

import styles from './PhillipsMediaCard.module.scss';
import { MediaCardImage } from './MediaCardImage';
import { MediaCardModal } from './MediaCardModal';
import { parseVideoUrl, VideoSourceType, IParsedVideo } from '../services/videoSource';
import { getYouTubeThumbnail } from '../services/videoThumbnail';

export interface IPhillipsMediaCardProps {
  eyebrow: string;
  header: string;
  description: string;
  // Uploaded image URL — always wins over auto-derived thumbnails when set.
  imageUrl: string;
  imageAlt: string;
  videoUrl: string;
  videoSourceType: VideoSourceType;
  // Resolved Vimeo thumbnail URL — populated by the WebPart's oEmbed fetch.
  // Empty string when the fetch hasn't run or failed.
  vimeoThumbnailUrl: string;
  unconfiguredMessage: string;
  playAriaLabel: string;
  closeAriaLabel: string;
}

export const PhillipsMediaCard: React.FC<IPhillipsMediaCardProps> = (props) => {
  const [modalOpen, setModalOpen] = React.useState<boolean>(false);

  // Parse video config once per props change. If the editor set a YouTube /
  // Vimeo URL that doesn't match a known shape, parsed is undefined and we
  // fall back to opening the URL in a new tab instead of rendering a broken
  // iframe.
  const parsed: IParsedVideo | undefined = React.useMemo(
    () => (props.videoUrl ? parseVideoUrl(props.videoUrl, props.videoSourceType) : undefined),
    [props.videoUrl, props.videoSourceType]
  );

  // Image resolution priority (first non-empty wins):
  //   1. Uploaded image — always overrides auto-derived thumbnails so an
  //      editor can replace a bad YouTube/Vimeo poster.
  //   2. Vimeo thumbnail — populated by the WebPart's oEmbed fetch at edit
  //      time and persisted in properties; rendered without a runtime call.
  //   3. YouTube derived URL — synchronous derivation from the parsed
  //      video ID at render time.
  //   4. Nothing — falls through to the unconfigured branch below.
  const resolvedImageUrl: string = React.useMemo(() => {
    if (props.imageUrl) {
      return props.imageUrl;
    }
    if (props.videoSourceType === 'vimeo' && props.vimeoThumbnailUrl) {
      return props.vimeoThumbnailUrl;
    }
    if (props.videoSourceType === 'youtube' && parsed?.type === 'youtube' && parsed.videoId) {
      return getYouTubeThumbnail(parsed.videoId);
    }
    return '';
  }, [props.imageUrl, props.videoSourceType, props.vimeoThumbnailUrl, parsed]);

  // Unconfigured state: header + description always required; a resolved
  // image (uploaded OR auto-derived) is also required. Editors see this when
  // a Vimeo oEmbed call fails before they've uploaded a manual fallback.
  if (!props.header || !props.description || !resolvedImageUrl) {
    return <div className={styles.unconfigured}>{props.unconfiguredMessage}</div>;
  }

  // onPlayClick is undefined when there's no video URL — MediaCardImage uses
  // that to decide whether to render the play overlay at all.
  const onPlayClick: (() => void) | undefined = props.videoUrl
    ? () => {
        if (parsed) {
          setModalOpen(true);
        } else {
          // Parser couldn't extract a YouTube / Vimeo ID — fall back to
          // opening the raw URL in a new tab. noopener for the usual
          // tabnabbing protection.
          window.open(props.videoUrl, '_blank', 'noopener,noreferrer');
        }
      }
    : undefined;

  return (
    <section className={styles.container}>
      <div className={styles.textColumn}>
        {props.eyebrow && <div className={styles.eyebrow}>{props.eyebrow}</div>}
        <h2 className={styles.header}>{props.header}</h2>
        <p className={styles.description}>{props.description}</p>
      </div>
      <div className={styles.imageColumn}>
        <MediaCardImage
          imageUrl={resolvedImageUrl}
          imageAlt={props.imageAlt || props.header}
          onPlayClick={onPlayClick}
          playAriaLabel={props.playAriaLabel}
        />
      </div>
      {modalOpen && parsed && (
        <MediaCardModal
          video={parsed}
          title={props.header}
          closeAriaLabel={props.closeAriaLabel}
          onClose={() => setModalOpen(false)}
        />
      )}
    </section>
  );
};
