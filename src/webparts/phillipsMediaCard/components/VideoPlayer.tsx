import * as React from 'react';

import { IParsedVideo } from '../services/videoSource';

export interface IVideoPlayerProps {
  video: IParsedVideo;
  // Title for the iframe / accessible label for the HTML5 video. Re-uses the
  // card's header so screen readers get a meaningful name.
  title: string;
}

// Dispatches to the right embed shape based on the parsed source type. The
// modal that wraps this component handles sizing and aspect ratio; here we
// just emit the right element and let CSS size it via 100% width/height.
export const VideoPlayer: React.FC<IVideoPlayerProps> = ({ video, title }) => {
  if (video.type === 'sharepoint') {
    return (
      <video
        src={video.embedUrl}
        controls
        autoPlay
        playsInline
        title={title}
        style={{ width: '100%', height: '100%', backgroundColor: '#000' }}
      />
    );
  }

  return (
    <iframe
      src={video.embedUrl}
      title={title}
      allow="autoplay; encrypted-media; picture-in-picture"
      allowFullScreen
      style={{ width: '100%', height: '100%', border: 'none', backgroundColor: '#000' }}
    />
  );
};
