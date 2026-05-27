import * as React from 'react';

import styles from './PhillipsPersonalizedHero.module.scss';
import { Greeting } from './Greeting';
import { DateTime } from './DateTime';
import { computeGreeting, extractFirstName } from '../utils/timeOfDay';

export type BackgroundType = 'Color' | 'Image';
export type GreetingAlignment = 'Left' | 'Center' | 'Right';

export interface IPhillipsPersonalizedHeroProps {
  backgroundType: BackgroundType;
  backgroundColor: string;
  backgroundImage: string;
  bannerHeight: number;
  greetingColor: string;
  greetingSize: number;
  greetingWeight: number;
  greetingAlignment: GreetingAlignment;
  displayName?: string;
  isEditMode: boolean;
}

const GREETING_ID = 'phil-hero-greeting';

const ALIGNMENT_CLASS: Record<GreetingAlignment, string> = {
  Left: styles.alignLeft,
  Center: styles.alignCenter,
  Right: styles.alignRight
};

export const PhillipsPersonalizedHero: React.FC<IPhillipsPersonalizedHeroProps> = (props) => {
  // Computed once per render; no timer (refresh updates after a boundary cross).
  const phrase = computeGreeting(new Date());
  const firstName = extractFirstName(props.displayName);

  const useImage = props.backgroundType === 'Image' && !!props.backgroundImage;
  // Image chosen but none picked: prompt in edit mode; fall back to color at runtime.
  const showEditPlaceholder =
    props.backgroundType === 'Image' && !props.backgroundImage && props.isEditMode;

  // backgroundColor is always the base layer so the banner is never blank — if an
  // image is set it's layered on top (cover/center via .hero), and a missing or
  // broken image simply reveals the color underneath.
  const sectionStyle: React.CSSProperties = {
    height: `${props.bannerHeight}px`,
    backgroundColor: props.backgroundColor
  };
  if (useImage) {
    sectionStyle.backgroundImage = `url("${props.backgroundImage}")`;
  }

  const alignmentClass = ALIGNMENT_CLASS[props.greetingAlignment] || styles.alignLeft;

  return (
    <section
      className={`${styles.hero} ${alignmentClass}`}
      style={sectionStyle}
      role="banner"
      aria-labelledby={showEditPlaceholder ? undefined : GREETING_ID}
    >
      {showEditPlaceholder ? (
        <div className={styles.editPlaceholder}>
          Pick a background image in the property pane
        </div>
      ) : (
        <>
          <Greeting
            id={GREETING_ID}
            phrase={phrase}
            firstName={firstName}
            color={props.greetingColor}
            sizePx={props.greetingSize}
            weight={props.greetingWeight}
          />
          <DateTime color={props.greetingColor} />
        </>
      )}
    </section>
  );
};
