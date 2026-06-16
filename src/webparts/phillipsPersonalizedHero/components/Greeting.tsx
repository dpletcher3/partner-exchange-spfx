import * as React from 'react';
import styles from './PhillipsPersonalizedHero.module.scss';
import { GreetingAlignment } from './PhillipsPersonalizedHero';

// Side padding the greeting anchors against, matching the hero's $phil-space-6
// (32px) horizontal padding so the absolutely-positioned greeting lines up with
// the in-flow date/time below it.
const SIDE_PADDING = '32px';

export interface IGreetingProps {
  // Time-of-day phrase, e.g. "Good morning".
  phrase: string;
  // First name, or undefined when none could be derived.
  firstName?: string;
  color: string;
  sizePx: number;
  weight: number;
  id: string;
  // Percent from the top of the hero (0–100); the greeting is absolutely
  // positioned at this `top`. Overlap at high values is allowed by design.
  verticalPositionPct: number;
  // Horizontal anchor, since the greeting is out of the flex flow.
  alignment: GreetingAlignment;
}

// The page-level greeting heading. Renders "{phrase}, {firstName}!" when a name
// is available, otherwise just "{phrase}!" (no trailing comma/space before the
// "!"). The single "!" is appended once to the fully assembled greeting, so it
// applies to every time-of-day variant and to the no-name fallback without
// being duplicated per-variant.
export const Greeting: React.FC<IGreetingProps> = (props) => {
  const text = (props.firstName ? `${props.phrase}, ${props.firstName}` : props.phrase) + '!';

  // Absolute positioning (merged with the font styles in styles.greeting):
  // vertical from the slider, horizontal anchored by alignment.
  const style: React.CSSProperties = {
    color: props.color,
    fontSize: `${props.sizePx}px`,
    fontWeight: props.weight,
    position: 'absolute',
    top: `${props.verticalPositionPct}%`
  };
  if (props.alignment === 'Center') {
    style.left = '50%';
    style.transform = 'translateX(-50%)';
  } else if (props.alignment === 'Right') {
    style.right = SIDE_PADDING;
  } else {
    style.left = SIDE_PADDING;
  }

  return (
    <h1 id={props.id} className={styles.greeting} style={style}>
      {text}
    </h1>
  );
};
