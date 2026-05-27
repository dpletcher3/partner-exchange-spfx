import * as React from 'react';
import styles from './PhillipsPersonalizedHero.module.scss';

export interface IGreetingProps {
  // Time-of-day phrase, e.g. "Good morning".
  phrase: string;
  // First name, or undefined when none could be derived.
  firstName?: string;
  color: string;
  sizePx: number;
  weight: number;
  id: string;
}

// The page-level greeting heading. Renders "{phrase}, {firstName}" when a name
// is available, otherwise just "{phrase}" (no trailing comma/space).
export const Greeting: React.FC<IGreetingProps> = (props) => {
  const text = props.firstName ? `${props.phrase}, ${props.firstName}` : props.phrase;

  return (
    <h1
      id={props.id}
      className={styles.greeting}
      style={{
        color: props.color,
        fontSize: `${props.sizePx}px`,
        fontWeight: props.weight
      }}
    >
      {text}
    </h1>
  );
};
