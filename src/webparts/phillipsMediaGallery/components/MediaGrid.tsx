import * as React from 'react';
import styles from './PhillipsMediaGallery.module.scss';
import { MediaCard } from './MediaCard';

export interface IMediaGridProps {
  // Desktop column count (3–5). Collapses responsively on narrow viewports.
  columns: number;
  // Turn 1: number of placeholder cards to render. Turn 2 swaps this for a real
  // items array.
  placeholderCount: number;
}

export const MediaGrid: React.FC<IMediaGridProps> = ({ columns, placeholderCount }) => {
  // The chosen column count drives the grid via a CSS custom property so the
  // SCSS owns the responsive breakpoints (see .grid).
  const gridStyle = { ['--phil-mg-cols']: String(columns) } as React.CSSProperties;
  return (
    <div className={styles.grid} style={gridStyle}>
      {Array.from({ length: placeholderCount }).map((_, i) => (
        <MediaCard key={`placeholder-${i}`} />
      ))}
    </div>
  );
};
