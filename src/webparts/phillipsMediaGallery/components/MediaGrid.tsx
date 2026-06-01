import * as React from 'react';
import { HttpClient } from '@microsoft/sp-http';
import styles from './PhillipsMediaGallery.module.scss';
import { IMediaCardItem } from '../services/models';
import { MediaCard } from './MediaCard';

export interface IMediaGridProps {
  items: IMediaCardItem[];
  // Desktop column count (3–5); collapses responsively on narrow viewports.
  columns: number;
  openInNewTab: boolean;
  httpClient: HttpClient;
}

export const MediaGrid: React.FC<IMediaGridProps> = ({ items, columns, openInNewTab, httpClient }) => {
  // The chosen column count drives the grid via a CSS custom property; the SCSS
  // owns the responsive breakpoints (see .grid).
  const gridStyle = { ['--phil-mg-cols']: String(columns) } as React.CSSProperties;
  return (
    <div className={styles.grid} style={gridStyle}>
      {items.map((item) => (
        <MediaCard key={item.id} item={item} openInNewTab={openInNewTab} httpClient={httpClient} />
      ))}
    </div>
  );
};
