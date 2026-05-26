import * as React from 'react';
import { INewsItem } from '../services/models';
import styles from './NewsCard.module.scss';

export interface INewsCardProps {
  item: INewsItem;
}

const DESCRIPTION_MAX_CHARS = 120;

export const NewsCard: React.FC<INewsCardProps> = ({ item }) => {
  const description = truncate(item.shortDescription, DESCRIPTION_MAX_CHARS);
  const hasThumb = !!(item.thumbnail && item.thumbnail.serverRelativeUrl);
  // Card shows a single category label; an item may carry several (MultiChoice).
  const primaryCategory = item.categories.length > 0 ? item.categories[0] : '';

  return (
    <a
      className={styles.card}
      href={item.linkUrl || undefined}
      target="_blank"
      rel="noopener noreferrer"
      aria-label={item.title}
    >
      <div className={styles.thumb}>
        {hasThumb ? (
          <img
            className={styles.thumbImg}
            src={item.thumbnail!.serverRelativeUrl}
            alt={item.thumbnail!.alt}
          />
        ) : (
          <div className={styles.fallback} aria-hidden="true" />
        )}
      </div>
      <div className={styles.content}>
        {primaryCategory && <div className={styles.category}>{primaryCategory}</div>}
        <div className={styles.cardTitle}>{item.title}</div>
        {item.publishedDate && (
          <div className={styles.date}>{formatDate(item.publishedDate)}</div>
        )}
        {description && <div className={styles.description}>{description}</div>}
      </div>
    </a>
  );
};

function truncate(text: string, max: number): string {
  if (!text || text.length <= max) {
    return text || '';
  }
  return `${text.substring(0, max).replace(/\s+$/, '')}…`;
}

// "May 23, 2026" — long month, numeric day, numeric year.
function formatDate(iso: string): string {
  const parsed = new Date(iso);
  if (isNaN(parsed.getTime())) {
    return '';
  }
  return parsed.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}
