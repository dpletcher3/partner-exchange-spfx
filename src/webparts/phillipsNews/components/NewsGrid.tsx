import * as React from 'react';
import { INewsItem } from '../services/models';
import { NewsCard } from './NewsCard';
import styles from './PhillipsNews.module.scss';

export interface INewsGridProps {
  items: INewsItem[];
}

export const NewsGrid: React.FC<INewsGridProps> = ({ items }) => {
  return (
    <div className={styles.grid}>
      {items.map((item) => (
        <NewsCard key={item.id} item={item} />
      ))}
    </div>
  );
};
