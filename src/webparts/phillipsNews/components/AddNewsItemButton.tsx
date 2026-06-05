import * as React from 'react';
import styles from './AddNewsItemButton.module.scss';
import { DataSource } from '../config/constants';

export interface IAddNewsItemButtonProps {
  sourceSiteUrl: string;
  listTitle: string;
  dataSource: DataSource;
}

// Editor-only affordance: opens a news-authoring entry point in a new tab.
// In list mode that's the News Repository list's NewForm; in pipeline mode it's
// the site's native "New news" page authoring. URLs are derived from existing
// property values — no new property-pane fields.
export const AddNewsItemButton: React.FC<IAddNewsItemButtonProps> = ({
  sourceSiteUrl,
  listTitle,
  dataSource
}) => {
  const newItemUrl = buildNewItemUrl(sourceSiteUrl, listTitle, dataSource);

  return (
    <a
      className={styles.addButton}
      href={newItemUrl}
      target="_blank"
      rel="noopener noreferrer"
    >
      + Add news item
    </a>
  );
};

function buildNewItemUrl(siteUrl: string, listTitle: string, dataSource: DataSource): string {
  const trimmed = siteUrl.replace(/\/+$/, '');
  if (dataSource === 'pipeline') {
    // Native news authoring (New → News post). CreateSitePage with the News
    // page type opens the same flow the OOTB "+ New" menu uses.
    return `${trimmed}/_layouts/15/CreateSitePage.aspx?pageType=News`;
  }
  return `${trimmed}/Lists/${encodeURIComponent(listTitle)}/NewForm.aspx`;
}
