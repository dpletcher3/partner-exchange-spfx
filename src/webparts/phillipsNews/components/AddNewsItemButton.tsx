import * as React from 'react';
import styles from './AddNewsItemButton.module.scss';

export interface IAddNewsItemButtonProps {
  sourceSiteUrl: string;
  listTitle: string;
}

// Editor-only affordance: opens the SharePoint new-item form for the configured
// News Repository list in a new tab. URL is derived from the existing
// sourceSiteUrl + listTitle property values — no new property-pane fields.
export const AddNewsItemButton: React.FC<IAddNewsItemButtonProps> = ({
  sourceSiteUrl,
  listTitle
}) => {
  const newItemUrl = buildNewFormUrl(sourceSiteUrl, listTitle);

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

function buildNewFormUrl(siteUrl: string, listTitle: string): string {
  const trimmed = siteUrl.replace(/\/+$/, '');
  return `${trimmed}/Lists/${encodeURIComponent(listTitle)}/NewForm.aspx`;
}
