import * as React from 'react';

import styles from './PhillipsNews.module.scss';
import { INewsRepositoryService } from '../services/INewsRepositoryService';
import { INewsItem, INewsFilters } from '../services/models';
import { NewsGrid } from './NewsGrid';
import { LoadingState } from './LoadingState';
import { EmptyState } from './EmptyState';
import { ErrorState } from './ErrorState';
import { AddNewsItemButton } from './AddNewsItemButton';
import { DataSource } from '../config/constants';

export interface IPhillipsNewsProps {
  service: INewsRepositoryService;
  // Selects list vs news-pipeline behavior for the header affordances (the data
  // itself arrives via `service`, which the web part injects per dataSource).
  dataSource: DataSource;
  sectionTitle: string;
  categoryFilter: string[];
  itemTypeFilter: string;
  maxItems: number;
  showViewAllLink: boolean;
  sourceSiteUrl: string;
  listTitle: string;
  // True when the page is in edit mode; gates the editor-only +Add affordance.
  isEditMode: boolean;
}

type Status = 'loading' | 'populated' | 'empty' | 'error';

const SKELETON_COUNT = 6;

export const PhillipsNews: React.FC<IPhillipsNewsProps> = (props) => {
  const [status, setStatus] = React.useState<Status>('loading');
  const [items, setItems] = React.useState<INewsItem[]>([]);
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  // Stable per-instance id so the <h2> can be referenced by aria-labelledby.
  const headerId = React.useMemo(
    () => `phil-news-title-${Math.random().toString(36).slice(2)}`,
    []
  );

  // Join the category array so the effect re-runs when its contents change,
  // not just its identity.
  const categoryKey = props.categoryFilter.join('|');

  const load = React.useCallback(() => {
    setStatus('loading');
    const filters: INewsFilters = {
      categories: props.categoryFilter,
      itemType: props.itemTypeFilter
    };
    props.service
      .getNewsItems(props.sourceSiteUrl, props.listTitle, filters, props.maxItems)
      .then((result) => {
        if (!result || result.length === 0) {
          setItems([]);
          setStatus('empty');
        } else {
          setItems(result);
          setStatus('populated');
        }
      })
      .catch((err: unknown) => {
        setErrorMessage(err instanceof Error ? err.message : String(err));
        setStatus('error');
      });
  }, [
    props.service,
    props.sourceSiteUrl,
    props.listTitle,
    categoryKey,
    props.itemTypeFilter,
    props.maxItems
  ]);

  React.useEffect(() => {
    load();
  }, [load]);

  const viewAllUrl = buildViewAllUrl(props.sourceSiteUrl, props.listTitle, props.dataSource);
  const hasHeader = !!props.sectionTitle || props.showViewAllLink || props.isEditMode;

  return (
    <section
      className={styles.phillipsNews}
      aria-labelledby={props.sectionTitle ? headerId : undefined}
    >
      {hasHeader && (
        <div className={styles.header}>
          {props.sectionTitle ? (
            <h2 id={headerId} className={styles.title}>
              {props.sectionTitle}
            </h2>
          ) : (
            <span />
          )}
          <div className={styles.headerActions}>
            {/* Editor-only: conditionally rendered, absent from the DOM in read mode. */}
            {props.isEditMode && (
              <AddNewsItemButton
                sourceSiteUrl={props.sourceSiteUrl}
                listTitle={props.listTitle}
                dataSource={props.dataSource}
              />
            )}
            {props.showViewAllLink && (
              <a className={styles.viewAll} href={viewAllUrl}>
                View all →
              </a>
            )}
          </div>
        </div>
      )}

      {status === 'loading' && <LoadingState count={SKELETON_COUNT} />}
      {status === 'error' && <ErrorState message={errorMessage} onRetry={load} />}
      {status === 'empty' && <EmptyState />}
      {status === 'populated' && <NewsGrid items={items} />}
    </section>
  );
};

function buildViewAllUrl(siteUrl: string, listTitle: string, dataSource: DataSource): string {
  const trimmed = siteUrl.replace(/\/+$/, '');
  if (dataSource === 'pipeline') {
    // Pipeline mode reads the Site Pages library; point "View all" at that
    // library's view rather than the News Repository list view.
    return `${trimmed}/SitePages/Forms/AllItems.aspx`;
  }
  return `${trimmed}/Lists/${encodeURIComponent(listTitle)}/AllItems.aspx`;
}
