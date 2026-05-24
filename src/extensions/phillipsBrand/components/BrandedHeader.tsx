import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import styles from './BrandedHeader.module.scss';

export const BrandedHeader: React.FC = () => {
  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.brandedHeader}>
        <span className={styles.placeholder}>BRANDED HEADER PLACEHOLDER</span>
      </div>
    </FluentProvider>
  );
};
