import * as React from 'react';
import { categories } from '../data/scenarioModel';
import styles from './FTLApp.module.scss';

export default function CategoryTags({ names }: { names: string[] }): React.ReactElement {
  return (
    <span className={styles.tags}>
      {names.map(name => {
        const category = categories.find(item => item.name === name);
        const color = category ? category.color : '#94a3b8';
        return <span key={name} className={styles.tag} style={{ color: `var(--category-text, ${color})`, borderColor: `var(--category-border, ${color}99)`, backgroundColor: `var(--category-background, ${color}26)` }}>{name}</span>;
      })}
    </span>
  );
}
