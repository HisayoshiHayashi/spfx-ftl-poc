import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IConsumerScenario, IScenario } from '../data/scenarioModel';
import { scenarioImages } from '../data/images';
import CategoryTags from './CategoryTags';
import ViewerIcon from './ViewerIcon';
import styles from './FTLApp.module.scss';

interface IScenarioDetailProps {
  scenario: IScenario;
  isFavorite: boolean;
  favoritesDisabled: boolean;
  favoritesSaving: boolean;
  favoriteStatus?: string;
  notification?: React.ReactNode;
  onFavorite: (id: string) => void;
  onDismiss: () => void;
}

export default function ScenarioDetail(props: IScenarioDetailProps): React.ReactElement {
  const { scenario, isFavorite, onFavorite, onDismiss } = props;
  const [consumers, setConsumers] = React.useState<IConsumerScenario[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();
  const [isOpen, setIsOpen] = React.useState(true);

  React.useEffect(() => {
    let active = true;
    setLoading(true);
    setError(undefined);
    import(/* webpackChunkName: "consumer-scenarios" */ '../data/consumers.json')
      .then(module => {
        if (active) {
          setConsumers(module.default.filter(item => item.scenarioId === scenario.id));
          setLoading(false);
        }
      })
      .catch((cause: unknown) => {
        console.error('FTLApp: 生活者シナリオの読み込みに失敗しました。', cause);
        if (active) {
          setError('生活者シナリオを読み込めませんでした。画面を閉じて再度お試しください。');
          setLoading(false);
        }
      });
    return () => { active = false; };
  }, [scenario.id]);

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={() => setIsOpen(false)}
      onDismissed={onDismiss}
      className={`${styles.detailModal} ${isOpen ? styles.detailModalOpening : ''}`}
      titleAriaId={`scenario-title-${scenario.id}`}
      layerProps={{ styles: { root: { zIndex: 100010 } } }}
      styles={{
        main: { background: '#11151d', color: '#e3e8ee', width: 'calc(100% - 40px)', maxWidth: 960, maxHeight: '92vh', borderRadius: 20, overflow: 'hidden', display: 'flex', flexDirection: 'column' },
        scrollableContent: { maxHeight: '92vh', minHeight: 0, overflowY: 'auto' }
      }}
    >
      <div className={styles.detail}>
        <div className={styles.detailToolbar}>
          <span>{scenario.id} <span className={styles.detailYear}>{scenario.year}年</span></span>
          <div>
            <button type="button" disabled={props.favoritesDisabled} aria-disabled={props.favoritesSaving || undefined} className={`${styles.iconButton} ${isFavorite ? styles.favorited : ''}`} aria-label={isFavorite ? 'お気に入りから削除' : 'お気に入りに追加'} aria-pressed={isFavorite} onClick={() => onFavorite(scenario.id)}><ViewerIcon name="heart" /></button>
            <button type="button" className={styles.iconButton} aria-label="詳細を閉じる" onClick={() => setIsOpen(false)}><ViewerIcon name="close" /></button>
          </div>
          {props.favoriteStatus && <p role="status" className={styles.error}>{props.favoriteStatus}</p>}
        </div>
        <img className={styles.detailImage} src={scenarioImages[scenario.id]} alt={scenario.title} />
        <div className={styles.detailContent}>
          <CategoryTags names={scenario.categories} />
          <h2 id={`scenario-title-${scenario.id}`}>{scenario.title}</h2>
          <p className={styles.summary}>{scenario.summary}</p>
          <h3>変化のポイント</h3>
          <div className={styles.changes}>
            <section><h4>これまで</h4><ul>{scenario.before.filter(Boolean).map((point, index) => <li key={index}>{point}</li>)}</ul></section>
            <span className={styles.changeArrow}><ViewerIcon name="arrow" /></span>
            <section><h4>これから · {scenario.year}</h4><ul>{scenario.after.filter(Boolean).map((point, index) => <li key={index}>{point}</li>)}</ul></section>
          </div>
          {scenario.background && <section className={styles.storySection}><span className={styles.sectionLabel}>BACKGROUND</span><h3>{scenario.backgroundTitle || '背景'}</h3><p>{scenario.background}</p></section>}
          {scenario.detail && <section className={styles.storySection}><span className={styles.sectionLabel}>FUTURE SCENARIO</span><h3>{scenario.detailTitle || '未来シナリオ'}</h3><p>{scenario.detail}</p></section>}
          <section className={styles.storySection}>
            <span className={styles.sectionLabel}>LIFE SCENARIOS</span>
            <h3>生活者の視点 <span className={styles.muted}>{loading ? '' : `· ${consumers.length}件`}</span></h3>
            {loading && <p role="status">生活者シナリオを読み込み中…</p>}
            {error && <p role="alert" className={styles.error}>{error}</p>}
            {!loading && !error && consumers.length === 0 && <p className={styles.muted}>このシナリオに関連する生活者データはありません。</p>}
            {!loading && !error && consumers.map(consumer => (
              <details key={consumer.id} className={styles.consumer}>
                <summary><span className={styles.persona}>生活者 {consumer.persona}</span>{consumer.title}</summary>
                <p>{consumer.story}</p>
              </details>
            ))}
          </section>
          {(scenario.sourceClip || scenario.socialData) && <footer className={styles.references}><p>元兆しクリップ：{scenario.sourceClip || '—'}</p><p>参照社会データ：{scenario.socialData || '—'}</p></footer>}
        </div>
        {props.notification}
      </div>
    </Modal>
  );
}
