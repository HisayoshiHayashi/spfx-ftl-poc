import * as React from 'react';
import { Modal } from '@fluentui/react/lib/Modal';
import { IConsumerScenario, IScenario } from '../data/scenarioModel';
import { scenarioImages } from '../data/images';
import CategoryTags from './CategoryTags';
import ViewerIcon from './ViewerIcon';
import styles from './FTLApp.module.scss';
import type { IScenarioEvaluation, IScenarioEvaluationSnapshot } from '../services/ScenarioEvaluationService';

interface IScenarioDetailProps {
  lightTheme: boolean;
  scenario: IScenario;
  isFavorite: boolean;
  favoritesDisabled: boolean;
  favoritesSaving: boolean;
  favoriteStatus?: string;
  notification?: React.ReactNode;
  loadEvaluation: (id: string) => Promise<IScenarioEvaluationSnapshot>;
  saveEvaluation: (id: string, evaluation: IScenarioEvaluation) => Promise<IScenarioEvaluationSnapshot>;
  onFavorite: (id: string) => void;
  onDismiss: () => void;
}

export default function ScenarioDetail(props: IScenarioDetailProps): React.ReactElement {
  const { scenario, isFavorite, onFavorite, onDismiss } = props;
  const [consumers, setConsumers] = React.useState<IConsumerScenario[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();
  const [isOpen, setIsOpen] = React.useState(true);
  const [evaluation, setEvaluation] = React.useState<IScenarioEvaluation>();
  const [answers, setAnswers] = React.useState<Partial<IScenarioEvaluation>>({});
  const [evaluationLoading, setEvaluationLoading] = React.useState(true);
  const [evaluationSaving, setEvaluationSaving] = React.useState(false);
  const [evaluationError, setEvaluationError] = React.useState<string>();
  const [evaluationStatus, setEvaluationStatus] = React.useState<string>();
  const [evaluationReload, setEvaluationReload] = React.useState(0);
  const [evaluationOpen, setEvaluationOpen] = React.useState(false);
  const evaluationTrigger = React.useRef<HTMLButtonElement>(null);

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

  React.useEffect(() => {
    let active = true;
    setEvaluationLoading(true);
    setEvaluationError(undefined);
    setEvaluationStatus(undefined);
    props.loadEvaluation(scenario.id).then(snapshot => {
      if (!active) return;
      setEvaluation(snapshot.evaluation);
      setAnswers(snapshot.evaluation || {});
      setEvaluationLoading(false);
    }).catch((cause: unknown) => {
      console.error('FTLApp: シナリオ評価の読み込みに失敗しました。', cause);
      if (!active) return;
      setEvaluationError(cause instanceof Error ? cause.message : '評価を読み込めませんでした。');
      setEvaluationLoading(false);
    });
    return () => { active = false; };
  }, [props.loadEvaluation, scenario.id, evaluationReload]);

  const submitEvaluation = (): void => {
    if (evaluationLoading || evaluationSaving ||
      (answers.willHappen !== -1 && answers.willHappen !== 1) ||
      (answers.lifeImpact !== -1 && answers.lifeImpact !== 1) ||
      (answers.wantItCome !== -1 && answers.wantItCome !== 1)) return;
    const submitted: IScenarioEvaluation = {
      willHappen: answers.willHappen,
      lifeImpact: answers.lifeImpact,
      wantItCome: answers.wantItCome
    };
    setEvaluationSaving(true);
    setEvaluationError(undefined);
    setEvaluationStatus(undefined);
    props.saveEvaluation(scenario.id, submitted).then(snapshot => {
      setEvaluation(snapshot.evaluation);
      setEvaluationStatus(evaluation ? '評価を更新しました。ポイントは付与されません。' : '評価を保存しました。1ptを獲得しました。');
      setEvaluationSaving(false);
    }).catch((cause: unknown) => {
      console.error('FTLApp: シナリオ評価の保存に失敗しました。', cause);
      setEvaluationError(cause instanceof Error ? cause.message : '評価を保存できませんでした。');
      setEvaluationSaving(false);
    });
  };

  const closeEvaluation = (): void => {
    setEvaluationOpen(false);
    window.requestAnimationFrame(() => evaluationTrigger.current?.focus());
  };

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={() => setIsOpen(false)}
      onDismissed={onDismiss}
      className={`${styles.detailModal} ${props.lightTheme ? styles.lightTheme : ''} ${isOpen ? styles.detailModalOpening : ''}`}
      titleAriaId={`scenario-title-${scenario.id}`}
      layerProps={{ styles: { root: { zIndex: 100010 } } }}
      styles={{
        main: { background: props.lightTheme ? '#ffffff' : '#11151d', color: props.lightTheme ? '#28313d' : '#e3e8ee', width: 'calc(100% - 40px)', maxWidth: 960, height: '92vh', maxHeight: '92vh', borderRadius: 20, overflow: 'hidden', display: 'flex', flexDirection: 'column' },
        scrollableContent: { flex: '1 1 auto', maxHeight: '92vh', minHeight: 0, overflow: 'hidden', display: 'flex', flexDirection: 'column' }
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
        <div className={styles.detailContent}>
          <img className={styles.detailImage} src={scenarioImages[scenario.id]} alt={scenario.title} />
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
        <div className={styles.evaluationDock}>
          <button ref={evaluationTrigger} type="button" className={styles.evaluationLauncher} aria-label="評価フォームを開く" aria-haspopup="dialog" aria-expanded={evaluationOpen} onClick={() => setEvaluationOpen(true)}>
            <strong>{evaluation ? '評価を編集する' : 'このシナリオを評価'}</strong>
          </button>
        </div>
        {evaluationOpen && <div className={styles.evaluationOverlay} role="presentation" onClick={event => { if (event.target === event.currentTarget) closeEvaluation(); }}>
          <section className={styles.evaluationSheet} role="dialog" aria-modal="true" aria-labelledby={`evaluation-title-${scenario.id}`} onKeyDown={event => { if (event.key === 'Escape') closeEvaluation(); }}>
            <header className={styles.evaluationSheetHeader}>
              <div><span className={styles.sectionLabel}>YOUR EVALUATION</span><h2 id={`evaluation-title-${scenario.id}`}>未来を想像してみよう</h2></div>
              <button type="button" className={styles.iconButton} aria-label="評価を閉じる" onClick={closeEvaluation}><ViewerIcon name="close" /></button>
            </header>
            {evaluationLoading && <p role="status">評価を読み込み中…</p>}
            {evaluationError && <p role="alert" className={styles.error}>{evaluationError} <button type="button" disabled={evaluationSaving} onClick={() => setEvaluationReload(value => value + 1)}>再読み込み</button></p>}
            {!evaluationLoading && !evaluationError && <>
              {([
                ['willHappen', '本当にくる？', 'きそう', 'こない'],
                ['lifeImpact', '生活はかわる？', 'かわる', 'かわらない'],
                ['wantItCome', '来たらうれしい？', '来てほしい', '来てほしくない']
              ] as const).map(([key, question, positive, negative]) => (
                <fieldset key={key} className={styles.evaluationQuestion} disabled={evaluationSaving}>
                  <legend>{question}</legend>
                  <div className={styles.evaluationChoices}>
                    {([
                      [1, positive],
                      [-1, negative]
                    ] as const).map(([score, label]) => {
                      const selectedAnswer = answers[key] === score;
                      return <label key={score} className={`${styles.evaluationChoice} ${selectedAnswer ? styles.evaluationChoiceSelected : ''}`}>
                        <input type="radio" name={`${scenario.id}-${key}`} value={score} checked={selectedAnswer} onChange={() => {
                          setAnswers(value => ({ ...value, [key]: score }));
                        }} />
                        <span>{label}</span>
                      </label>;
                    })}
                  </div>
                </fieldset>
              ))}
              <button type="button" className={styles.evaluationSubmit} disabled={evaluationSaving || answers.willHappen === undefined || answers.lifeImpact === undefined || answers.wantItCome === undefined} onClick={submitEvaluation}>
                {evaluationSaving ? '保存中…' : evaluation ? '評価を更新する' : '評価する'}
              </button>
              {evaluation && <p className={styles.muted}>回答を変更できます。編集ではポイントは増えません。</p>}
              {evaluationStatus && <p role="status" className={styles.evaluationStatus}>{evaluationStatus}</p>}
            </>}
          </section>
        </div>}
        {props.notification}
      </div>
    </Modal>
  );
}
