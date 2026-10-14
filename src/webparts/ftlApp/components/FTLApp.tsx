import * as React from 'react';
import { Spinner } from '@fluentui/react/lib/Spinner';
import styles from './FTLApp.module.scss';
import type { IFTLAppProps } from './IFTLAppProps';
import scenarioData from '../data/scenarios.json';
import { categories, filterScenarios, IScenario, parseFavorites } from '../data/scenarioModel';
import { scenarioImages } from '../data/images';
import ViewerIcon from './ViewerIcon';
import CategoryTags from './CategoryTags';
import ScenarioDetail from './ScenarioDetail';
import FavoritesService from '../services/FavoritesService';
import ScenarioEvaluationService, { IScenarioEvaluation } from '../services/ScenarioEvaluationService';
import type { IScenarioEvaluationSnapshot } from '../services/ScenarioEvaluationService';
import FavoriteToast from './FavoriteToast';

const scenarios: IScenario[] = scenarioData;
const years = [2030, 2035, 2040, 2045, 2050];
const favoriteStorageKey = 'ftl-app:scenario-favorites:v1';
const themeStorageKey = 'ftl-app:theme:v1';

export default function FTLApp(props: IFTLAppProps): React.ReactElement<IFTLAppProps> {
  const [year, setYear] = React.useState<number>();
  const [category, setCategory] = React.useState<string>();
  const [query, setQuery] = React.useState('');
  const [favoritesOnly, setFavoritesOnly] = React.useState(false);
  const [favorites, setFavorites] = React.useState<string[]>([]);
  const [points, setPoints] = React.useState(0);
  const [storageError, setStorageError] = React.useState<string>();
  const [migrationWarning, setMigrationWarning] = React.useState<string>();
  const [favoritesReady, setFavoritesReady] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [saveSuccess, setSaveSuccess] = React.useState(false);
  const [pendingSave, setPendingSave] = React.useState(false);
  const [migrationPending, setMigrationPending] = React.useState(false);
  const [reload, setReload] = React.useState(0);
  const savedFavorites = React.useRef<string[]>([]);
  const saveInFlight = React.useRef(false);
  const service = React.useMemo(() => new FavoritesService(props.spHttpClient, props.webUrl), [props.spHttpClient, props.webUrl]);
  const evaluationService = React.useMemo(() => new ScenarioEvaluationService(props.spHttpClient, props.webUrl), [props.spHttpClient, props.webUrl]);
  const [view, setView] = React.useState<'grid' | 'list'>('grid');
  const [selected, setSelected] = React.useState<IScenario>();
  const [limit, setLimit] = React.useState(40);
  const [showFilters, setShowFilters] = React.useState(false);
  const [page, setPage] = React.useState<'home' | 'about'>('home');
  const [menuOpen, setMenuOpen] = React.useState(false);
  const menuButton = React.useRef<HTMLButtonElement>(null);
  const [listReady, setListReady] = React.useState(false);
  const [userPhotoFailed, setUserPhotoFailed] = React.useState(false);
  const [appearance, setAppearance] = React.useState<{ light: boolean; error?: string }>(() => {
    try {
      const stored = window.localStorage.getItem(themeStorageKey);
      if (stored !== null && stored !== 'light' && stored !== 'dark') throw new Error('Invalid saved theme');
      return { light: stored === 'light' };
    } catch (cause) {
      console.error('FTLApp: 表示テーマを読み込めませんでした。', cause);
      return { light: false, error: '表示テーマの設定を読み込めませんでした。ダークモードで表示しています。' };
    }
  });

  const toggleTheme = (): void => {
    const light = !appearance.light;
    try {
      window.localStorage.setItem(themeStorageKey, light ? 'light' : 'dark');
      setAppearance({ light });
    } catch (cause) {
      console.error('FTLApp: 表示テーマを保存できませんでした。', cause);
      setAppearance({ light, error: '表示テーマを保存できませんでした。この画面には反映していますが、再読み込み後は保持されません。' });
    }
  };
  const userPhotoAccount = props.userEmail || props.userLoginName;
  const userPhotoUrl = userPhotoAccount
    ? `${window.location.origin}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(userPhotoAccount)}`
    : undefined;

  React.useEffect(() => {
    let frame = window.requestAnimationFrame(() => {
      frame = window.requestAnimationFrame(() => setListReady(true));
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  React.useEffect(() => {
    let active = true;
    setFavoritesReady(false);
    setStorageError(undefined);
    setMigrationWarning(undefined);
    setMigrationPending(false);
    service.load().then(snapshot => {
      if (!active) return;
      savedFavorites.current = snapshot.ids;
      setPoints(snapshot.points);
      let ids = snapshot.ids;
      if (!snapshot.exists) {
        try {
          if (!window.localStorage.getItem(service.migrationKey)) {
            const legacy = parseFavorites(window.localStorage.getItem(favoriteStorageKey) ?? undefined);
            ids = legacy.filter(id => scenarios.some(item => item.id === id));
            setMigrationPending(ids.length > 0);
          }
        } catch (cause) {
          console.error('FTLApp: 旧お気に入りの移行データを読み込めませんでした。', cause);
          setMigrationWarning('ブラウザーの旧お気に入りを読み込めませんでした。移行を行う場合は保存設定や旧データを確認して再読み込みしてください。');
        }
      }
      setFavorites(ids);
      setPendingSave(false);
      setFavoritesReady(true);
    }).catch((cause: unknown) => {
      console.error('FTLApp: リストのお気に入りを読み込めませんでした。', cause);
      if (active) setStorageError(cause instanceof Error ? cause.message : 'お気に入りを読み込めませんでした。');
    });
    return () => { active = false; };
  }, [service, reload]);

  React.useEffect(() => { setLimit(40); }, [year, category, query, favoritesOnly]);
  React.useEffect(() => {
    if (!saveSuccess) return;
    const timer = window.setTimeout(() => setSaveSuccess(false), 3000);
    return () => window.clearTimeout(timer);
  }, [saveSuccess]);

  const filtered = React.useMemo(() => filterScenarios(scenarios, {
    year, category, query, favoritesOnly, favorites
  }), [year, category, query, favoritesOnly, favorites]);
  const favoriteCount = scenarios.filter(item => favorites.indexOf(item.id) !== -1).length;

  const persistFavorites = (next: string[]): void => {
    if (!favoritesReady || saveInFlight.current) return;
    saveInFlight.current = true;
    setSaving(true);
    setSaveSuccess(false);
    setStorageError(undefined);
    setFavorites(next);
    setPendingSave(true);
    const finishSaving = (): void => {
      saveInFlight.current = false;
      setSaving(false);
    };
    service.save(savedFavorites.current, next).then(stored => {
      savedFavorites.current = stored;
      setFavorites(stored);
      setPendingSave(false);
      setMigrationPending(false);
      setSaveSuccess(true);
      try {
        window.localStorage.setItem(service.migrationKey, 'completed');
      } catch (cause) {
        console.error('FTLApp: 移行済みマーカーを保存できませんでした。', cause);
        setMigrationWarning('リストへの保存は完了しましたが、ブラウザーに移行済み情報を保存できませんでした。');
      }
      finishSaving();
    }, (cause: unknown) => {
      console.error('FTLApp: リストのお気に入りを保存できませんでした。', cause);
      setStorageError(`${cause instanceof Error ? cause.message : 'お気に入りの保存に失敗しました。'} 未保存の変更があります。再試行してください。`);
      finishSaving();
    });
  };

  const toggleFavorite = (id: string): void => {
    if (!favoritesReady || saving || pendingSave) return;
    const next = favorites.indexOf(id) === -1 ? [...favorites, id] : favorites.filter(item => item !== id);
    persistFavorites(next);
  };

  const loadEvaluation = React.useCallback(async (id: string): Promise<IScenarioEvaluationSnapshot> => {
    if (!favoritesReady) throw new Error('ポイント情報を読み込み中です。評価を再読み込みしてください。');
    const snapshot = await evaluationService.load(id);
    const updatedPoints = await service.reconcileEvaluationPoints(snapshot.evaluationCount);
    setPoints(updatedPoints);
    return snapshot;
  }, [evaluationService, favoritesReady, service]);

  const saveEvaluation = React.useCallback(async (id: string, evaluation: IScenarioEvaluation): Promise<IScenarioEvaluationSnapshot> => {
    if (!favoritesReady) throw new Error('ポイント情報が準備できていません。評価を保存できません。');
    const snapshot = await evaluationService.save(id, evaluation);
    const updatedPoints = await service.reconcileEvaluationPoints(snapshot.evaluationCount);
    setPoints(updatedPoints);
    return snapshot;
  }, [evaluationService, favoritesReady, service]);

  const clearFilters = (): void => {
    setYear(undefined);
    setCategory(undefined);
    setQuery('');
    setFavoritesOnly(false);
  };

  const notification = <FavoriteToast lightTheme={appearance.light} saving={saving} error={pendingSave ? storageError : undefined} success={saveSuccess}
    inline={selected !== undefined} onRetry={() => persistFavorites(favorites)} onDismiss={() => setSaveSuccess(false)} />;

  return (
    <section data-theme={appearance.light ? 'light' : 'dark'} className={`${styles.ftlApp} ${appearance.light ? styles.lightTheme : ''} ${props.isFullScreen ? styles.fullScreen : ''}`}>
      <header className={styles.header}>
        <div className={styles.brand}>
          <h1><span>未来年表</span><span className={styles.brandEnglish} lang="en">Future Time Line</span></h1>
        </div>
        <div className={styles.headerActions}
          onBlur={event => {
            if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) setMenuOpen(false);
          }}
          onKeyDown={event => {
            if (event.key === 'Escape' && menuOpen) {
              setMenuOpen(false);
              menuButton.current?.focus();
            }
          }}>
          <button ref={menuButton} type="button" className={styles.menuToggle} aria-label={menuOpen ? 'ページメニューを閉じる' : 'ページメニューを開く'} aria-expanded={menuOpen} onClick={() => setMenuOpen(!menuOpen)}>
            <ViewerIcon name={menuOpen ? 'close' : 'menu'} />
          </button>
          <nav className={`${styles.pageNavigation} ${menuOpen ? styles.pageNavigationOpen : ''}`} aria-label="ページ">
            <button type="button" aria-current={page === 'home' ? 'page' : undefined} onClick={() => { setPage('home'); setMenuOpen(false); if (menuOpen) menuButton.current?.focus(); }}>HOME</button>
            <button type="button" aria-current={page === 'about' ? 'page' : undefined} onClick={() => { setPage('about'); setMenuOpen(false); if (menuOpen) menuButton.current?.focus(); }}>ABOUT</button>
          </nav>
          <button type="button" className={styles.themeToggle} aria-label={appearance.light ? 'ダークモードに切り替える' : 'ライトモードに切り替える'} onClick={toggleTheme} title={appearance.light ? 'ダークモードに切り替える' : 'ライトモードに切り替える'}>
            <ViewerIcon name={appearance.light ? 'moon' : 'sun'} />
          </button>
          {props.userDisplayName && <div className={styles.headerUser} title={`${points.toLocaleString()} pt`} aria-label={`ログイン中のポイント: ${points.toLocaleString()}ポイント`}>
            <span className={styles.userAvatar} aria-label="プロフィール写真">
              {userPhotoUrl && !userPhotoFailed
                ? <img src={userPhotoUrl} alt="" onError={() => setUserPhotoFailed(true)} />
                : <span aria-hidden="true">{Array.from(props.userDisplayName.trim())[0]}</span>}
            </span>
            <span className={styles.userPoints}><strong>{points.toLocaleString()}</strong><span>pt</span></span>
          </div>}
        </div>
      </header>
      {appearance.error && <div role="alert" className={`${styles.error} ${styles.themeError}`}>{appearance.error}<button type="button" onClick={() => setAppearance({ light: appearance.light })}>閉じる</button></div>}
      <div className={styles.layout} hidden={page !== 'home'}>
        <div className={styles.layoutContent}>
        <aside className={`${styles.sidebar} ${showFilters ? styles.sidebarOpen : ''}`} aria-label="シナリオ絞り込み">
          <div className={styles.filterGroup}>
            <h2>お気に入り</h2>
            <button type="button" className={`${styles.filterButton} ${favoritesOnly ? styles.active : ''}`} aria-pressed={favoritesOnly} onClick={() => setFavoritesOnly(!favoritesOnly)}>
              <ViewerIcon name="heart" /><span>お気に入りのみ</span><span className={styles.count}>{favoriteCount}</span>
            </button>
          </div>
          <div className={styles.filterGroup}>
            <h2>年度</h2>
            <button type="button" className={`${styles.filterButton} ${year === undefined ? styles.active : ''}`} aria-pressed={year === undefined} onClick={() => setYear(undefined)}><span>すべて</span><span className={styles.count}>{scenarios.length}</span></button>
            {years.map(value => <button type="button" key={value} className={`${styles.filterButton} ${year === value ? styles.active : ''}`} aria-pressed={year === value} onClick={() => setYear(value)}><span>{value}</span><span className={styles.count}>{scenarios.filter(item => item.year === value).length}</span></button>)}
          </div>
          <div className={styles.filterGroup}>
            <h2>カテゴリ</h2>
            <button type="button" className={`${styles.filterButton} ${category === undefined ? styles.active : ''}`} aria-pressed={category === undefined} onClick={() => setCategory(undefined)}><span className={styles.dot} style={{ background: '#94a3b8' }} /><span>すべて</span><span className={styles.count}>{scenarios.length}</span></button>
            {categories.map(item => <button type="button" key={item.name} className={`${styles.filterButton} ${category === item.name ? styles.active : ''}`} aria-pressed={category === item.name} onClick={() => setCategory(item.name)}><span className={styles.dot} style={{ background: item.color }} /><span>{item.name}</span><span className={styles.count}>{scenarios.filter(scenario => scenario.categories.indexOf(item.name) !== -1).length}</span></button>)}
          </div>
          <div className={styles.sidebarFooter}><span>FTLApp</span><small>{scenarios.length} SCENARIOS / {years.length} FUTURES</small></div>
        </aside>
        <main className={styles.main}>
          <div className={styles.toolbar}>
            <div role="status" aria-live="polite"><strong>{filtered.length}</strong> 件{(year !== undefined || category || query || favoritesOnly) && <button type="button" className={styles.reset} onClick={clearFilters}>絞り込みを解除</button>}</div>
            <label className={styles.search}>
              <ViewerIcon name="search" />
              <input type="search" aria-label="シナリオをキーワード検索" placeholder="キーワード検索…" value={query} onChange={event => setQuery(event.target.value)} />
            </label>
            <button type="button" className={styles.mobileFilter} aria-expanded={showFilters} onClick={() => setShowFilters(!showFilters)}><ViewerIcon name="filter" />絞り込み</button>
            <div className={styles.viewSwitch} aria-label="表示形式">
              <button type="button" className={view === 'grid' ? styles.viewActive : ''} aria-label="カード表示" aria-pressed={view === 'grid'} onClick={() => setView('grid')}><ViewerIcon name="grid" /></button>
              <button type="button" className={view === 'list' ? styles.viewActive : ''} aria-label="リスト表示" aria-pressed={view === 'list'} onClick={() => setView('list')}><ViewerIcon name="list" /></button>
            </div>
          </div>
          {(!listReady || (!favoritesReady && !storageError)) && <div className={styles.listLoading} role="status"><Spinner label={!listReady ? 'シナリオを読み込み中…' : 'お気に入りをリストから読み込み中…'} /></div>}
          {migrationPending && !pendingSave && <div className={styles.persistenceNotice}>このブラウザーの旧お気に入り{favorites.length}件を移行できます。旧データにはユーザー情報がありません。ご自身のデータであることを確認して保存してください。<button type="button" disabled={saving} onClick={() => persistFavorites(favorites)}>リストに移行して保存</button></div>}
          {migrationWarning && <p role="alert" className={styles.error}>{migrationWarning}</p>}
          {storageError && !pendingSave && <div role="alert" className={`${styles.error} ${styles.storageError}`}>{storageError}<button type="button" disabled={saving} onClick={() => setReload(reload + 1)}>再試行</button></div>}
          {listReady && (filtered.length === 0 ? <div className={styles.empty}><ViewerIcon name="search" /><h2>シナリオが見つかりませんでした</h2><p>キーワードや絞り込み条件を変更してください。</p><button type="button" onClick={clearFilters}>すべてのシナリオを見る</button></div> : (
            <div key={`${view}:${year}:${category}:${favoritesOnly}`} className={`${styles.grid} ${view === 'list' ? styles.listView : ''}`}>
              {filtered.slice(0, limit).map((scenario, index) => {
                const favorite = favorites.indexOf(scenario.id) !== -1;
                return (
                  <article key={scenario.id} className={styles.card} style={{ animationDelay: `${(index % 40) * 35}ms` }}>
                    <button type="button" className={styles.cardOpen} onClick={() => setSelected(scenario)} aria-label={`${scenario.year}年 ${scenario.title}の詳細`}>
                      <span className={styles.cardImage}>
                        <img src={scenarioImages[scenario.id]} alt="" loading="lazy" />
                        <span className={styles.yearBadge}>{scenario.year}</span>
                        <CategoryTags names={scenario.categories} />
                      </span>
                      <span className={styles.cardBody}>
                        <span className={styles.cardTitle}>{scenario.title}</span>
                        <span className={styles.cardSummary}>{scenario.summary}</span>
                      </span>
                    </button>
                    <button type="button" disabled={!favoritesReady || (pendingSave && !saving)} aria-disabled={saving || undefined} className={`${styles.favoriteButton} ${favorite ? styles.favorited : ''}`} aria-label={`${scenario.id}をお気に入り${favorite ? 'から削除' : 'に追加'}`} aria-pressed={favorite} onClick={() => toggleFavorite(scenario.id)}><ViewerIcon name="heart" /></button>
                  </article>
                );
              })}
            </div>
          ))}
          {listReady && limit < filtered.length && <div className={styles.loadMore}><button type="button" onClick={() => setLimit(limit + 40)}>さらに表示</button><span>{Math.min(limit, filtered.length)} / {filtered.length} 件</span></div>}
        </main>
        </div>
      </div>
      {page === 'about' && <div className={styles.layout}>
        <div className={styles.layoutContent}>
        <main className={styles.main} aria-labelledby="about-title">
          <section className={styles.about}>
            <span className={styles.sectionLabel}>ABOUT</span>
            <h2 id="about-title">未来年表について</h2>
            <p>未来年表は、未来の社会や暮らしのシナリオを探索するためのアプリです。</p>
            <p className={styles.muted}>このページの紹介文は仮の内容です。詳しい説明やコンテンツは、今後追加します。</p>
            <button type="button" className={styles.filterButton} onClick={() => setPage('home')}>HOMEに戻る</button>
          </section>
        </main>
        </div>
      </div>}
      {selected && <ScenarioDetail lightTheme={appearance.light} scenario={selected} isFavorite={favorites.indexOf(selected.id) !== -1} favoritesDisabled={!favoritesReady || (pendingSave && !saving)} favoritesSaving={saving} favoriteStatus={!favoritesReady ? storageError || 'お気に入りを読み込み中…' : undefined} notification={notification} onFavorite={toggleFavorite} loadEvaluation={loadEvaluation} saveEvaluation={saveEvaluation} onDismiss={() => setSelected(undefined)} />}
      {!selected && notification}
    </section>
  );
}
