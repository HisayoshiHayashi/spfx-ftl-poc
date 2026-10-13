export interface IScenario {
  id: string;
  title: string;
  summary: string;
  year: number;
  categories: string[];
  before: string[];
  after: string[];
  backgroundTitle: string;
  background: string;
  detailTitle: string;
  detail: string;
  sourceClip: string;
  socialData: string;
}

export interface IConsumerScenario {
  id: string;
  scenarioId: string;
  persona: number;
  title: string;
  story: string;
  year: number;
}

export const categories = [
  'エコノミー',
  'エシカル・ガバナンス',
  'エネルギー・サステナビリティ',
  'カルチャー・エンターテインメント',
  'デジタルインテリジェンス',
  'モビリティ・ロボティクス',
  'ライフデザイン'
].map(name => ({ name, color: '#94a3b8' }));

export interface IScenarioFilters {
  year: number | undefined;
  category: string | undefined;
  query: string;
  favoritesOnly: boolean;
  favorites: readonly string[];
}

export function filterScenarios(scenarios: readonly IScenario[], filters: IScenarioFilters): IScenario[] {
  const words = filters.query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
  return scenarios.filter(scenario => {
    if (filters.year !== undefined && scenario.year !== filters.year) return false;
    if (filters.category !== undefined && scenario.categories.indexOf(filters.category) === -1) return false;
    if (filters.favoritesOnly && filters.favorites.indexOf(scenario.id) === -1) return false;
    const text = `${scenario.id} ${scenario.title} ${scenario.summary} ${scenario.categories.join(' ')} ${scenario.background} ${scenario.detail}`.toLocaleLowerCase();
    return words.every(word => text.indexOf(word) !== -1);
  });
}

export function parseFavorites(value: string | undefined): string[] {
  if (value === undefined) return [];
  const parsed: unknown = JSON.parse(value);
  if (!Array.isArray(parsed) || !parsed.every(item => typeof item === 'string')) {
    throw new Error('お気に入りの保存データ形式が正しくありません。');
  }
  return parsed.filter((item, index) => parsed.indexOf(item) === index);
}
