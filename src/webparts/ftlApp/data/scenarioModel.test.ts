import scenarios from './scenarios.json';
import consumers from './consumers.json';
import { categories, filterScenarios, IScenarioFilters, parseFavorites } from './scenarioModel';

const all: IScenarioFilters = { year: undefined, category: undefined, query: '', favoritesOnly: false, favorites: [] };

describe('Excel scenario dataset', () => {
  it('preserves all scenarios and the screenshot year counts', () => {
    expect(scenarios).toHaveLength(559);
    expect(new Set(scenarios.map(item => item.id)).size).toBe(559);
    expect([2030, 2035, 2040, 2045, 2050].map(year => scenarios.filter(item => item.year === year).length))
      .toEqual([116, 124, 117, 97, 105]);
    expect(categories.map(category => scenarios.filter(item => item.categories.indexOf(category.name) !== -1).length))
      .toEqual([239, 484, 69, 21, 447, 99, 71]);
  });

  it('links every consumer story to a valid parent scenario', () => {
    expect(consumers).toHaveLength(3304);
    expect(new Set(consumers.map(item => item.id)).size).toBe(3304);
    const ids = new Set(scenarios.map(item => item.id));
    expect(consumers.every(item => ids.has(item.scenarioId))).toBe(true);
    expect(consumers.filter(item => item.scenarioId === 'MS00001')).toHaveLength(7);
    expect(scenarios.every(item => item.categories.every(name => categories.some(category => category.name === name)))).toBe(true);
  });
});

describe('Scenario filtering', () => {
  it('shows everything without filters and handles empty search terms', () => {
    expect(filterScenarios(scenarios, all)).toHaveLength(559);
    expect(filterScenarios(scenarios, { ...all, query: '   ' })).toHaveLength(559);
  });

  it('combines year, category, case-insensitive search and favorites', () => {
    const filters = { ...all, year: 2035, category: 'エコノミー', query: 'ms00001 景観', favoritesOnly: true, favorites: ['MS00001'] };
    expect(filterScenarios(scenarios, filters).map(item => item.id)).toEqual(['MS00001']);
    expect(filterScenarios(scenarios, { ...filters, year: 2050 })).toHaveLength(0);
    expect(filterScenarios(scenarios, { ...filters, category: 'ライフデザイン' })).toHaveLength(0);
  });

  it('handles no matching query and an empty favorites collection', () => {
    expect(filterScenarios(scenarios, { ...all, query: 'not-a-real-scenario' })).toHaveLength(0);
    expect(filterScenarios(scenarios, { ...all, favoritesOnly: true })).toHaveLength(0);
  });
});

describe('Favorite persistence', () => {
  it('reads and deduplicates saved IDs', () => {
    expect(parseFavorites(undefined)).toEqual([]);
    expect(parseFavorites('["MS00001","MS00001","MS00002"]')).toEqual(['MS00001', 'MS00002']);
  });

  it('rejects corrupted data rather than silently accepting it', () => {
    expect(() => parseFavorites('{')).toThrow();
    expect(() => parseFavorites('{}')).toThrow();
    expect(() => parseFavorites('[123]')).toThrow();
  });
});
