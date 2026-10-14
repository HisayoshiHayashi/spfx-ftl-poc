jest.mock('@fluentui/react/lib/Modal', () => ({
  Modal: (props: { children: React.ReactNode }): React.ReactElement => <div>{props.children}</div>
}));
jest.mock('../data/images', () => ({ scenarioImages: {} }));
jest.mock('./FTLApp.module.scss', () => ({}), { virtual: true });
jest.mock('./ViewerIcon', () => ({ __esModule: true, default: () => null }));
jest.mock('./CategoryTags', () => ({ __esModule: true, default: () => null }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import scenarioData from '../data/scenarios.json';
import type { IScenarioEvaluationSnapshot } from '../services/ScenarioEvaluationService';
import ScenarioDetail from './ScenarioDetail';

describe('Scenario evaluation UI', () => {
  let container: HTMLDivElement;
  let consoleError: jest.SpyInstance;
  const loadEvaluation = jest.fn<Promise<IScenarioEvaluationSnapshot>, [string]>();
  const saveEvaluation = jest.fn<Promise<IScenarioEvaluationSnapshot>, [string, { willHappen: -1 | 1; lifeImpact: -1 | 1; wantItCome: -1 | 1 }]>();

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    consoleError = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    loadEvaluation.mockReset();
    saveEvaluation.mockReset();
  });

  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    consoleError.mockRestore();
  });

  async function render(snapshot: IScenarioEvaluationSnapshot): Promise<void> {
    loadEvaluation.mockResolvedValue(snapshot);
    await act(async () => {
      ReactDom.render(
        <ScenarioDetail lightTheme={false} scenario={scenarioData[0]} isFavorite={false}
          favoritesDisabled={false} favoritesSaving={false} loadEvaluation={loadEvaluation}
          saveEvaluation={saveEvaluation} onFavorite={jest.fn()} onDismiss={jest.fn()} />,
        container
      );
    });
  }

  function evaluationButton(): HTMLButtonElement | undefined {
    return Array.from(container.querySelectorAll<HTMLButtonElement>('button'))
      .find(button => button.textContent?.includes('評価して1pt') || button.textContent?.includes('評価を更新する'));
  }

  async function openEvaluation(): Promise<void> {
    const button = container.querySelector<HTMLButtonElement>('button[aria-label="評価フォームを開く"]');
    if (!button) throw new Error('Evaluation launcher not rendered');
    await act(async () => { button.click(); });
    expect(container.querySelector('[role="dialog"][aria-modal="true"]')).not.toBeNull();
  }

  it('requires all three answers before submitting and awards one point on first save', async () => {
    await render({ evaluationCount: 0 });
    expect(container.querySelector('[role="dialog"]')).toBeNull();
    await openEvaluation();
    expect(Array.from(container.querySelectorAll<HTMLInputElement>('input[name="MS00001-willHappen"]')).map(input => input.value)).toEqual(['1', '-1']);
    const submit = evaluationButton();
    expect(submit?.disabled).toBe(true);
    const choices = [
      ['MS00001-willHappen', '1'],
      ['MS00001-lifeImpact', '1'],
      ['MS00001-wantItCome', '1']
    ];
    for (const [name, value] of choices) {
      const radio = container.querySelector<HTMLInputElement>(`input[name="${name}"][value="${value}"]`);
      if (!radio) throw new Error(`Missing rating choice ${name}=${value}`);
      await act(async () => { radio.click(); });
    }
    expect(submit?.disabled).toBe(false);
    saveEvaluation.mockResolvedValue({ evaluation: { willHappen: 1, lifeImpact: 1, wantItCome: 1 }, evaluationCount: 1 });
    await act(async () => { submit?.click(); });
    expect(saveEvaluation).toHaveBeenCalledWith('MS00001', { willHappen: 1, lifeImpact: 1, wantItCome: 1 });
    expect(container.textContent).toContain('1ptを獲得しました');
    expect(evaluationButton()?.textContent).toContain('評価を更新する');
  });

  it('shows an existing evaluation for editing without promising additional points', async () => {
    await render({ evaluation: { willHappen: 1, lifeImpact: -1, wantItCome: 1 }, evaluationCount: 1 });
    await openEvaluation();
    expect(container.querySelector<HTMLInputElement>('input[name="MS00001-willHappen"][value="1"]')?.checked).toBe(true);
    expect(evaluationButton()?.textContent).toContain('評価を更新する');
    expect(container.textContent).toContain('編集ではポイントは増えません');
  });

  it('closes the evaluation overlay with Escape while keeping the detail modal open', async () => {
    await render({ evaluationCount: 0 });
    await openEvaluation();
    const dialog = container.querySelector<HTMLElement>('[role="dialog"]');
    if (!dialog) throw new Error('Evaluation dialog not rendered');
    await act(async () => {
      dialog.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    });
    expect(container.querySelector('[role="dialog"]')).toBeNull();
    expect(container.querySelector('button[aria-label="評価フォームを開く"]')).not.toBeNull();
  });
});
