const mockLoad = jest.fn();
const mockSave = jest.fn();

jest.mock('../data/images', () => ({ scenarioImages: {} }));
jest.mock('./FTLApp.module.scss', () => ({ listLoading: 'listLoading' }), { virtual: true });
jest.mock('./ScenarioDetail', () => ({ __esModule: true, default: () => undefined }));
jest.mock('../services/FavoritesService', () => ({
  __esModule: true,
  default: jest.fn().mockImplementation(() => ({
    load: mockLoad,
    save: mockSave,
    migrationKey: 'test-migration-current-user'
  }))
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import type { SPHttpClient } from '@microsoft/sp-http';
import FTLApp from './FTLApp';

const client: Pick<SPHttpClient, 'get' | 'post'> = { get: jest.fn(), post: jest.fn() };

describe('Viewer SharePoint favorites UI', () => {
  let container: HTMLDivElement;
  let consoleError: jest.SpyInstance;
  let requestFrame: jest.SpyInstance;
  let cancelFrame: jest.SpyInstance;
  let frames: FrameRequestCallback[];

  beforeEach(() => {
    mockLoad.mockReset();
    mockSave.mockReset();
    window.localStorage.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
    consoleError = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    frames = [];
    requestFrame = jest.spyOn(window, 'requestAnimationFrame').mockImplementation(callback => frames.push(callback));
    cancelFrame = jest.spyOn(window, 'cancelAnimationFrame').mockImplementation(() => undefined);
  });

  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    consoleError.mockRestore();
    requestFrame.mockRestore();
    cancelFrame.mockRestore();
  });

  function paintFrame(): void {
    const callback = frames.shift();
    if (!callback) throw new Error('No animation frame scheduled');
    act(() => { callback(0); });
  }

  async function render(ready = true): Promise<void> {
    await act(async () => {
      ReactDom.render(
        <FTLApp description="test" isFullScreen={false} isServedFromLocalhost
          spHttpClient={client} webUrl="https://example.sharepoint.com/sites/-TMC-" />,
        container
      );
    });
    if (ready) {
      paintFrame();
      paintFrame();
    }
  }

  function favoriteButton(): HTMLButtonElement {
    const button = container.querySelector<HTMLButtonElement>('button[aria-label^="MS00001をお気に入り"]');
    if (!button) throw new Error('Favorite button not rendered');
    return button;
  }

  async function clickButton(text: string): Promise<void> {
    const button = Array.from(document.querySelectorAll('button')).find(item => item.textContent === text);
    if (!button) throw new Error(`Button not rendered: ${text}`);
    await act(async () => { button.click(); });
  }

  it('starts the initial stagger after a loading paint without waiting for favorites', async () => {
    mockLoad.mockReturnValue(new Promise(() => undefined));
    await render(false);
    expect(container.textContent).toContain('シナリオを読み込み中');
    expect(container.querySelectorAll('.listLoading')).toHaveLength(1);
    expect(container.querySelector('.listLoading')?.parentElement?.tagName).toBe('MAIN');
    expect(container.querySelectorAll('article')).toHaveLength(0);
    paintFrame();
    expect(container.querySelectorAll('article')).toHaveLength(0);
    paintFrame();
    const cards = container.querySelectorAll('article');
    expect(cards).toHaveLength(40);
    expect(cards[0].style.animationDelay).toBe('0ms');
    expect(cards[1].style.animationDelay).toBe('35ms');
    expect(cards[39].style.animationDelay).toBe('1365ms');
    expect(container.textContent).not.toContain('シナリオを読み込み中');
    expect(container.querySelectorAll('.listLoading')).toHaveLength(1);
    expect(container.querySelector('.listLoading')?.textContent).toContain('お気に入りをリストから読み込み中');
    await clickButton('さらに表示');
    const expanded = container.querySelectorAll('article');
    expect(expanded).toHaveLength(80);
    expect(expanded[0]).toBe(cards[0]);
    expect(expanded[40].style.animationDelay).toBe('0ms');
    expect(expanded[41].style.animationDelay).toBe('35ms');
  });

  it('cancels the pending initial frame when unmounted', async () => {
    mockLoad.mockResolvedValue({ ids: [], exists: true });
    await render(false);
    paintFrame();
    const pendingFrame = requestFrame.mock.results[requestFrame.mock.results.length - 1].value;
    act(() => { ReactDom.unmountComponentAtNode(container); });
    expect(cancelFrame).toHaveBeenCalledWith(pendingFrame);
  });

  it('switches between HOME and ABOUT while preserving the list and keeping search outside the header', async () => {
    mockLoad.mockResolvedValue({ ids: [], exists: true });
    await render();
    expect(container.querySelector('header input[type="search"]')).toBeNull();
    expect(container.querySelector('main input[type="search"]')).not.toBeNull();
    await clickButton('さらに表示');
    const home = container.querySelector('article')?.closest('main')?.parentElement;
    const firstCard = container.querySelector('article');
    expect(home).not.toBeNull();
    await clickButton('ABOUT');
    expect(home?.hidden).toBe(true);
    expect(container.querySelector('#about-title')?.textContent).toBe('未来年表について');
    expect(container.querySelector('#about-title')?.closest('[hidden]')).toBeNull();
    expect(container.querySelector('nav button[aria-current="page"]')?.textContent).toBe('ABOUT');
    await clickButton('HOME');
    expect(home?.hidden).toBe(false);
    expect(container.querySelector('#about-title')).toBeNull();
    expect(container.querySelectorAll('article')).toHaveLength(80);
    expect(container.querySelector('article')).toBe(firstCard);
  });

  it('closes mobile navigation after selection, Escape and focus leaving the menu', async () => {
    mockLoad.mockResolvedValue({ ids: [], exists: true });
    await render();
    const toggle = container.querySelector<HTMLButtonElement>('button[aria-label="ページメニューを開く"]');
    if (!toggle) throw new Error('Menu toggle not rendered');
    act(() => { toggle.focus(); toggle.click(); });
    expect(toggle.getAttribute('aria-expanded')).toBe('true');
    await clickButton('ABOUT');
    expect(toggle.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(toggle);
    act(() => { toggle.click(); });
    act(() => { toggle.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); });
    expect(toggle.getAttribute('aria-expanded')).toBe('false');
    act(() => { toggle.click(); });
    const back = Array.from(container.querySelectorAll('button')).find(button => button.textContent === 'HOMEに戻る');
    if (!back) throw new Error('About back button not rendered');
    act(() => { back.focus(); });
    expect(toggle.getAttribute('aria-expanded')).toBe('false');
  });

  it('prefers an existing server row and never imports the old browser data again', async () => {
    window.localStorage.setItem('ftl-app:scenario-favorites:v1', '["MS00001"]');
    mockLoad.mockResolvedValue({ ids: ['MS00002'], exists: true });
    await render();
    expect(favoriteButton().getAttribute('aria-pressed')).toBe('false');
    expect(container.textContent).not.toContain('リストに移行して保存');
    expect(mockSave).not.toHaveBeenCalled();
  });

  it('moves legacy IDs on first save and marks migration complete only after success', async () => {
    window.localStorage.setItem('ftl-app:scenario-favorites:v1', '["MS00001"]');
    mockLoad.mockResolvedValue({ ids: [], exists: false });
    mockSave.mockResolvedValue(['MS00001']);
    await render();
    expect(favoriteButton().getAttribute('aria-pressed')).toBe('true');
    expect(window.localStorage.getItem('test-migration-current-user')).toBeNull();
    await clickButton('リストに移行して保存');
    expect(mockSave).toHaveBeenCalledWith([], ['MS00001']);
    expect(window.localStorage.getItem('test-migration-current-user')).toBe('completed');
    expect(window.localStorage.getItem('ftl-app:scenario-favorites:v1')).toBe('["MS00001"]');
  });

  it('keeps failed changes for retry without pretending they were persisted', async () => {
    mockLoad.mockResolvedValue({ ids: [], exists: false });
    mockSave.mockRejectedValueOnce(new Error('HTTP 403')).mockResolvedValueOnce(['MS00001']);
    await render();
    await act(async () => { favoriteButton().click(); });
    expect(document.body.textContent).toContain('未保存の変更');
    expect(favoriteButton().disabled).toBe(true);
    expect(window.localStorage.getItem('test-migration-current-user')).toBeNull();
    await clickButton('再試行');
    expect(mockSave).toHaveBeenNthCalledWith(2, [], ['MS00001']);
    expect(document.body.textContent).not.toContain('未保存の変更');
    expect(favoriteButton().disabled).toBe(false);
  });

  it('disables favorite edits on load failure and can retry loading', async () => {
    mockLoad.mockRejectedValueOnce(new Error('List missing')).mockResolvedValueOnce({ ids: [], exists: false });
    await render();
    expect(container.textContent).toContain('List missing');
    expect(favoriteButton().disabled).toBe(true);
    await clickButton('再試行');
    expect(favoriteButton().disabled).toBe(false);
    expect(mockSave).not.toHaveBeenCalled();
  });

  it('does not resurrect old favorites after this user has already migrated', async () => {
    window.localStorage.setItem('ftl-app:scenario-favorites:v1', '["MS00001"]');
    window.localStorage.setItem('test-migration-current-user', 'completed');
    mockLoad.mockResolvedValue({ ids: [], exists: false });
    await render();
    expect(favoriteButton().getAttribute('aria-pressed')).toBe('false');
    expect(container.textContent).not.toContain('リストに移行して保存');
  });

  it('shows a saving toast without moving the list, then auto-dismisses success after three seconds', async () => {
    mockLoad.mockResolvedValue({ ids: [], exists: true });
    let complete: ((ids: string[]) => void) | undefined;
    mockSave.mockImplementation(() => new Promise<string[]>(resolve => { complete = resolve; }));
    await render();
    await act(async () => { favoriteButton().click(); });
    expect(document.body.textContent).toContain('お気に入りを保存中…');
    expect(container.textContent).not.toContain('お気に入りを保存中…');
    expect(favoriteButton().disabled).toBe(false);
    expect(favoriteButton().getAttribute('aria-disabled')).toBe('true');
    await act(async () => { favoriteButton().click(); });
    expect(mockSave).toHaveBeenCalledTimes(1);
    jest.useFakeTimers();
    try {
      await act(async () => {
        if (!complete) throw new Error('Save was not started');
        complete(['MS00001']);
      });
      expect(document.body.textContent).toContain('お気に入りを保存しました');
      act(() => { jest.advanceTimersByTime(2999); });
      expect(document.body.textContent).toContain('お気に入りを保存しました');
      act(() => { jest.advanceTimersByTime(1); });
      expect(document.body.textContent).not.toContain('お気に入りを保存しました');
    } finally {
      jest.useRealTimers();
    }
  });
});
