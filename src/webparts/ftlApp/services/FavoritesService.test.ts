jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));

import FavoritesService, { mergeFavoriteChanges, parseFavoritesJson } from './FavoritesService';

const webUrl = 'https://example.sharepoint.com/sites/-TMC-';
const user = { Id: 42, LoginName: 'i:0#.f|membership|user@example.com', Title: 'Test User' };
const fields = { value: [
  { Title: 'UserKey', InternalName: 'UserKey', EntityPropertyName: 'UserKey', TypeAsString: 'Text' },
  { Title: 'FavoritesJson', InternalName: 'FavoritesJson', EntityPropertyName: 'FavoritesJson', TypeAsString: 'Note' },
  { Title: 'Points', InternalName: 'Points', EntityPropertyName: 'Points', TypeAsString: 'Number' },
  { Title: 'EvaluationPoints', InternalName: 'EvaluationPoints', EntityPropertyName: 'EvaluationPoints', TypeAsString: 'Number' }
] };

function response(payload: unknown, status: number = 200): object {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => payload,
    text: async () => JSON.stringify(payload)
  };
}

function item(ids: string[], extra: Record<string, unknown> = {}): object {
  return {
    Id: 7, AuthorId: 42, UserKey: user.LoginName, Points: 125, EvaluationPoints: 0,
    FavoritesJson: JSON.stringify({ version: 1, scenarioIds: ids }),
    '@odata.etag': '"3"',
    ...extra
  };
}

function setup(): { client: { get: jest.Mock; post: jest.Mock }; service: FavoritesService } {
  const client = { get: jest.fn(), post: jest.fn() };
  client.get.mockResolvedValueOnce(response(user)).mockResolvedValueOnce(response(fields));
  return { client, service: new FavoritesService(client, webUrl) };
}

describe('Favorite data and concurrent changes', () => {
  it('validates versioned JSON and deduplicates IDs', () => {
    expect(parseFavoritesJson('{"version":1,"scenarioIds":["MS00001","MS00001"]}')).toEqual(['MS00001']);
    for (const value of ['', '{}', '{"version":2,"scenarioIds":[]}', '{"version":1,"scenarioIds":[123]}', '{']) {
      expect(() => parseFavoritesJson(value)).toThrow();
    }
  });

  it('preserves another device additions and applies explicit removals', () => {
    expect(mergeFavoriteChanges(['MS00001'], ['MS00002'], ['MS00001', 'MS00003'])).toEqual(['MS00003', 'MS00002']);
    expect(mergeFavoriteChanges([], ['MS00001'], ['MS00002'])).toEqual(['MS00002', 'MS00001']);
    expect(mergeFavoriteChanges(['MS00001'], ['MS00001'], [])).toEqual([]);
  });
});

describe('SharePoint favorites', () => {
  it('loads only rows authored by the signed-in user, even for a list administrator', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'])] }));
    await expect(service.load()).resolves.toEqual({ ids: ['MS00001'], exists: true, points: 125, evaluationPoints: 0 });
    expect(client.get.mock.calls[2][0]).toContain('$filter=AuthorId eq 42');
    expect(client.get.mock.calls[2][0]).toContain(',Points,EvaluationPoints&$filter');
    expect(client.get.mock.calls[2][0]).toContain("GetList('/sites/-TMC-/Lists/List2')");
    expect(service.migrationKey).toContain(user.LoginName);
  });

  it('creates a user row with zero points and empty favorites on first visit', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [] }));
    client.post.mockResolvedValueOnce(response({}, 201));
    client.get.mockResolvedValueOnce(response({ value: [item([], { Points: 0 })] }));
    await expect(service.load()).resolves.toEqual({ ids: [], exists: false, points: 0, evaluationPoints: 0 });
    const body = JSON.parse(client.post.mock.calls[0][2].body);
    expect(body.UserKey).toBe(user.LoginName);
    expect(body.Points).toBe(0);
    expect(body.EvaluationPoints).toBe(0);
    expect(JSON.parse(body.FavoritesJson)).toEqual({ version: 1, scenarioIds: [] });
  });

  it('initializes legacy user rows with zero points using their ETag', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'], { Points: null })] }));
    client.post.mockResolvedValueOnce(response({}, 204));
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'], { Points: 0 })] }));
    await expect(service.load()).resolves.toEqual({ ids: ['MS00001'], exists: true, points: 0, evaluationPoints: 0 });
    expect(client.post.mock.calls[0][2].headers['IF-MATCH']).toBe('"3"');
    expect(JSON.parse(client.post.mock.calls[0][2].body)).toEqual({ Points: 0 });
  });

  it('identifies which point column is still unavailable after initialization', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'], { Points: null })] }));
    client.post.mockResolvedValueOnce(response({}, 204));
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'], { Points: 0, EvaluationPoints: undefined })] }));
    await expect(service.load()).rejects.toThrow('EvaluationPoints列の値を確認できませんでした');
  });

  it('rejects invalid point values and missing Points column', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item([], { Points: -1 })] }));
    await expect(service.load()).rejects.toThrow('0以上の整数');
    const missing = setup();
    missing.client.get.mockReset().mockResolvedValueOnce(response(user)).mockResolvedValueOnce(response({ value: fields.value.slice(0, 2) }));
    await expect(missing.service.load()).rejects.toThrow('Points列');
  });

  it('creates the initial row with migrated IDs and verifies the result', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [] }));
    client.post.mockResolvedValueOnce(response({}, 201));
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001', 'MS00002'])] }));
    await expect(service.save([], ['MS00001', 'MS00002'])).resolves.toEqual(['MS00001', 'MS00002']);
    const body = JSON.parse(client.post.mock.calls[0][2].body);
    expect(body.UserKey).toBe(user.LoginName);
    expect(JSON.parse(body.FavoritesJson)).toEqual({ version: 1, scenarioIds: ['MS00001', 'MS00002'] });
    expect(body.Points).toBe(0);
    expect(body.EvaluationPoints).toBe(0);
    expect(client.post.mock.calls[0][2].headers['IF-MATCH']).toBeUndefined();
  });

  it('retries ETag conflicts and merges changes without losing another device additions', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001'])] }));
    client.post.mockResolvedValueOnce(response({}, 412));
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00001', 'MS00003'], { '@odata.etag': '"4"' })] }));
    client.post.mockResolvedValueOnce(response({}, 204));
    client.get.mockResolvedValueOnce(response({ value: [item(['MS00003', 'MS00002'])] }));
    await expect(service.save(['MS00001'], ['MS00002'])).resolves.toEqual(['MS00003', 'MS00002']);
    const options = client.post.mock.calls[1][2];
    expect(options.headers['IF-MATCH']).toBe('"4"');
    expect(options.headers['X-HTTP-Method']).toBe('MERGE');
    expect(JSON.parse(JSON.parse(options.body).FavoritesJson).scenarioIds).toEqual(['MS00003', 'MS00002']);
    expect(JSON.parse(options.body).Points).toBe(125);
    expect(JSON.parse(options.body).EvaluationPoints).toBe(0);
  });

  it('reconciles newly saved evaluation awards once while preserving other points', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item([], { Points: 125, EvaluationPoints: 2 })] }));
    client.post.mockResolvedValueOnce(response({}, 204));
    client.get.mockResolvedValueOnce(response({ value: [item([], { Points: 126, EvaluationPoints: 3 })] }));
    await expect(service.reconcileEvaluationPoints(3)).resolves.toBe(126);
    expect(JSON.parse(client.post.mock.calls[0][2].body)).toEqual({ Points: 126, EvaluationPoints: 3 });
    expect(client.post.mock.calls[0][2].headers['IF-MATCH']).toBe('"3"');
  });

  it('does not add points again when an evaluation has already been reconciled', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item([], { Points: 126, EvaluationPoints: 3 })] }));
    await expect(service.reconcileEvaluationPoints(3)).resolves.toBe(126);
    expect(client.post).not.toHaveBeenCalled();
  });

  it('reports repeated conflicts instead of overwriting with a wildcard ETag', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValue(response({ value: [item(['MS00001'])] }));
    client.post.mockResolvedValue(response({}, 412));
    await expect(service.save(['MS00001'], [])).rejects.toThrow('再試行');
    expect(client.post).toHaveBeenCalledTimes(3);
  });

  it('rejects duplicates, including simultaneous initial creation', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item([]), item([])] }));
    await expect(service.load()).rejects.toThrow('重複');
    expect(client.post).not.toHaveBeenCalled();
    const second = setup();
    second.client.get.mockResolvedValueOnce(response({ value: [] }));
    second.client.post.mockResolvedValueOnce(response({}, 201));
    second.client.get.mockResolvedValueOnce(response({ value: [item([]), item([])] }));
    await expect(second.service.save([], ['MS00001'])).rejects.toThrow('重複');
  });

  it('rejects mismatched ownership, missing ETags and corrupt JSON', async () => {
    for (const extra of [{ AuthorId: 43 }, { UserKey: 'another-user' }, { '@odata.etag': undefined }, { FavoritesJson: '{}' }]) {
      const { client, service } = setup();
      client.get.mockResolvedValueOnce(response({ value: [item([], extra)] }));
      await expect(service.load()).rejects.toThrow();
      expect(client.post).not.toHaveBeenCalled();
    }
  });

  it('reports permission, missing list and missing field errors', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({}, 404));
    await expect(service.load()).rejects.toThrow('HTTP 404');
    const denied = setup();
    denied.client.get.mockResolvedValueOnce(response({ value: [] }));
    denied.client.post.mockResolvedValueOnce(response({}, 403));
    await expect(denied.service.save([], [])).rejects.toThrow('HTTP 403');
    const wrong = setup();
    wrong.client.get.mockReset().mockResolvedValueOnce(response(user)).mockResolvedValueOnce(response({ value: [] }));
    await expect(wrong.service.load()).rejects.toThrow('UserKey');
  });
});
