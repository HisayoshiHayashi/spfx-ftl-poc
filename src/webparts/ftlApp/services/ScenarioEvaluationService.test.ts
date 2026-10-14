jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));

import ScenarioEvaluationService from './ScenarioEvaluationService';

const webUrl = 'https://example.sharepoint.com/sites/-TMC-';
const user = { Id: 42, LoginName: 'i:0#.f|membership|user@example.com', Title: 'Test User' };
const fields = { value: [
  { Title: 'UserKey', InternalName: 'UserKey', EntityPropertyName: 'UserKey', TypeAsString: 'Text' },
  { Title: 'ScenarioId', InternalName: 'ScenarioId', EntityPropertyName: 'ScenarioId', TypeAsString: 'Text' },
  { Title: 'WillHappen', InternalName: 'WillHappen', EntityPropertyName: 'WillHappen', TypeAsString: 'Number' },
  { Title: 'LifeImpact', InternalName: 'LifeImpact', EntityPropertyName: 'LifeImpact', TypeAsString: 'Number' },
  { Title: 'WantItCome', InternalName: 'WantItCome', EntityPropertyName: 'WantItCome', TypeAsString: 'Number' },
  { Title: 'EvaluationKey', InternalName: 'EvaluationKey', EntityPropertyName: 'EvaluationKey', TypeAsString: 'Text' }
] };
const answer = { willHappen: -1 as const, lifeImpact: 1 as const, wantItCome: -1 as const };

function response(payload: unknown, status: number = 200): object {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => payload,
    text: async () => JSON.stringify(payload)
  };
}

function item(scenarioId: string, values: Record<string, unknown> = {}): object {
  return {
    Id: 7, AuthorId: 42, UserKey: user.LoginName, ScenarioId: scenarioId,
    EvaluationKey: `42:${scenarioId}`, WillHappen: -1, LifeImpact: 1, WantItCome: -1,
    '@odata.etag': '"3"',
    ...values
  };
}

function setup(): { client: { get: jest.Mock; post: jest.Mock }; service: ScenarioEvaluationService } {
  const client = { get: jest.fn(), post: jest.fn() };
  client.get.mockResolvedValueOnce(response(user)).mockResolvedValueOnce(response(fields));
  return { client, service: new ScenarioEvaluationService(client, webUrl) };
}

describe('SharePoint scenario evaluations', () => {
  it('loads the current user evaluation and counts distinct evaluated scenarios', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item('MS00001'), item('MS00002')] }));
    await expect(service.load('MS00001')).resolves.toEqual({ evaluation: answer, evaluationCount: 2 });
    expect(client.get.mock.calls[2][0]).toContain("GetList('/sites/-TMC-/Lists/List3')");
    expect(client.get.mock.calls[2][0]).toContain('$filter=AuthorId%20eq%2042');
  });

  it('creates a first evaluation with the signed-in user and a unique evaluation key', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [] }));
    client.post.mockResolvedValueOnce(response({}, 201));
    client.get.mockResolvedValueOnce(response({ value: [item('MS00001')] }));
    await expect(service.save('MS00001', answer)).resolves.toEqual({ evaluation: answer, evaluationCount: 1 });
    const body = JSON.parse(client.post.mock.calls[0][2].body);
    expect(body).toEqual({
      Title: 'MS00001',
      UserKey: user.LoginName,
      ScenarioId: 'MS00001',
      EvaluationKey: '42:MS00001',
      WillHappen: -1,
      LifeImpact: 1,
      WantItCome: -1
    });
  });

  it('updates an existing evaluation with its ETag without creating another row', async () => {
    const { client, service } = setup();
    client.get.mockResolvedValueOnce(response({ value: [item('MS00001')] }));
    client.post.mockResolvedValueOnce(response({}, 204));
    client.get.mockResolvedValueOnce(response({ value: [item('MS00001', { WillHappen: 1 })] }));
    await expect(service.save('MS00001', { ...answer, willHappen: 1 })).resolves.toEqual({
      evaluation: { ...answer, willHappen: 1 }, evaluationCount: 1
    });
    expect(client.post.mock.calls[0][0]).toContain('/items(7)');
    expect(client.post.mock.calls[0][2].headers['IF-MATCH']).toBe('"3"');
    expect(client.post.mock.calls[0][2].headers['X-HTTP-Method']).toBe('MERGE');
  });

  it('rejects duplicate records and invalid scores while allowing a non-unique key field', async () => {
    const duplicate = setup();
    duplicate.client.get.mockResolvedValueOnce(response({ value: [item('MS00001'), item('MS00001', { Id: 8 })] }));
    await expect(duplicate.service.load('MS00001')).rejects.toThrow('重複');

    const invalid = setup();
    invalid.client.get.mockResolvedValueOnce(response({ value: [item('MS00001', { WillHappen: 0 })] }));
    await expect(invalid.service.load('MS00001')).rejects.toThrow('-1または1');

    const nonUnique = setup();
    nonUnique.client.get.mockResolvedValueOnce(response({ value: [] }));
    await expect(nonUnique.service.load('MS00001')).resolves.toEqual({ evaluation: undefined, evaluationCount: 0 });
  });

  it('does not write an evaluation for malformed scenario IDs or answers', async () => {
    const { client, service } = setup();
    await expect(service.save('../MS00001', answer)).rejects.toThrow('シナリオID');
    expect(client.post).not.toHaveBeenCalled();
  });
});
