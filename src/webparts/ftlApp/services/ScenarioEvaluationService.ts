import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IScenarioEvaluation {
  willHappen: -1 | 1;
  lifeImpact: -1 | 1;
  wantItCome: -1 | 1;
}

export interface IScenarioEvaluationSnapshot {
  evaluation?: IScenarioEvaluation;
  evaluationCount: number;
}

interface IUser {
  id: number;
  login: string;
  name: string;
}

interface IEvaluationItem {
  id: number;
  etag: string;
  scenarioId: string;
  evaluationKey: string;
  evaluation: IScenarioEvaluation;
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('SharePointから想定した形式の評価データを取得できませんでした。');
  }
  return value as Record<string, unknown>;
}

function score(value: unknown, fieldName: string): -1 | 1 {
  if (value !== -1 && value !== 1) {
    throw new Error(`${fieldName}列には-1または1を設定してください。`);
  }
  return value;
}

export default class ScenarioEvaluationService {
  private readonly _listUrl: string;
  private _user: IUser | undefined;
  private _fields: Record<string, string> | undefined;

  public constructor(private readonly _client: Pick<SPHttpClient, 'get' | 'post'>, webUrl: string) {
    this._listUrl = `${webUrl}/_api/web/GetList('/sites/-TMC-/Lists/List3')`;
  }

  private async _check(response: SPHttpClientResponse, action: string): Promise<void> {
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 400);
      throw new Error(`${action}に失敗しました（HTTP ${response.status}）。評価リストの列・権限を確認してください。${detail}`);
    }
  }

  private async _get(url: string): Promise<Record<string, unknown>> {
    const response = await this._client.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=minimalmetadata' }
    });
    await this._check(response, 'シナリオ評価の取得');
    return record(await response.json());
  }

  private async _initialize(): Promise<void> {
    if (this._user && this._fields) return;
    const user = await this._get(`${this._listUrl.split('/_api/')[0]}/_api/web/currentuser?$select=Id,LoginName,Title`);
    if (typeof user.Id !== 'number' || typeof user.LoginName !== 'string' || !user.LoginName || typeof user.Title !== 'string') {
      throw new Error('現在のユーザー情報を取得できませんでした。');
    }
    const fields = await this._get(`${this._listUrl}/fields?$select=Title,InternalName,EntityPropertyName,TypeAsString`);
    if (!Array.isArray(fields.value)) throw new Error('評価リストの列情報が不正です。');
    const fieldRecords = fields.value.map(record);
    const definitions: Array<[string, string]> = [
      ['UserKey', 'Text'],
      ['ScenarioId', 'Text'],
      ['WillHappen', 'Number'],
      ['LifeImpact', 'Number'],
      ['WantItCome', 'Number'],
      ['EvaluationKey', 'Text']
    ];
    const resolved: Record<string, string> = {};
    for (const [name, type] of definitions) {
      const field = fieldRecords.find(item => item.InternalName === name || item.Title === name);
      if (!field || field.TypeAsString !== type) throw new Error(`${name}列がありません、または列の種類が不正です。`);
      const property = field.EntityPropertyName || field.InternalName;
      if (typeof property !== 'string' || !/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(property)) {
        throw new Error(`${name}列の内部名を取得できませんでした。`);
      }
      resolved[name] = property;
    }
    this._user = { id: user.Id, login: user.LoginName, name: user.Title };
    this._fields = resolved;
  }

  private _parseItem(value: unknown): IEvaluationItem {
    if (!this._user || !this._fields) throw new Error('評価リストの初期化が完了していません。');
    const item = record(value);
    const scenarioId = item[this._fields.ScenarioId];
    const evaluationKey = item[this._fields.EvaluationKey];
    const etag = item['@odata.etag'] || item['odata.etag'];
    if (typeof item.Id !== 'number' || item.AuthorId !== this._user.id ||
      item[this._fields.UserKey] !== this._user.login || typeof scenarioId !== 'string' ||
      typeof evaluationKey !== 'string' || typeof etag !== 'string' || !etag) {
      throw new Error('評価行の作成者、UserKey、シナリオID、EvaluationKeyまたはETagが不正です。');
    }
    return {
      id: item.Id,
      etag,
      scenarioId,
      evaluationKey,
      evaluation: {
        willHappen: score(item[this._fields.WillHappen], 'WillHappen'),
        lifeImpact: score(item[this._fields.LifeImpact], 'LifeImpact'),
        wantItCome: score(item[this._fields.WantItCome], 'WantItCome')
      }
    };
  }

  private async _readAll(): Promise<IEvaluationItem[]> {
    await this._initialize();
    if (!this._user || !this._fields) throw new Error('評価リストの初期化が完了していません。');
    const select = `Id,AuthorId,${this._fields.UserKey},${this._fields.ScenarioId},${this._fields.EvaluationKey},${this._fields.WillHappen},${this._fields.LifeImpact},${this._fields.WantItCome}`;
    const items: IEvaluationItem[] = [];
    let url = `${this._listUrl}/items?$select=${select}&$filter=AuthorId%20eq%20${this._user.id}&$top=5000`;
    while (url) {
      const data = await this._get(url);
      if (!Array.isArray(data.value)) throw new Error('評価リストの取得結果が不正です。');
      items.push(...data.value.map(value => this._parseItem(value)));
      const next = data['@odata.nextLink'] || data['odata.nextLink'];
      if (next !== undefined && typeof next !== 'string') throw new Error('評価リストの続きURLが不正です。');
      url = typeof next === 'string' ? next : '';
    }
    return items;
  }

  public async load(scenarioId: string): Promise<IScenarioEvaluationSnapshot> {
    if (!/^MS\d+$/.test(scenarioId)) throw new Error('シナリオIDが不正です。');
    if (!this._user) await this._initialize();
    const items = await this._readAll();
    const evaluationKey = `${this._user?.id}:${scenarioId}`;
    const matching = items.filter(item => item.evaluationKey === evaluationKey);
    if (matching.length > 1) throw new Error('同じシナリオの評価が重複しています。管理者に確認してください。');
    return {
      evaluation: matching[0]?.evaluation,
      evaluationCount: new Set(items.map(item => item.scenarioId)).size
    };
  }

  public async save(scenarioId: string, evaluation: IScenarioEvaluation): Promise<IScenarioEvaluationSnapshot> {
    if (!/^MS\d+$/.test(scenarioId)) throw new Error('シナリオIDが不正です。');
    const validEvaluation: IScenarioEvaluation = {
      willHappen: score(evaluation.willHappen, 'WillHappen'),
      lifeImpact: score(evaluation.lifeImpact, 'LifeImpact'),
      wantItCome: score(evaluation.wantItCome, 'WantItCome')
    };
    for (let attempt = 0; attempt < 3; attempt++) {
      const items = await this._readAll();
      if (!this._user || !this._fields) throw new Error('評価リストの初期化が完了していません。');
      const evaluationKey = `${this._user.id}:${scenarioId}`;
      const matching = items.filter(item => item.evaluationKey === evaluationKey);
      if (matching.length > 1) throw new Error('同じシナリオの評価が重複しています。管理者に確認してください。');
      const existing = matching[0];
      const body = {
        Title: scenarioId,
        [this._fields.UserKey]: this._user.login,
        [this._fields.ScenarioId]: scenarioId,
        [this._fields.EvaluationKey]: evaluationKey,
        [this._fields.WillHappen]: validEvaluation.willHappen,
        [this._fields.LifeImpact]: validEvaluation.lifeImpact,
        [this._fields.WantItCome]: validEvaluation.wantItCome
      };
      const response = await this._client.post(
        existing ? `${this._listUrl}/items(${existing.id})` : `${this._listUrl}/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=minimalmetadata',
            'Content-Type': 'application/json;odata=nometadata',
            ...(existing ? { 'IF-MATCH': existing.etag, 'X-HTTP-Method': 'MERGE' } : {})
          },
          body: JSON.stringify(body)
        }
      );
      if (response.status === 412) continue;
      if (!response.ok && !existing) {
        const afterConflict = await this._readAll();
        if (afterConflict.some(item => item.evaluationKey === evaluationKey)) continue;
      }
      await this._check(response, 'シナリオ評価の保存');
      const verified = await this._readAll();
      const saved = verified.filter(item => item.evaluationKey === evaluationKey);
      if (saved.length !== 1) throw new Error('保存した評価を一意に確認できませんでした。');
      return {
        evaluation: saved[0].evaluation,
        evaluationCount: new Set(verified.map(item => item.scenarioId)).size
      };
    }
    throw new Error('別の端末で評価が更新されているため保存できませんでした。再試行してください。');
  }
}
