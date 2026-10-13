import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IFavoritesSnapshot {
  ids: string[];
  exists: boolean;
}

interface IUser {
  id: number;
  login: string;
  name: string;
}

interface IFavoriteItem {
  id: number;
  etag: string;
  ids: string[];
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('SharePointから想定した形式のデータを取得できませんでした。');
  }
  return value as Record<string, unknown>;
}

export function parseFavoritesJson(value: unknown): string[] {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error('FavoritesJsonが空、または文字列ではありません。');
  }
  const data = record(JSON.parse(value));
  if (data.version !== 1 || !Array.isArray(data.scenarioIds) ||
    !data.scenarioIds.every(id => typeof id === 'string' && /^MS\d+$/.test(id))) {
    throw new Error('FavoritesJsonの形式が不正です。version: 1とscenarioIdsの配列が必要です。');
  }
  return data.scenarioIds.filter((id, index, ids) => ids.indexOf(id) === index);
}

export function mergeFavoriteChanges(base: readonly string[], desired: readonly string[], latest: readonly string[]): string[] {
  const removed = base.filter(id => desired.indexOf(id) === -1);
  const added = desired.filter(id => base.indexOf(id) === -1);
  return latest.filter(id => removed.indexOf(id) === -1)
    .concat(added.filter(id => latest.indexOf(id) === -1));
}

export default class FavoritesService {
  private readonly _listUrl: string;
  private _user: IUser | undefined;
  private _keyField: string = 'UserKey';
  private _jsonField: string = 'FavoritesJson';
  private _initialized: boolean = false;

  public constructor(private readonly _client: Pick<SPHttpClient, 'get' | 'post'>, private readonly _webUrl: string) {
    this._listUrl = `${_webUrl}/_api/web/GetList('/sites/-TMC-/Lists/List2')`;
  }

  private async _check(response: SPHttpClientResponse, action: string): Promise<void> {
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 400);
      throw new Error(`${action}に失敗しました（HTTP ${response.status}）。リストの列・権限を確認してください。${detail}`);
    }
  }

  private async _get(url: string): Promise<Record<string, unknown>> {
    const response = await this._client.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=minimalmetadata' }
    });
    await this._check(response, 'お気に入りの取得');
    return record(await response.json());
  }

  private async _initialize(): Promise<void> {
    if (this._initialized) return;
    const user = await this._get(`${this._webUrl}/_api/web/currentuser?$select=Id,LoginName,Title`);
    if (typeof user.Id !== 'number' || typeof user.LoginName !== 'string' || !user.LoginName ||
      typeof user.Title !== 'string') {
      throw new Error('現在のユーザー情報を取得できませんでした。');
    }
    this._user = { id: user.Id, login: user.LoginName, name: user.Title };
    const fields = await this._get(`${this._listUrl}/fields?$select=Title,InternalName,EntityPropertyName,TypeAsString`);
    if (!Array.isArray(fields.value)) throw new Error('お気に入りリストの列情報が不正です。');
    const fieldRecords = fields.value.map(record);
    const resolveField = (name: string, type: string): string => {
      const field = fieldRecords.find(item => item.InternalName === name || item.Title === name);
      if (!field || field.TypeAsString !== type) throw new Error(`${name}列がありません、または列の種類が不正です。`);
      const property = field.EntityPropertyName || field.InternalName;
      if (typeof property !== 'string' || !/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(property)) {
        throw new Error(`${name}列の内部名を取得できませんでした。`);
      }
      return property;
    };
    this._keyField = resolveField('UserKey', 'Text');
    this._jsonField = resolveField('FavoritesJson', 'Note');
    this._initialized = true;
  }

  public get migrationKey(): string {
    if (!this._user) throw new Error('ユーザー情報の読み込みが完了していません。');
    return `ftl-app:favorite-migration:v1:${this._webUrl}:${this._user.login}`;
  }

  private async _readItem(): Promise<IFavoriteItem | undefined> {
    await this._initialize();
    if (!this._user) throw new Error('ユーザー情報がありません。');
    const data = await this._get(`${this._listUrl}/items?$select=Id,AuthorId,${this._keyField},${this._jsonField}&$filter=AuthorId eq ${this._user.id}&$orderby=Id&$top=2`);
    if (!Array.isArray(data.value)) throw new Error('お気に入りの取得結果が不正です。');
    if (data.value.length > 1 || data['@odata.nextLink'] || data['odata.nextLink']) {
      throw new Error('本人のお気に入り行が重複しています。管理者に行の統合を依頼してください。データは上書きしていません。');
    }
    if (data.value.length === 0) return undefined;
    const item = record(data.value[0]);
    if (typeof item.Id !== 'number' || item.AuthorId !== this._user.id ||
      item[this._keyField] !== this._user.login) {
      throw new Error('お気に入り行の作成者またはUserKeyが現在のユーザーと一致しません。');
    }
    const etag = item['@odata.etag'] || item['odata.etag'];
    if (typeof etag !== 'string' || !etag) throw new Error('お気に入り行のETagを取得できませんでした。');
    return { id: item.Id, etag, ids: parseFavoritesJson(item[this._jsonField]) };
  }

  public async load(): Promise<IFavoritesSnapshot> {
    const item = await this._readItem();
    return { ids: item ? item.ids : [], exists: item !== undefined };
  }

  public async save(base: readonly string[], desired: readonly string[]): Promise<string[]> {
    // Reapply only this screen's additions/removals to the latest server state.
    for (let attempt = 0; attempt < 3; attempt++) {
      const item = await this._readItem();
      if (!this._user) throw new Error('ユーザー情報がありません。');
      const ids = mergeFavoriteChanges(base, desired, item ? item.ids : []);
      const body = JSON.stringify({
        Title: this._user.name,
        [this._keyField]: this._user.login,
        [this._jsonField]: JSON.stringify({ version: 1, scenarioIds: ids })
      });
      const response = await this._client.post(
        item ? `${this._listUrl}/items(${item.id})` : `${this._listUrl}/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=minimalmetadata',
            'Content-Type': 'application/json;odata=nometadata',
            ...(item ? { 'IF-MATCH': item.etag, 'X-HTTP-Method': 'MERGE' } : {})
          },
          body
        }
      );
      if (response.status === 412 && item) continue;
      await this._check(response, 'お気に入りの保存');
      // Also detect simultaneous first-time creation while uniqueness is disabled.
      const verified = await this._readItem();
      if (!verified) throw new Error('保存したお気に入り行を確認できませんでした。');
      return verified.ids;
    }
    throw new Error('別の端末で更新が続いているため保存できませんでした。再試行してください。');
  }
}
