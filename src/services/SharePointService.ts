import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

interface ISharePointField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  AllowMultipleValues: boolean;
}

export interface ISharePointItem {
  [key: string]: unknown;
  ID?: number;
}

export class SharePointService {
  private readonly fieldCache: Map<string, Map<string, ISharePointField>> = new Map();

  public constructor(public readonly sp: SPFI) {}

  public async resolveField(listTitle: string, candidates: string[]): Promise<ISharePointField> {
    const fields = this.fieldCache.get(listTitle) || await this.loadFields(listTitle);

    const field = candidates.reduce<ISharePointField | undefined>((match, candidate) => {
      if (match) {
        return match;
      }
      const candidateName = candidate.toLowerCase();
      return fields.get(candidateName) || Array.from(fields.values())
        .find(item => typeof item.Title === 'string' && item.Title.toLowerCase() === candidateName);
    }, undefined);

    if (!field) {
      throw new Error(`No se encontró el campo "${candidates.join(' / ')}" en la lista "${listTitle}".`);
    }

    if (process.env.NODE_ENV !== 'production') {
      console.debug('[ProcessView] Campo SharePoint resuelto', {
        listTitle,
        requestedNames: candidates,
        internalName: field.InternalName,
        title: field.Title,
        type: field.TypeAsString,
        allowMultipleValues: field.AllowMultipleValues
      });
    }
    return field;
  }

  private async loadFields(listTitle: string): Promise<Map<string, ISharePointField>> {
    const fieldItems = await this.sp.web.lists.getByTitle(listTitle).fields
      .select('InternalName', 'Title', 'TypeAsString', 'AllowMultipleValues')();
    const typedFields = fieldItems as unknown as ISharePointField[];
    const fields = new Map(typedFields
      .filter(field => typeof field.InternalName === 'string' && field.InternalName.length > 0)
      .map(field => [field.InternalName.toLowerCase(), field]));
    this.fieldCache.set(listTitle, fields);
    return fields;
  }

  public static getText(item: ISharePointItem, fieldName: string): string {
    const value = item[fieldName];
    return typeof value === 'string' ? value.trim() : '';
  }

  public static getLookupText(item: ISharePointItem, fieldName: string): string {
    const value = item[fieldName];
    if (typeof value === 'object' && value !== null && 'Title' in value) {
      const title = value.Title;
      return typeof title === 'string' ? title.trim() : '';
    }
    return '';
  }

  public static getId(item: ISharePointItem, fieldName: string): number {
    const value = item[fieldName];
    if (typeof value === 'number') {
      return value;
    }
    if (typeof value === 'string') {
      const parsed = Number(value);
      return Number.isFinite(parsed) ? parsed : 0;
    }
    return 0;
  }
}