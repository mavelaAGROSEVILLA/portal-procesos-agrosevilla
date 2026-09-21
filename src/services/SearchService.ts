import type { SPFI } from '@pnp/sp';
import type { ISearchResult } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';

const SEARCH_LIMIT = 8;

export class SearchService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async search(term: string): Promise<ISearchResult[]> {
    const normalizedTerm = term.trim();
    if (normalizedTerm.length < 2) {
      return [];
    }

    const [departments, groups, processes] = await Promise.all([
      this.searchDepartments(normalizedTerm).catch(error => this.handleSearchError('departamentos', error)),
      this.searchGroups(normalizedTerm).catch(error => this.handleSearchError('grupos', error)),
      this.searchProcesses(normalizedTerm).catch(error => this.handleSearchError('procesos', error))
    ]);
    return [...departments, ...groups, ...processes].slice(0, SEARCH_LIMIT);
  }

  private async searchDepartments(term: string): Promise<ISearchResult[]> {
    const listTitle = 'Departamentos';
    const codeField = await this.sharePoint.resolveField(listTitle, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(listTitle, ['NombreDepartamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(listTitle).items
      .select('ID', codeField.InternalName, nameField.InternalName)
      .filter(this.buildFilter([codeField.InternalName, nameField.InternalName], term))
      .top(SEARCH_LIMIT)();
    return (items as ISharePointItem[]).map(item => ({
      id: item.ID || 0,
      kind: 'department' as const,
      code: SharePointService.getText(item, codeField.InternalName),
      name: SharePointService.getText(item, nameField.InternalName)
    }));
  }

  private async searchGroups(term: string): Promise<ISearchResult[]> {
    const listTitle = 'Grupo de procesos';
    const codeField = await this.sharePoint.resolveField(listTitle, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(listTitle, ['NombreGrupo']);
    const departmentField = await this.sharePoint.resolveField(listTitle, ['Departamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(listTitle).items
      .select('ID', codeField.InternalName, nameField.InternalName, `${departmentField.InternalName}Id`)
      .filter(this.buildFilter([codeField.InternalName, nameField.InternalName], term))
      .top(SEARCH_LIMIT)();
    return (items as ISharePointItem[]).map(item => ({
      id: item.ID || 0,
      kind: 'group' as const,
      code: SharePointService.getText(item, codeField.InternalName),
      name: SharePointService.getText(item, nameField.InternalName),
      departmentId: SharePointService.getId(item, `${departmentField.InternalName}Id`)
    }));
  }

  private async searchProcesses(term: string): Promise<ISearchResult[]> {
    const listTitle = 'Procesos';
    const codeField = await this.sharePoint.resolveField(listTitle, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(listTitle, ['NombreProceso']);
    const descriptionField = await this.sharePoint.resolveField(listTitle, ['DescripcionProceso', 'DescripciónProceso']);
    const groupField = await this.sharePoint.resolveField(listTitle, ['GrupoProcesoLookup', 'Grupo de procesos']);
    const departmentField = await this.sharePoint.resolveField(listTitle, ['DepartamentoLookup', 'Departamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(listTitle).items
      .select('ID', codeField.InternalName, nameField.InternalName, descriptionField.InternalName, `${groupField.InternalName}Id`, `${departmentField.InternalName}Id`)
      .filter(this.buildFilter([codeField.InternalName, nameField.InternalName], term))
      .top(SEARCH_LIMIT)();
    return (items as ISharePointItem[]).map(item => ({
      id: item.ID || 0,
      kind: 'process' as const,
      code: SharePointService.getText(item, codeField.InternalName),
      name: SharePointService.getText(item, nameField.InternalName),
      description: SharePointService.getText(item, descriptionField.InternalName),
      groupId: SharePointService.getId(item, `${groupField.InternalName}Id`),
      departmentId: SharePointService.getId(item, `${departmentField.InternalName}Id`)
    }));
  }

  private buildFilter(fieldNames: string[], term: string): string {
    const escapedTerm = term.replace(/'/g, "''");
    return fieldNames.map(fieldName => `substringof('${escapedTerm}',${fieldName})`).join(' or ');
  }

  private handleSearchError(level: string, error: unknown): ISearchResult[] {
    console.error(`[ProcessView] Error en la búsqueda de ${level}`, error);
    return [];
  }
}