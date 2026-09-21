import type { SPFI } from '@pnp/sp';
import type { IProcessGroup } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';

const LIST_TITLE = 'Grupo de procesos';

export class ProcessGroupService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async getGroupsByDepartment(departmentId: number): Promise<IProcessGroup[]> {
    const codeField = await this.sharePoint.resolveField(LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['NombreGrupo']);
    const departmentField = await this.sharePoint.resolveField(LIST_TITLE, ['Departamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select('ID', codeField.InternalName, nameField.InternalName, `${departmentField.InternalName}Id`)
      .filter(`${departmentField.InternalName}Id eq ${departmentId}`)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        code: SharePointService.getText(item, codeField.InternalName),
        name: SharePointService.getText(item, nameField.InternalName),
        departmentId: SharePointService.getId(item, `${departmentField.InternalName}Id`)
      }))
      .filter(group => group.id > 0)
      .sort((left, right) => `${left.code} ${left.name}`.localeCompare(`${right.code} ${right.name}`));
  }

  public async createGroup(code: string, name: string, departmentId: number): Promise<void> {
    const codeField = await this.sharePoint.resolveField(LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['NombreGrupo']);
    const departmentField = await this.sharePoint.resolveField(LIST_TITLE, ['Departamento']);
    await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items.add({
      [codeField.InternalName]: code,
      [nameField.InternalName]: name,
      [`${departmentField.InternalName}Id`]: departmentId
    });
  }

  public async getAllGroups(): Promise<IProcessGroup[]> {
    const codeField = await this.sharePoint.resolveField(LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['NombreGrupo']);
    const departmentField = await this.sharePoint.resolveField(LIST_TITLE, ['Departamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select('ID', codeField.InternalName, nameField.InternalName, `${departmentField.InternalName}Id`)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        code: SharePointService.getText(item, codeField.InternalName),
        name: SharePointService.getText(item, nameField.InternalName),
        departmentId: SharePointService.getId(item, `${departmentField.InternalName}Id`)
      }))
      .filter(group => group.id > 0)
      .sort((left, right) => `${left.code} ${left.name}`.localeCompare(`${right.code} ${right.name}`));
  }
}