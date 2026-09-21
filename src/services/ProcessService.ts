import type { SPFI } from '@pnp/sp';
import type { IProcess } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';
import '@pnp/sp/site-users/web';

const LIST_TITLE = 'Procesos';
const CREATE_LIST_TITLE = 'Cartera de procesos';

export class ProcessService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async getProcessesByGroup(groupId: number): Promise<IProcess[]> {
    const codeField = await this.sharePoint.resolveField(LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['NombreProceso']);
    const descriptionField = await this.sharePoint.resolveField(LIST_TITLE, ['DescripcionProceso', 'DescripciónProceso']);
    const groupField = await this.sharePoint.resolveField(LIST_TITLE, ['GrupoProcesoLookup', 'Grupo de procesos']);
    const departmentField = await this.sharePoint.resolveField(LIST_TITLE, ['DepartamentoLookup', 'Departamento']);
    const ownerField = await this.sharePoint.resolveField(LIST_TITLE, ['PropietarioProceso', 'Propietario del proceso']);
    const statusField = await this.sharePoint.resolveField(LIST_TITLE, ['Status', 'Estado']);
    const typeField = await this.sharePoint.resolveField(LIST_TITLE, ['TipoProceso', 'Tipo de proceso']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select(
        'ID',
        codeField.InternalName,
        nameField.InternalName,
        descriptionField.InternalName,
        `${groupField.InternalName}Id`,
        `${departmentField.InternalName}Id`,
        `${ownerField.InternalName}/Title`,
        statusField.InternalName,
        typeField.InternalName
      )
      .expand(ownerField.InternalName)
      .filter(`${groupField.InternalName}Id eq ${groupId}`)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        code: SharePointService.getText(item, codeField.InternalName),
        name: SharePointService.getText(item, nameField.InternalName),
        description: SharePointService.getText(item, descriptionField.InternalName),
        groupId: SharePointService.getId(item, `${groupField.InternalName}Id`),
        departmentId: SharePointService.getId(item, `${departmentField.InternalName}Id`) || undefined,
        owner: SharePointService.getLookupText(item, ownerField.InternalName) || undefined,
        status: SharePointService.getText(item, statusField.InternalName) || undefined,
        type: SharePointService.getText(item, typeField.InternalName) || undefined
      }))
      .filter(process => process.id > 0)
      .sort((left, right) => `${left.code} ${left.name}`.localeCompare(`${right.code} ${right.name}`));
  }

  public async createProcess(
    code: string,
    name: string,
    description: string,
    ownerEmail: string,
    groupId: number,
    departmentId: number
  ): Promise<void> {
    const codeField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['NombreProceso']);
    const descriptionField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['DescripcionProceso', 'DescripciónProceso']);
    const ownerField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['PropietarioProceso', 'Propietario del proceso']);
    const groupField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['GrupoProcesoLookup', 'Grupo de procesos']);
    const departmentField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['DepartamentoLookup', 'Departamento']);
    const owner = await this.sharePoint.sp.web.ensureUser(ownerEmail);
    const values: Record<string, unknown> = {
      [codeField.InternalName]: code,
      [nameField.InternalName]: name,
      [descriptionField.InternalName]: description,
      [ownerField.InternalName]: owner.Id,
      [`${groupField.InternalName}Id`]: groupId,
      [`${departmentField.InternalName}Id`]: departmentId
    };
    await this.sharePoint.sp.web.lists.getByTitle(CREATE_LIST_TITLE).items.add(values);
  }
}