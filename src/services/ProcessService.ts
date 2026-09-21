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

  public async getProcessCountsByGroup(): Promise<Map<number, number>> {
    const groupField = await this.sharePoint.resolveField(LIST_TITLE, ['GrupoProcesoLookup', 'Grupo de procesos']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select(`${groupField.InternalName}Id`)();
    const counts = new Map<number, number>();
    (items as ISharePointItem[]).forEach(item => {
      const groupId = SharePointService.getId(item, `${groupField.InternalName}Id`);
      if (groupId > 0) {
        counts.set(groupId, (counts.get(groupId) || 0) + 1);
      }
    });
    return counts;
  }

  public async createProcess(
    code: string,
    name: string,
    description: string,
    ownerEmail: string,
    groupId: number,
    departmentId: number,
    departmentValue: string
  ): Promise<void> {
    const codeField = await this.sharePoint.tryResolveField(CREATE_LIST_TITLE, ['Codigo', 'Código']);
    const titleField = await this.sharePoint.resolveField(CREATE_LIST_TITLE, ['Title', 'Título', 'NombreProceso']);
    const descriptionField = await this.sharePoint.tryResolveField(CREATE_LIST_TITLE, ['Contexto', 'Contexto - Descripción', 'DescripcionProceso', 'DescripciónProceso']);
    const ownerField = await this.sharePoint.tryResolveField(CREATE_LIST_TITLE, ['PropietarioProceso', 'Propietario del proceso', 'Involucrados']);
    const groupField = await this.sharePoint.tryResolveField(CREATE_LIST_TITLE, ['GrupoProcesoLookup', 'Grupo de procesos']);
    const departmentField = await this.sharePoint.resolveFieldByTitleFragment(CREATE_LIST_TITLE, 'departamento');
    const departmentChoice = departmentField.Choices?.find(choice =>
      choice.toLowerCase().includes(departmentValue.toLowerCase()));
    const departmentStoredValue = departmentChoice || departmentValue;
    const publisherField = await this.sharePoint.tryResolveField(CREATE_LIST_TITLE, ['Involucrados', 'Publicado por', 'Publicador']);
    const currentUser = publisherField ? await this.sharePoint.sp.web.currentUser() : undefined;
    const values: Record<string, unknown> = {
      [titleField.InternalName]: name,
      [departmentField.InternalName]: departmentField.AllowMultipleValues ? [departmentStoredValue] : departmentStoredValue
    };
    if (codeField) {
      values[codeField.InternalName] = code;
    }
    if (descriptionField && description) {
      values[descriptionField.InternalName] = description;
    }
    if (groupField && groupId) {
      values[`${groupField.InternalName}Id`] = groupId;
    }
    if (ownerField && ownerEmail) {
      const owner = await this.sharePoint.sp.web.ensureUser(ownerEmail);
      values[`${ownerField.InternalName}Id`] = owner.Id;
    }
    if (publisherField && currentUser) {
      values[`${publisherField.InternalName}Id`] = publisherField.AllowMultipleValues
        ? { results: [currentUser.Id] }
        : currentUser.Id;
    }
    await this.sharePoint.sp.web.lists.getByTitle(CREATE_LIST_TITLE).items.add(values);
  }
}