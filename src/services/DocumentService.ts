import type { SPFI } from '@pnp/sp';
import type { IDocumentoProceso } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';

const LIST_TITLE = 'Documentos de procesos';

export class DocumentService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async getDocumentsByProcess(processId: number): Promise<IDocumentoProceso[]> {
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['Name', 'Nombre']);
    const processField = await this.sharePoint.resolveField(LIST_TITLE, ['ProcesoRelated']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select('ID', nameField.InternalName, 'FileRef', `${processField.InternalName}Id`)
      .filter(`${processField.InternalName}Id eq ${processId}`)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        name: SharePointService.getText(item, nameField.InternalName),
        processId: SharePointService.getId(item, `${processField.InternalName}Id`) || undefined,
        url: SharePointService.getText(item, 'FileRef') || undefined
      }))
      .filter(document => document.id > 0)
      .sort((left, right) => left.name.localeCompare(right.name));
  }
}