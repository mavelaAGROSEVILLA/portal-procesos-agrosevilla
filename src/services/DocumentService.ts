import type { SPFI } from '@pnp/sp';
import type { IDocumentoProceso } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';
import '@pnp/sp/folders';
import '@pnp/sp/files';

const LIST_TITLE = 'Documentos de procesos';

export class DocumentService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async getDocumentsByProcess(processId: number): Promise<IDocumentoProceso[]> {
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['Name', 'Nombre']);
    const processField = await this.sharePoint.resolveField(LIST_TITLE, ['ProcesoRelated']);
    const documentTypeField = await this.sharePoint.resolveField(LIST_TITLE, ['TipoDocumental', 'Tipo documental']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select('ID', nameField.InternalName, documentTypeField.InternalName, 'FileRef', `${processField.InternalName}Id`)
      .filter(`${processField.InternalName}Id eq ${processId}`)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        name: SharePointService.getText(item, nameField.InternalName),
        processId: SharePointService.getId(item, `${processField.InternalName}Id`) || undefined,
        url: SharePointService.getText(item, 'FileRef') || undefined,
        documentType: SharePointService.getText(item, documentTypeField.InternalName) || undefined
      }))
      .filter(document => document.id > 0)
      .sort((left, right) => left.name.localeCompare(right.name));
  }

  public async getFolderContents(folderUrl: string): Promise<Array<{ name: string; url: string; isFolder: boolean }>> {
    const folder = this.sharePoint.sp.web.getFolderByServerRelativePath(folderUrl);
    const [files, folders] = await Promise.all([
      folder.files.select('Name', 'ServerRelativeUrl')(),
      folder.folders.select('Name', 'ServerRelativeUrl')()
    ]);
    const fileItems = (files as unknown as ISharePointItem[]).map(file => ({
      name: SharePointService.getText(file, 'Name'),
      url: SharePointService.getText(file, 'ServerRelativeUrl'),
      isFolder: false
    }));
    const folderItems = (folders as unknown as ISharePointItem[]).map(subfolder => ({
      name: SharePointService.getText(subfolder, 'Name'),
      url: SharePointService.getText(subfolder, 'ServerRelativeUrl'),
      isFolder: true
    }));
    return [...folderItems, ...fileItems].filter(item => item.name && item.url).sort((left, right) => left.name.localeCompare(right.name));
  }
}