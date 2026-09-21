import type { SPFI } from '@pnp/sp';
import type { IDepartment } from '../models/ProcessModels';
import { SharePointService, type ISharePointItem } from './SharePointService';

const LIST_TITLE = 'Departamentos';

export class DepartmentService {
  private readonly sharePoint: SharePointService;

  public constructor(sp: SPFI) {
    this.sharePoint = new SharePointService(sp);
  }

  public async getDepartments(): Promise<IDepartment[]> {
    const codeField = await this.sharePoint.resolveField(LIST_TITLE, ['Codigo', 'Código', 'Title', 'Título']);
    const nameField = await this.sharePoint.resolveField(LIST_TITLE, ['NombreDepartamento']);
    const items = await this.sharePoint.sp.web.lists.getByTitle(LIST_TITLE).items
      .select('ID', codeField.InternalName, nameField.InternalName)();

    return (items as ISharePointItem[])
      .map(item => ({
        id: item.ID || 0,
        code: SharePointService.getText(item, codeField.InternalName),
        name: SharePointService.getText(item, nameField.InternalName)
      }))
      .filter(department => department.id > 0)
      .sort((left, right) => `${left.code} ${left.name}`.localeCompare(`${right.code} ${right.name}`));
  }
}