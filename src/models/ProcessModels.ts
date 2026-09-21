export interface IDepartment {
  id: number;
  code: string;
  name: string;
}

export interface IProcessGroup {
  id: number;
  code: string;
  name: string;
  departmentId: number;
}

export interface IProcess {
  id: number;
  code: string;
  name: string;
  description: string;
  groupId: number;
  departmentId?: number;
  owner?: string;
  status?: string;
  type?: string;
}

export interface IDocumentoProceso {
  id: number;
  name: string;
  title?: string;
  processId?: number;
  url?: string;
  documentType?: string;
}

export interface IDocumentoContenido {
  name: string;
  url: string;
  isFolder: boolean;
}

export type SearchResultKind = 'department' | 'group' | 'process';

export interface ISearchResult {
  id: number;
  kind: SearchResultKind;
  code: string;
  name: string;
  description?: string;
  departmentId?: number;
  groupId?: number;
}