import * as React from 'react';
import type { SPFI } from '@pnp/sp';
import { DepartmentService } from '../../services/DepartmentService';
import { ProcessGroupService } from '../../services/ProcessGroupService';
import { ProcessService } from '../../services/ProcessService';
import { DocumentService } from '../../services/DocumentService';
import type { IDepartment, IDocumentoContenido, IDocumentoProceso, IProcess, IProcessGroup, ISearchResult } from '../../models/ProcessModels';
import Department from '../Department/Department';
import ProcessGroup from '../ProcessGroup/ProcessGroup';
import Process from '../Process/Process';
import ProcessDetail from '../Process/ProcessDetail';
import CreateItemPanel, { type CreateItemMode } from './CreateItemPanel';
import GlobalSearch from './GlobalSearch';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IHomeProps {
  sp: SPFI;
}

const Home: React.FunctionComponent<IHomeProps> = ({ sp }) => {
  const [departments, setDepartments] = React.useState<IDepartment[]>([]);
  const [selectedDepartmentId, setSelectedDepartmentId] = React.useState<number>();
  const [selectedGroupId, setSelectedGroupId] = React.useState<number>();
  const [selectedProcess, setSelectedProcess] = React.useState<IProcess>();
  const [createMode, setCreateMode] = React.useState<CreateItemMode>();
  const [isAddMenuOpen, setIsAddMenuOpen] = React.useState(false);
  const [groups, setGroups] = React.useState<IProcessGroup[]>([]);
  const [availableGroups, setAvailableGroups] = React.useState<IProcessGroup[]>([]);
    const [groupCounts, setGroupCounts] = React.useState<Map<number, number>>(new Map());
    const [processCounts, setProcessCounts] = React.useState<Map<number, number>>(new Map());
  const [processes, setProcesses] = React.useState<IProcess[]>([]);
  const [documents, setDocuments] = React.useState<IDocumentoProceso[]>([]);
  const [documentsLoading, setDocumentsLoading] = React.useState(false);
  const [documentsError, setDocumentsError] = React.useState<string>();
  const [folderContents, setFolderContents] = React.useState<Map<number, IDocumentoContenido[]>>(new Map());
  const [loadingDepartments, setLoadingDepartments] = React.useState(true);
  const [loadingGroups, setLoadingGroups] = React.useState(false);
  const [loadingProcesses, setLoadingProcesses] = React.useState(false);
  const [departmentError, setDepartmentError] = React.useState<string>();
  const [groupError, setGroupError] = React.useState<string>();
  const [processError, setProcessError] = React.useState<string>();
  const groupsCache = React.useRef<Map<number, IProcessGroup[]>>(new Map());
  const processesCache = React.useRef<Map<number, IProcess[]>>(new Map());

  React.useEffect(() => {
    const loadDepartments = async (): Promise<void> => {
      try {
        setDepartmentError(undefined);
        const loadedDepartments = await new DepartmentService(sp).getDepartments();
        setDepartments(loadedDepartments);
        new ProcessGroupService(sp).getAllGroups().then(allGroups => {
          const counts = new Map<number, number>();
          allGroups.forEach(group => counts.set(group.departmentId, (counts.get(group.departmentId) || 0) + 1));
          setGroupCounts(counts);
        }, error => console.error('[ProcessView] Error al contar grupos por departamento', error));
        new ProcessService(sp).getProcessCountsByGroup().then(setProcessCounts, error => console.error('[ProcessView] Error al contar procesos por grupo', error));
      } catch (error) {
        console.error('[ProcessView] Error al cargar departamentos', error);
        setDepartmentError('No se han podido cargar los departamentos.');
      } finally {
        setLoadingDepartments(false);
      }
    };
    loadDepartments().catch(() => undefined);
  }, [sp]);

  const loadGroups = async (departmentId: number, force = false): Promise<void> => {
    const cachedGroups = groupsCache.current.get(departmentId);
    if (cachedGroups && !force) {
      setGroups(cachedGroups);
      return;
    }
    try {
      setLoadingGroups(true);
      setGroupError(undefined);
      const loadedGroups = await new ProcessGroupService(sp).getGroupsByDepartment(departmentId);
      groupsCache.current.set(departmentId, loadedGroups);
      setGroups(loadedGroups);
    } catch (error) {
      console.error('[ProcessView] Error al cargar grupos de procesos', error);
      setGroupError('No se han podido cargar los grupos de procesos.');
    } finally {
      setLoadingGroups(false);
    }
  };

  const loadProcesses = async (groupId: number, force = false): Promise<void> => {
    const cachedProcesses = processesCache.current.get(groupId);
    if (cachedProcesses && !force) {
      setProcesses(cachedProcesses);
      return;
    }
    try {
      setLoadingProcesses(true);
      setProcessError(undefined);
      const loadedProcesses = await new ProcessService(sp).getProcessesByGroup(groupId);
      processesCache.current.set(groupId, loadedProcesses);
      setProcesses(loadedProcesses);
    } catch (error) {
      console.error('[ProcessView] Error al cargar procesos', error);
      setProcessError('No se han podido cargar los procesos.');
    } finally {
      setLoadingProcesses(false);
    }
  };

  const selectDepartment = async (departmentId: number): Promise<void> => {
    if (selectedDepartmentId === departmentId) {
      setSelectedDepartmentId(undefined);
      setSelectedGroupId(undefined);
      setSelectedProcess(undefined);
      setGroups([]);
      setProcesses([]);
      return;
    }
    setSelectedDepartmentId(departmentId);
    setSelectedGroupId(undefined);
    setSelectedProcess(undefined);
    setProcesses([]);
    await loadGroups(departmentId);
  };

  const selectGroup = async (groupId: number): Promise<void> => {
    if (selectedGroupId === groupId) {
      setSelectedGroupId(undefined);
      setSelectedProcess(undefined);
      setProcesses([]);
      return;
    }
    setSelectedGroupId(groupId);
    setSelectedProcess(undefined);
    await loadProcesses(groupId);
  };

  const selectProcess = async (process: IProcess): Promise<void> => {
    setSelectedProcess(process);
    setDocuments([]);
    setDocumentsError(undefined);
    setFolderContents(new Map());
    try {
      setDocumentsLoading(true);
      const documentService = new DocumentService(sp);
      const loadedDocuments = await documentService.getDocumentsByProcess(process.id);
      setDocuments(loadedDocuments);
      const folderDocuments = loadedDocuments.filter(document => document.documentType?.toLowerCase() === 'carpeta documental' && document.url);
      const contents = await Promise.all(folderDocuments.map(async folder => [folder.id, await documentService.getFolderContents(folder.url as string)] as [number, IDocumentoContenido[]]));
      setFolderContents(new Map(contents));
    } catch (error) {
      console.error('[ProcessView] Error al cargar documentos del proceso', error);
      setDocumentsError('No se han podido cargar los documentos asociados.');
    } finally {
      setDocumentsLoading(false);
    }
  };

  const closeBranch = (): void => {
    setSelectedDepartmentId(undefined);
    setSelectedGroupId(undefined);
    setSelectedProcess(undefined);
    setGroups([]);
    setProcesses([]);
    setCreateMode(undefined);
    setIsAddMenuOpen(false);
  };

  const compactTree = selectedProcess !== undefined;

  const handleCreated = async (): Promise<void> => {
    if (createMode === 'group' && selectedDepartmentId !== undefined) {
      groupsCache.current.delete(selectedDepartmentId);
      await loadGroups(selectedDepartmentId, true);
    }
    if (createMode === 'process' && selectedGroupId !== undefined) {
      processesCache.current.delete(selectedGroupId);
      await loadProcesses(selectedGroupId, true);
    }
    setCreateMode(undefined);
    setIsAddMenuOpen(false);
  };

  const openProcessForm = async (): Promise<void> => {
    const allGroups = await new ProcessGroupService(sp).getAllGroups();
    setAvailableGroups(allGroups);
    setCreateMode('process');
    setIsAddMenuOpen(false);
  };

  const focusSearchResult = async (result: ISearchResult): Promise<void> => {
    if (result.kind === 'department') {
      setSelectedDepartmentId(result.id);
      setSelectedGroupId(undefined);
      setSelectedProcess(undefined);
      setProcesses([]);
      await loadGroups(result.id);
      return;
    }

    if (result.kind === 'group' && result.departmentId !== undefined) {
      setSelectedDepartmentId(result.departmentId);
      setSelectedGroupId(result.id);
      setSelectedProcess(undefined);
      await loadGroups(result.departmentId);
      await loadProcesses(result.id);
      return;
    }

    if (result.kind === 'process' && result.departmentId !== undefined && result.groupId !== undefined) {
      setSelectedDepartmentId(result.departmentId);
      setSelectedGroupId(result.groupId);
      await loadGroups(result.departmentId);
      const loadedProcesses = await new ProcessService(sp).getProcessesByGroup(result.groupId);
      processesCache.current.set(result.groupId, loadedProcesses);
      setProcesses(loadedProcesses);
      const selected = loadedProcesses.find(process => process.id === result.id);
      if (selected) {
        await selectProcess(selected);
      }
    }
  };

  return (
    <section className={styles.home} aria-label="Catálogo de procesos">
      <header className={styles.header}>
        <div>
          <p className={styles.eyebrow}>Catálogo corporativo</p>
          <h2>Procesos</h2>
          <p className={styles.intro}>Explora los procesos organizados por departamento y grupo.</p>
        </div>
        <GlobalSearch sp={sp} onSelect={result => { focusSearchResult(result).catch(() => undefined); }} />
        <div className={styles.headerActions}>
          <button className={styles.addButton} type="button" onClick={() => setIsAddMenuOpen(!isAddMenuOpen)} aria-expanded={isAddMenuOpen}>Añadir</button>
          {isAddMenuOpen && <div className={styles.addMenu} role="menu">
            <button type="button" role="menuitem" onClick={() => { setCreateMode('group'); setIsAddMenuOpen(false); }}>Grupo de procesos</button>
            <button type="button" role="menuitem" onClick={() => { openProcessForm().catch(() => undefined); }}>Proceso</button>
          </div>}
          {selectedDepartmentId !== undefined && <button className={styles.closeButton} type="button" onClick={closeBranch}>Cerrar rama</button>}
        </div>
      </header>
      <div className={styles.tree}>
        <Department
          departments={compactTree ? departments.filter(department => department.id === selectedDepartmentId) : departments}
          selectedId={selectedDepartmentId}
          loading={loadingDepartments}
          error={departmentError}
          groupCounts={groupCounts}
          onSelect={selectDepartment}
        />
        {selectedDepartmentId !== undefined && <ProcessGroup
          groups={compactTree ? groups.filter(group => group.id === selectedGroupId) : groups}
          selectedId={selectedGroupId}
          loading={loadingGroups}
          error={groupError}
          processCounts={processCounts}
          onSelect={selectGroup}
        />}
        {selectedGroupId !== undefined && <Process
          processes={compactTree ? processes.filter(process => process.id === selectedProcess?.id) : processes}
          selectedId={selectedProcess?.id}
          loading={loadingProcesses}
          error={processError}
          onSelect={process => { selectProcess(process).catch(() => undefined); }}
        />}
      </div>
      {createMode && <CreateItemPanel
        key={createMode}
        mode={createMode}
        sp={sp}
        departments={departments}
        groups={createMode === 'process' ? availableGroups : groups}
        selectedDepartmentId={selectedDepartmentId}
        selectedGroupId={selectedGroupId}
        onClose={() => setCreateMode(undefined)}
        onCreated={handleCreated}
      />}
      {selectedProcess && <ProcessDetail
        process={selectedProcess}
        documents={documents}
        documentsLoading={documentsLoading}
        documentsError={documentsError}
        folderContents={folderContents}
        onClose={() => setSelectedProcess(undefined)}
      />}
    </section>
  );
};

export default Home;