import * as React from 'react';
import { DefaultButton, Dropdown, IDropdownOption, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import type { IDepartment, IProcess, IProcessGroup } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IAllProcessesProps {
  processes: IProcess[];
  selectedProcessId?: number;
  departments: IDepartment[];
  groups: IProcessGroup[];
  loading: boolean;
  error?: string;
  onSelect: (process: IProcess) => void;
}

const AllProcesses: React.FunctionComponent<IAllProcessesProps> = ({ processes, selectedProcessId, departments, groups, loading, error, onSelect }) => {
  const departmentNames = new Map(departments.map(department => [department.id, `${department.code} - ${department.name}`]));
  const groupNames = new Map(groups.map(group => [group.id, `${group.code} - ${group.name}`]));
  const [term, setTerm] = React.useState('');
  const [status, setStatus] = React.useState('');
  const [sortBy, setSortBy] = React.useState<'code' | 'name'>('code');
  const [visibleCount, setVisibleCount] = React.useState(30);

  const statuses = Array.from(new Set(processes.map(process => process.status?.trim()).filter((value): value is string => Boolean(value))));
  const statusOptions: IDropdownOption[] = [
    { key: '', text: 'Todos los estados' },
    ...statuses.map(value => ({ key: value, text: value }))
  ];
  const filteredProcesses = processes
    .filter(process => !status || process.status === status)
    .filter(process => {
      const normalizedTerm = term.trim().toLocaleLowerCase();
      return !normalizedTerm || `${process.code} ${process.name} ${process.description}`.toLocaleLowerCase().includes(normalizedTerm);
    })
    .sort((left, right) => `${sortBy === 'code' ? left.code : left.name} ${left.code}`.localeCompare(`${sortBy === 'code' ? right.code : right.name} ${right.code}`));
  const visibleProcesses = filteredProcesses.slice(0, visibleCount);

  React.useEffect(() => {
    setVisibleCount(30);
  }, [term, status, sortBy]);

  const clearFilters = (): void => {
    setTerm('');
    setStatus('');
    setSortBy('code');
  };

  return (
    <section className={`${styles.level} ${styles.allProcesses}`} aria-labelledby="all-processes-heading">
      <div className={styles.levelHeading}><span className={styles.levelNumber}>TODOS</span><h3 id="all-processes-heading">Todos los procesos</h3></div>
      {loading && <Spinner size={SpinnerSize.small} label="Cargando todos los procesos" />}
      {error && <p className={styles.message} role="alert">{error}</p>}
      {!loading && !error && processes.length === 0 && <p className={styles.message}>No hay procesos disponibles.</p>}
      {!loading && !error && processes.length > 0 && <>
        <div className={styles.filterPanel} aria-label="Filtros de procesos">
          <TextField label="Buscar" placeholder="Código, nombre o descripción" value={term} onChange={(_, value) => setTerm(value || '')} />
          <Dropdown label="Estado" options={statusOptions} selectedKey={status} onChange={(_, option) => setStatus(typeof option?.key === 'string' ? option.key : '')} />
          <Dropdown label="Ordenar por" options={[
            { key: 'code', text: 'Código' },
            { key: 'name', text: 'Nombre' }
          ]} selectedKey={sortBy} onChange={(_, option) => setSortBy(option?.key === 'name' ? 'name' : 'code')} />
          <DefaultButton text="Limpiar filtros" onClick={clearFilters} />
        </div>
        <p className={styles.resultCount} aria-live="polite">{filteredProcesses.length} {filteredProcesses.length === 1 ? 'proceso' : 'procesos'} encontrados</p>
        {filteredProcesses.length === 0
          ? <p className={styles.message}>No hay procesos que coincidan con los filtros. Prueba a cambiar o limpiar los filtros.</p>
          : <div className={styles.allProcessItems}>
        {visibleProcesses.map(process => <article className={`${styles.allProcessItem} ${selectedProcessId === process.id ? styles.selectedProcess : ''}`} key={process.id}>
          <button className={styles.processSelect} type="button" onClick={() => onSelect(process)}>
            <span className={styles.processCopy}><strong>{process.code}</strong><h4>{process.name}</h4></span>
            <span className={styles.processAction}>Ver ficha</span>
          </button>
          <div className={styles.allProcessData}>
            <span><strong>Departamento</strong>{departmentNames.get(process.departmentId || 0) || 'Sin departamento'}</span>
            <span><strong>Grupo de proceso</strong>{groupNames.get(process.groupId) || 'Sin grupo'}</span>
          </div>
        </article>)}
        </div>}
        {visibleCount < filteredProcesses.length && <div className={styles.loadMore}><DefaultButton text={`Mostrar ${Math.min(30, filteredProcesses.length - visibleCount)} más`} onClick={() => setVisibleCount(count => count + 30)} /></div>}
      </>}
    </section>
  );
};

export default AllProcesses;