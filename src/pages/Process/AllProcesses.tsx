import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import type { IDepartment, IProcess, IProcessGroup } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IAllProcessesProps {
  processes: IProcess[];
  departments: IDepartment[];
  groups: IProcessGroup[];
  loading: boolean;
  error?: string;
  onSelect: (process: IProcess) => void;
}

const AllProcesses: React.FunctionComponent<IAllProcessesProps> = ({ processes, departments, groups, loading, error, onSelect }) => {
  const departmentNames = new Map(departments.map(department => [department.id, `${department.code} - ${department.name}`]));
  const groupNames = new Map(groups.map(group => [group.id, `${group.code} - ${group.name}`]));

  return (
    <section className={`${styles.level} ${styles.allProcesses}`} aria-labelledby="all-processes-heading">
      <div className={styles.levelHeading}><span className={styles.levelNumber}>TODOS</span><h3 id="all-processes-heading">Todos los procesos</h3></div>
      {loading && <Spinner size={SpinnerSize.small} label="Cargando todos los procesos" />}
      {error && <p className={styles.message} role="alert">{error}</p>}
      {!loading && !error && processes.length === 0 && <p className={styles.message}>No hay procesos disponibles.</p>}
      {!loading && !error && processes.length > 0 && <div className={styles.allProcessItems}>
        {processes.map(process => <article className={styles.allProcessItem} key={process.id}>
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
    </section>
  );
};

export default AllProcesses;