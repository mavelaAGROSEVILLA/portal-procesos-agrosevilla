import * as React from 'react';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import type { IDepartment } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IDepartmentProps {
  departments: IDepartment[];
  selectedId?: number;
  loading: boolean;
  error?: string;
  groupCounts: Map<number, number>;
  onSelect: (id: number) => Promise<void>;
}

const Department: React.FunctionComponent<IDepartmentProps> = ({ departments, selectedId, loading, error, groupCounts, onSelect }) => (
  <section className={styles.level} aria-labelledby="departments-heading">
    <div className={styles.levelHeading}><span className={styles.levelNumber}>01</span><h3 id="departments-heading">Departamentos</h3></div>
    {loading && <Spinner size={SpinnerSize.small} label="Cargando departamentos" />}
    {error && <p className={styles.message} role="alert">{error}</p>}
    {!loading && !error && departments.length === 0 && <p className={styles.message}>No hay departamentos disponibles.</p>}
    <div className={styles.items}>
      {departments.map(department => <button
        className={`${styles.item} ${selectedId === department.id ? styles.selected : ''}`}
        type="button"
        key={department.id}
        aria-expanded={selectedId === department.id}
        onClick={() => { onSelect(department.id).catch(() => undefined); }}
      >
        <Icon iconName={selectedId === department.id ? 'ChevronDown' : 'ChevronRight'} />
        <span><strong>{department.code}</strong><span>{department.name}</span><small className={styles.itemCount}>{groupCounts.get(department.id) || 0} grupos</small></span>
      </button>)}
    </div>
  </section>
);

export default Department;