import * as React from 'react';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import type { IProcessGroup } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IProcessGroupProps {
  groups: IProcessGroup[];
  selectedId?: number;
  loading: boolean;
  error?: string;
  onSelect: (id: number) => Promise<void>;
}

const ProcessGroup: React.FunctionComponent<IProcessGroupProps> = ({ groups, selectedId, loading, error, onSelect }) => (
  <section className={`${styles.level} ${styles.childLevel}`} aria-labelledby="groups-heading">
    <div className={styles.levelHeading}><span className={styles.levelNumber}>02</span><h3 id="groups-heading">Grupos de procesos</h3></div>
    {loading && <Spinner size={SpinnerSize.small} label="Cargando grupos de procesos" />}
    {error && <p className={styles.message} role="alert">{error}</p>}
    {!loading && !error && groups.length === 0 && <p className={styles.message}>Este departamento no tiene grupos.</p>}
    <div className={styles.items}>
      {groups.map(group => <button
        className={`${styles.item} ${selectedId === group.id ? styles.selected : ''}`}
        type="button"
        key={group.id}
        aria-expanded={selectedId === group.id}
        onClick={() => { onSelect(group.id).catch(() => undefined); }}
      >
        <Icon iconName={selectedId === group.id ? 'ChevronDown' : 'ChevronRight'} />
        <span><strong>{group.code}</strong><span>{group.name}</span></span>
      </button>)}
    </div>
  </section>
);

export default ProcessGroup;