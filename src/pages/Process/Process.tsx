import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import type { IProcess } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IProcessProps {
  processes: IProcess[];
  selectedId?: number;
  loading: boolean;
  error?: string;
  onSelect: (process: IProcess) => void;
}

const Process: React.FunctionComponent<IProcessProps> = ({ processes, selectedId, loading, error, onSelect }) => (
  <section className={`${styles.level} ${styles.processLevel}`} aria-labelledby="processes-heading">
    <div className={styles.levelHeading}><span className={styles.levelNumber}>03</span><h3 id="processes-heading">Procesos</h3></div>
    {loading && <Spinner size={SpinnerSize.small} label="Cargando procesos" />}
    {error && <p className={styles.message} role="alert">{error}</p>}
    {!loading && !error && processes.length === 0 && <p className={styles.message}>Este grupo no tiene procesos.</p>}
    <div className={styles.processItems}>
      {processes.map(process => <article className={`${styles.processItem} ${selectedId === process.id ? styles.selectedProcess : ''}`} key={process.id}>
        <button className={styles.processSelect} type="button" onClick={() => onSelect(process)} aria-pressed={selectedId === process.id}>
          <span className={styles.processCopy}><strong>{process.code}</strong><h4>{process.name}</h4></span>
          <span className={styles.processAction}>Ver ficha</span>
        </button>
        {process.description && <p>{process.description}</p>}
        {(process.status || process.type) && <div className={styles.meta}>{process.status || process.type}</div>}
      </article>)}
    </div>
  </section>
);

export default Process;