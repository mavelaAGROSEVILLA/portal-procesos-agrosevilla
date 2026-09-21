import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import type { IDocumentoProceso, IProcess } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IProcessDetailProps {
  process: IProcess;
  documents: IDocumentoProceso[];
  documentsLoading: boolean;
  documentsError?: string;
  onClose: () => void;
}

const ProcessDetail: React.FunctionComponent<IProcessDetailProps> = ({
  process,
  documents,
  documentsLoading,
  documentsError,
  onClose
}) => (
  <section className={styles.detail} aria-labelledby="process-detail-heading">
    <div className={styles.detailHeader}>
      <div>
        <p className={styles.eyebrow}>Ficha del proceso</p>
        <h3 id="process-detail-heading">{process.name}</h3>
      </div>
      <button className={styles.closeButton} type="button" onClick={onClose}>Cerrar detalle</button>
    </div>
    <div className={styles.detailGrid}>
      <div><span>Código</span><strong>{process.code || 'Sin código'}</strong></div>
      <div><span>Responsable</span><strong>{process.owner || 'No informado'}</strong></div>
      <div><span>Estado</span><strong>{process.status || 'No informado'}</strong></div>
      <div><span>Tipo</span><strong>{process.type || 'No informado'}</strong></div>
    </div>
    <div className={styles.detailSection}>
      <h4>Descripción</h4>
      <p>{process.description || 'Este proceso no tiene una descripción registrada.'}</p>
    </div>
    <div className={styles.detailSection}>
      <h4>Documentos asociados</h4>
      {documentsLoading && <Spinner size={SpinnerSize.small} label="Cargando documentos" />}
      {documentsError && <p className={styles.message} role="alert">{documentsError}</p>}
      {!documentsLoading && !documentsError && documents.length === 0 && (
        <p className={styles.message}>Todavía no hay documentos asociados a este proceso.</p>
      )}
      {!documentsLoading && !documentsError && documents.length > 0 && (
        <ul className={styles.documentList}>
          {documents.map(document => <li key={document.id}>
            {document.url ? <a href={document.url} target="_blank" rel="noreferrer">{document.name}</a> : document.name}
          </li>)}
        </ul>
      )}
    </div>
  </section>
);

export default ProcessDetail;