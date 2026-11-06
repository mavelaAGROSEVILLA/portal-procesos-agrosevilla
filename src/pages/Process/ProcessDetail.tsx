import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import type { IDocumentoContenido, IDocumentoProceso, IProcess } from '../../models/ProcessModels';
import DocumentPreview from './DocumentPreview';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IProcessDetailProps {
  process: IProcess;
  documents: IDocumentoProceso[];
  documentsLoading: boolean;
  documentsError?: string;
  folderContents: Map<number, IDocumentoContenido[]>;
  departmentName?: string;
  groupName?: string;
  returnLabel?: string;
  onClose: () => void;
}

const ProcessDetail: React.FunctionComponent<IProcessDetailProps> = ({
  process,
  documents,
  documentsLoading,
  documentsError,
  folderContents,
  departmentName,
  groupName,
  returnLabel = 'Cerrar detalle',
  onClose
}) => (
  <section className={styles.detail} aria-labelledby="process-detail-heading">
    <div className={styles.detailHeader}>
      <div>
        <p className={styles.eyebrow}>Ficha del proceso</p>
        <h3 id="process-detail-heading">{process.name}</h3>
      </div>
      <button className={styles.closeButton} type="button" onClick={onClose}>{returnLabel}</button>
    </div>
    <div className={styles.detailGrid}>
      <div><span>Código</span><strong>{process.code || 'Sin código'}</strong></div>
      <div><span>Departamento</span><strong>{departmentName || 'No informado'}</strong></div>
      <div><span>Grupo de proceso</span><strong>{groupName || 'Sin grupo'}</strong></div>
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
    <div className={styles.previewGrid}>
      <div className={styles.detailSection}>
        <h4>Ficha Maestra</h4>
        <DocumentPreview document={documents.find(document => document.documentType?.toLowerCase() === 'ficha maestra')} kind="master" />
      </div>
      <div className={styles.detailSection}>
        <h4>Diagrama BPMN</h4>
        <DocumentPreview document={documents.find(document => document.documentType?.toLowerCase() === 'bpmn')} kind="bpmn" />
      </div>
    </div>
    {documents.filter(document => document.documentType?.toLowerCase() === 'carpeta documental').map(folder => (
      <div className={styles.folderSection} key={folder.id}>
        <h4>Contenido de la carpeta documental</h4>
        {folder.url && <a className={styles.folderLink} href={folder.url} target="_blank" rel="noreferrer">{folder.name}</a>}
        {(folderContents.get(folder.id) || []).length === 0 && <p className={styles.message}>La carpeta no contiene elementos visibles.</p>}
        {(folderContents.get(folder.id) || []).length > 0 && <ul className={styles.documentList}>
          {(folderContents.get(folder.id) || []).map(content => <li key={content.url}>
            <a href={content.url} target="_blank" rel="noreferrer">{content.isFolder ? `Carpeta: ${content.name}` : content.name}</a>
          </li>)}
        </ul>}
      </div>
    ))}
  </section>
);

export default ProcessDetail;