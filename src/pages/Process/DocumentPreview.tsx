import * as React from 'react';
import type { IDocumentoProceso } from '../../models/ProcessModels';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IDocumentPreviewProps {
  document?: IDocumentoProceso;
  kind: 'master' | 'bpmn';
}

const isPdf = (document: IDocumentoProceso): boolean => /\.pdf(?:$|[?#])/i.test(document.name) || /\.pdf(?:$|[?#])/i.test(document.url || '');
const isPng = (document: IDocumentoProceso): boolean => /\.png(?:$|[?#])/i.test(document.name) || /\.png(?:$|[?#])/i.test(document.url || '');

const DocumentPreview: React.FunctionComponent<IDocumentPreviewProps> = ({ document, kind }) => {
  const [scale, setScale] = React.useState(1);
  const [position, setPosition] = React.useState({ x: 0, y: 0 });
  const [dragging, setDragging] = React.useState(false);
  const dragStart = React.useRef({ x: 0, y: 0 });
  const positionStart = React.useRef({ x: 0, y: 0 });

  React.useEffect(() => {
    setScale(1);
    setPosition({ x: 0, y: 0 });
  }, [document?.id]);

  if (!document || !document.url) {
    return <p className={styles.message}>No hay un documento de tipo {kind === 'master' ? 'Ficha Maestra' : 'BPMN'} asociado.</p>;
  }

  if (kind === 'master' && !isPdf(document)) {
    return <p className={styles.message}>La Ficha Maestra asociada no es un PDF.</p>;
  }

  if (kind === 'bpmn' && !isPng(document)) {
    return <p className={styles.message}>El documento BPMN asociado no es un PNG.</p>;
  }

  if (kind === 'master') {
    return <div className={styles.pdfPreview}>
      <iframe className={styles.pdfFrame} src={document.url} title={`Vista previa de ${document.name}`} />
      <a className={styles.previewLink} href={document.url} target="_blank" rel="noreferrer">Abrir PDF en una pestaña nueva</a>
    </div>;
  }

  const handleWheel = (event: React.WheelEvent<HTMLDivElement>): void => {
    event.preventDefault();
    setScale(current => Math.min(10, Math.max(0.5, current - event.deltaY * 0.001)));
  };

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>): void => {
    event.currentTarget.setPointerCapture(event.pointerId);
    dragStart.current = { x: event.clientX, y: event.clientY };
    positionStart.current = position;
    setDragging(true);
  };

  const handlePointerMove = (event: React.PointerEvent<HTMLDivElement>): void => {
    if (!dragging) {
      return;
    }
    setPosition({
      x: positionStart.current.x + event.clientX - dragStart.current.x,
      y: positionStart.current.y + event.clientY - dragStart.current.y
    });
  };

  const reset = (): void => {
    setScale(1);
    setPosition({ x: 0, y: 0 });
  };

  return <div className={styles.imagePreview}>
    <div className={styles.previewToolbar}>
      <span>Zoom {Math.round(scale * 100)}%</span>
      <button type="button" onClick={() => setScale(current => Math.min(10, current + 0.25))}>+</button>
      <button type="button" onClick={() => setScale(current => Math.max(0.5, current - 0.25))}>-</button>
      <button type="button" onClick={reset}>Restablecer</button>
      <a href={document.url} target="_blank" rel="noreferrer">Abrir imagen</a>
    </div>
    <div
      className={`${styles.imageViewport} ${dragging ? styles.dragging : ''}`}
      onWheel={handleWheel}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={() => setDragging(false)}
      onPointerCancel={() => setDragging(false)}
      role="application"
      aria-label="Vista previa BPMN; usa la rueda para zoom y arrastra para desplazarte"
    >
      <img
        className={styles.previewImage}
        src={document.url}
        alt={`Diagrama BPMN ${document.name}`}
        style={{ transform: `translate(${position.x}px, ${position.y}px) scale(${scale})` }}
        draggable={false}
      />
    </div>
  </div>;
};

export default DocumentPreview;