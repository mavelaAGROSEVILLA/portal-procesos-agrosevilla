import * as React from 'react';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import type { IDepartment, IProcessGroup, ISearchResult } from '../../models/ProcessModels';
import { SearchService } from '../../services/SearchService';
import type { SPFI } from '@pnp/sp';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IGlobalSearchProps {
  sp: SPFI;
  departments: IDepartment[];
  groups: IProcessGroup[];
  onSelect: (result: ISearchResult) => void;
}

const resultLabels: Record<ISearchResult['kind'], string> = {
  department: 'Departamento',
  group: 'Grupo',
  process: 'Proceso'
};

const GlobalSearch: React.FunctionComponent<IGlobalSearchProps> = ({ sp, departments, groups, onSelect }) => {
  const [term, setTerm] = React.useState('');
  const [results, setResults] = React.useState<ISearchResult[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState(false);
  const departmentNames = new Map(departments.map(department => [department.id, `${department.code} - ${department.name}`]));
  const groupNames = new Map(groups.map(group => [group.id, `${group.code} - ${group.name}`]));

  React.useEffect(() => {
    const normalizedTerm = term.trim();
    if (normalizedTerm.length < 2) {
      setResults([]);
      setLoading(false);
      return undefined;
    }

    let cancelled = false;
    const timer = window.setTimeout(() => {
      setLoading(true);
      setError(false);
      new SearchService(sp).search(normalizedTerm)
        .then(foundResults => {
          if (!cancelled) {
            setResults(foundResults);
            setLoading(false);
          }
        }, searchError => {
          console.error('[ProcessView] Error en la búsqueda global', searchError);
          if (!cancelled) {
            setResults([]);
            setError(true);
            setLoading(false);
          }
        });
    }, 300);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [sp, term]);

  const selectResult = (result: ISearchResult): void => {
    setTerm('');
    setResults([]);
    onSelect(result);
  };

  return (
    <div className={styles.search}>
      <label className={styles.searchLabel} htmlFor="process-global-search">Buscar en el catálogo</label>
      <div className={styles.searchInputWrap}>
        <Icon iconName="Search" aria-hidden="true" />
        <input
          id="process-global-search"
          type="search"
          value={term}
          placeholder="Código, nombre o descripción"
          onChange={event => setTerm(event.target.value)}
          aria-autocomplete="list"
          aria-controls="process-search-results"
          aria-expanded={term.trim().length >= 2}
        />
        {term && <button className={styles.clearSearch} type="button" onClick={() => { setTerm(''); setResults([]); setError(false); }} aria-label="Limpiar búsqueda"><Icon iconName="Cancel" /></button>}
        {loading && <Spinner size={SpinnerSize.small} ariaLabel="Buscando" />}
      </div>
      {term.trim().length === 1 && <p className={styles.searchHint}>Escribe al menos 2 caracteres.</p>}
      {term.trim().length >= 2 && <div className={styles.searchResults} id="process-search-results" role="listbox" aria-label="Resultados de búsqueda">
        {error && <p className={styles.searchMessage}>No se ha podido buscar en SharePoint. Comprueba tu conexión y vuelve a intentarlo.</p>}
        {!loading && !error && results.length === 0 && <p className={styles.searchMessage}>No hay coincidencias.</p>}
        {results.map(result => <button className={styles.searchResult} key={`${result.kind}-${result.id}`} type="button" role="option" onClick={() => selectResult(result)}>
          <span className={styles.searchResultIcon}><Icon iconName={result.kind === 'process' ? 'Flow' : result.kind === 'group' ? 'BulletedList' : 'Org'} /></span>
          <span className={styles.searchResultCopy}><strong>{result.code}</strong><span>{result.name}</span>
            {(result.departmentId !== undefined || result.groupId !== undefined) && <small>{[
              result.departmentId !== undefined ? departmentNames.get(result.departmentId) : undefined,
              result.groupId !== undefined ? groupNames.get(result.groupId) : undefined
            ].filter(Boolean).join(' › ')}</small>}
          </span>
          <span className={styles.searchResultKind}>{resultLabels[result.kind]}</span>
        </button>)}
      </div>}
    </div>
  );
};

export default GlobalSearch;