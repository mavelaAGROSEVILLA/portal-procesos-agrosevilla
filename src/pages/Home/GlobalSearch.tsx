import * as React from 'react';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import type { ISearchResult } from '../../models/ProcessModels';
import { SearchService } from '../../services/SearchService';
import type { SPFI } from '@pnp/sp';
import styles from '../../webparts/processView/components/Home.module.scss';

interface IGlobalSearchProps {
  sp: SPFI;
  onSelect: (result: ISearchResult) => void;
}

const resultLabels: Record<ISearchResult['kind'], string> = {
  department: 'Departamento',
  group: 'Grupo',
  process: 'Proceso'
};

const GlobalSearch: React.FunctionComponent<IGlobalSearchProps> = ({ sp, onSelect }) => {
  const [term, setTerm] = React.useState('');
  const [results, setResults] = React.useState<ISearchResult[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState(false);

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
        />
        {loading && <Spinner size={SpinnerSize.small} ariaLabel="Buscando" />}
      </div>
      {term.trim().length >= 2 && <div className={styles.searchResults} id="process-search-results" role="listbox">
        {error && <p className={styles.searchMessage}>No se ha podido realizar la búsqueda.</p>}
        {!loading && !error && results.length === 0 && <p className={styles.searchMessage}>No hay coincidencias.</p>}
        {results.map(result => <button className={styles.searchResult} key={`${result.kind}-${result.id}`} type="button" role="option" onClick={() => selectResult(result)}>
          <span className={styles.searchResultIcon}><Icon iconName={result.kind === 'process' ? 'Flow' : result.kind === 'group' ? 'BulletedList' : 'Org'} /></span>
          <span className={styles.searchResultCopy}><strong>{result.code}</strong><span>{result.name}</span></span>
          <span className={styles.searchResultKind}>{resultLabels[result.kind]}</span>
        </button>)}
      </div>}
    </div>
  );
};

export default GlobalSearch;