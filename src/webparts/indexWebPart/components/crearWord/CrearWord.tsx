import * as React from 'react';
import styles from './CrearWord.module.scss';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useNovedades } from './useNovedadesDownload';

type Props = { context: WebPartContext };

const CrearWord: React.FC<Props> = ({ context }) => {
  const { handleDownload, loading } = useNovedades({
    context,
    listName: 'ListaPruebaGraph',
  });

  return (
    <button className={styles.button} onClick={handleDownload} disabled={loading}>
      {loading ? 'Generando...' : 'Descargar Word'}
    </button>
  );
};

export default CrearWord;
