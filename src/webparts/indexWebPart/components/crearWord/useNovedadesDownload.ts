import { useState, useCallback } from 'react';
import { BuildDocInput } from '../../../../core/utilsWord/types';
import { DOCX_FILENAME } from '../../../../core/utils/Constants';
import createNovedadesDoc from '../../../../core/utilsWord/buildDoc';
import { downloadDoc } from '../../../../core/utilsWord/downloadDoc';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphListService } from '../datosSharepoint/listaConGraph';


type UseNovedadesArgs = {
  context: WebPartContext;
  filename?: string;
  listName?: string; // default: ListaPruebaGraph
  /** Si pasás esto, se usa tal cual y NO se consultará Graph */
  inputOverride?: Omit<BuildDocInput, 'maxItemsPerSection'>;
};

export function useNovedades({
  context,
  filename = DOCX_FILENAME,
  listName = 'Word',
  inputOverride,
}: UseNovedadesArgs) {
  const [loading, setLoading] = useState(false);

  const handleDownload = useCallback(async () => {
    try {
      setLoading(true);

      let input: BuildDocInput;

      if (inputOverride) {
        input = { ...inputOverride, confidentialityLabel: inputOverride.confidentialityLabel ?? 'YPF-Confidencial' };
      } else {
        const svc = new GraphListService(context);
        await svc.init();
        const data = await svc.fetchNovedades(listName);

        input = {
          sectorGeneral: data.area ?? '',
          novedad: data.novedad ?? [],
          confidentialityLabel: 'YPF-Confidencial',
        };
      }

      const doc = await createNovedadesDoc(input);
      await downloadDoc(doc, filename);
    } finally {
      setLoading(false);
    }
  }, [context, filename, inputOverride, listName]);

  return { handleDownload, loading };
}
