import { useState, useCallback } from 'react';
import { BuildDocInput } from '../../../../core/utilsWord/types';
import { DOCX_FILENAME } from '../../../../core/utils/Constants';
import { createNovedadesDoc } from '../../../../core/utilsWord/buildDoc';
import { downloadDoc } from '../../../../core/utilsWord/downloadDoc';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphListService } from '../datosSharepoint/listaConGraph';
import { NovedadItem } from '../datosSharepoint/listaConGraph';

type UseNovedadesArgs = {
    context: WebPartContext;
    filename?: string;
    listName?: string;
    inputOverride?: Omit<BuildDocInput, 'maxItemsPerSection'>;
};

type UseNovedadesResult = {
    handleDownload: () => Promise<void>;
    loading: boolean;
};

export function useNovedades({
    context,
    filename = DOCX_FILENAME,
    listName = 'Word',
    inputOverride,
}: UseNovedadesArgs): UseNovedadesResult {
    const [loading, setLoading] = useState(false);

    const handleDownload = useCallback(async () => {
        try {
            setLoading(true);

            let input: BuildDocInput;

            if (inputOverride) {
                input = {
                    ...inputOverride,
                    confidentialityLabel:
                        inputOverride.confidentialityLabel ??
                        'YPF-Confidencial',
                };
            } else {
                const svc = new GraphListService(context);
                await svc.init();
                const data = await svc.fetchNovedades(listName);

                const novedades: BuildDocInput['novedades'] = [];

                data.novedad.forEach((entry) => {
                    const { areaNovedad, items } = entry;

                    items.forEach(
                        (item: NovedadItem & { sectorGeneral: string }) => {
                            novedades.push({
                                sectorGeneral: item.sectorGeneral, // ✅ Usa el valor real
                                areaNovedad,
                                tituloNovedad: item.tituloNovedad,
                                detalleNovedad: item.detalleNovedad,
                                imagenesNovedad: item.imagenesNovedad,
                            });
                        }
                    );
                });

                input = {
                    novedades,
                    confidentialityLabel: 'YPF-Confidencial',
                };

                // // DEBUG si querés ver el input:
                // console.log('INPUT al doc:', JSON.stringify(input, null, 2));
            }

            const doc = await createNovedadesDoc(input);
            await downloadDoc(doc, filename);
        } finally {
            setLoading(false);
        }
    }, [context, filename, inputOverride, listName]);

    return { handleDownload, loading };
}
