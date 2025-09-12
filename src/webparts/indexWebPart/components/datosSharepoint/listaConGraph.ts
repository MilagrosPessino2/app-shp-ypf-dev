import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';

export type ListaPruebaFields = {
    SectorGeneral?: string;
    AreaNovedad?: string;
    TituloNovedad?: string;
    DetalleNovedad?: string;
    Resumen?: string;
    Title?: string;
};

export type NovedadItem = {
    sectorGeneral: string;
    tituloNovedad: string;
    resumenHtml: string;
    detalleNovedad: string;
    imagenesNovedad: string[];
};

export type NovedadesData = {
    area: string;
    novedad: Array<{ areaNovedad: string; items: NovedadItem[] }>;
};

export class GraphListService {
    private client!: MSGraphClientV3;
    private siteId!: string;

    constructor(private context: WebPartContext) {}

    public async init(): Promise<void> {
        this.client = await this.context.msGraphClientFactory.getClient('3');
        this.siteId = this.buildGraphSiteId();
    }

    public async fetchNovedades(
        listDisplayName: string
    ): Promise<NovedadesData> {
        const list = await this.getList(listDisplayName);
        const items = await this.getItems(list);

        const groups = new Map<string, NovedadItem[]>();

        for (const it of items) {
            const f: ListaPruebaFields = it.fields || {};

            const areaNovedad = (f.AreaNovedad ?? 'General').trim();
            const titulo = (
                f.TituloNovedad ??
                f.Title ??
                '(Sin título)'
            ).trim();
            const detalleHtml = f.DetalleNovedad ?? '';

            const urls: string[] = [];

            // 1) Intento con Graph
            if (list?.id) {
                try {
                    const resp: {
                        value?: Array<{
                            file?: { mimeType?: string };
                            ['@microsoft.graph.downloadUrl']?: string;
                        }>;
                    } = await this.client
                        .api(
                            `/sites/${this.siteId}/lists/${list.id}/items/${it.id}/driveItem/children`
                        )
                        .select('id,name,@microsoft.graph.downloadUrl,file')
                        .get();

                    const files = resp?.value ?? [];
                    for (const ch of files) {
                        const mime = ch?.file?.mimeType || '';
                        const dl = ch?.['@microsoft.graph.downloadUrl'];
                        if (dl && /^image\//i.test(mime)) urls.push(dl);
                    }
                } catch {
                    // ignorar
                }
            }

            // 2) Fallback REST
            if (urls.length === 0) {
                try {
                    const listTitleEncoded =
                        encodeURIComponent(listDisplayName);
                    const apiUrl =
                        `${this.context.pageContext.web.absoluteUrl}` +
                        `/_api/web/lists/getByTitle('${listTitleEncoded}')/items(${it.id})/AttachmentFiles` +
                        `?$select=FileName,ServerRelativeUrl`;

                    const res = await this.context.spHttpClient.get(
                        apiUrl,
                        SPHttpClient.configurations.v1
                    );
                    if (res.ok) {
                        const json: {
                            value: Array<{
                                FileName: string;
                                ServerRelativeUrl: string;
                            }>;
                        } = await res.json();

                        for (const af of json?.value ?? []) {
                            const fileName = af.FileName || '';
                            const rel = af.ServerRelativeUrl || '';
                            if (
                                rel &&
                                /\.(png|jpe?g|gif|bmp|webp)$/i.test(fileName)
                            ) {
                                const abs = new URL(
                                    rel,
                                    this.context.pageContext.web.absoluteUrl
                                ).toString();
                                urls.push(abs);
                            }
                        }
                    }
                } catch {
                    // ignorar
                }
            }

            const arr = groups.get(areaNovedad) ?? [];
            arr.push({
                tituloNovedad: titulo,
                detalleNovedad: detalleHtml,
                resumenHtml: f.Resumen ?? '',
                imagenesNovedad: Array.from(new Set(urls)),
                sectorGeneral: (f.SectorGeneral ?? f.Title ?? 'General').trim(),
            });
            groups.set(areaNovedad, arr);
        }

        const sectorGeneral = this.pickMostFrequent(
            items.map((it) =>
                (
                    it.fields?.SectorGeneral ??
                    it.fields?.Title ??
                    'General'
                ).trim()
            )
        );

        const novedad = Array.from(groups.entries())
            .map(([areaNovedad, items]) => ({ areaNovedad, items }))
            .sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

        return { area: sectorGeneral, novedad };
    }

    private buildGraphSiteId(): string {
        const url = new URL(this.context.pageContext.web.absoluteUrl);
        const hostname = url.hostname;
        const siteId = this.context.pageContext.site.id.toString();
        const webId = this.context.pageContext.web.id.toString();
        return `${hostname},${siteId},${webId}`;
    }

    private async getList(
        displayName: string
    ): Promise<{ id: string; name: string } | null> {
        const response: {
            value?: Array<{ id: string; name?: string; displayName?: string }>;
        } = await this.client
            .api(`/sites/${this.siteId}/lists`)
            .select('id,name,displayName')
            .get();

        const lists = response?.value ?? [];

        const match = lists.find(
            (l) =>
                (l.displayName || '').toLowerCase() ===
                    displayName.toLowerCase() ||
                (l.name || '').toLowerCase() === displayName.toLowerCase()
        );

        return match
            ? { id: match.id, name: match.displayName || match.name || '' }
            : null;
    }

    private async getItems(
        list: { id: string } | null
    ): Promise<Array<{ id: string; fields: ListaPruebaFields }>> {
        if (!list?.id) return [];

        const response: {
            value?: Array<{ id: string; fields: ListaPruebaFields }>;
        } = await this.client
            .api(`/sites/${this.siteId}/lists/${list.id}/items`)
            .expand(
                'fields($select=Title,SectorGeneral,AreaNovedad,TituloNovedad,DetalleNovedad,Resumen)'
            )
            .top(200)
            .get();

        return response?.value ?? [];
    }

    private pickMostFrequent(values: string[]): string {
        const counts: Record<string, number> = {};
        for (const v of values) counts[v] = (counts[v] || 0) + 1;

        let best = 'General';
        let max = -1;

        for (const [k, v] of Object.entries(counts)) {
            if (v > max) {
                max = v;
                best = k;
            }
        }

        return best;
    }
}
