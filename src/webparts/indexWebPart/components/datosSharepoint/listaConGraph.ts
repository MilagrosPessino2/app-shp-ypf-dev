import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';

/** Columnas (internal names = display names) según tu captura */
export type ListaPruebaFields = {
  SectorGeneral?: string;    // Texto
  AreaNovedad?: string;      // Texto
  TituloNovedad?: string;    // Texto
  DetalleNovedad?: string;   // Texto enriquecido (HTML)
  Resumen?: string;          // (opcional)
  Title?: string;            // fallback por si renombraron Title
};

export type NovedadItem = {
  tituloNovedad: string;
  detalleNovedad: string;     // HTML
  imagenesNovedad: string[];  // URLs directas a imágenes
};

export type NovedadesData = {
  area: string; // SectorGeneral más frecuente
  novedad: Array<{ areaNovedad: string; items: NovedadItem[] }>;
};

export class GraphListService {
  private client!: MSGraphClientV3;
  private siteId!: string;

  constructor(private context: WebPartContext) {}

  public async init() {
    this.client = await this.context.msGraphClientFactory.getClient("3");
    this.siteId = this.buildGraphSiteId();
  }

  /** Trae y mapea: agrupa por AreaNovedad y define SectorGeneral (más frecuente) */
  public async fetchNovedades(listDisplayName: string): Promise<NovedadesData> {
    const list  = await this.getList(listDisplayName);
    const items = await this.getItems(list);

    const groups = new Map<string, NovedadItem[]>();

    for (const it of items) {
      const f = (it.fields || {}) as ListaPruebaFields;

      const areaNovedad = (f.AreaNovedad ?? 'General').trim();
      const titulo      = (f.TituloNovedad ?? f.Title ?? '(Sin título)').trim();
      const detalleHtml = f.DetalleNovedad ?? '';

      const urls: string[] = [];

      // 1) Intento con Graph (driveItem/children) y filtro image/*
      if (list?.id) {
        try {
          const resp = await this.client
            .api(`/sites/${this.siteId}/lists/${list.id}/items/${it.id}/driveItem/children`)
            .select('id,name,@microsoft.graph.downloadUrl,file')
            .get();

          const files = resp?.value ?? [];
          for (const ch of files) {
            const mime = ch?.file?.mimeType || '';
            const dl = ch?.['@microsoft.graph.downloadUrl'];
            if (dl && /^image\//i.test(mime)) urls.push(dl);
          }
        } catch { /* ignorar */ }
      }

      // 2) Fallback REST (AttachmentFiles) → URLs de mismo origen (evita CORS)
      if (urls.length === 0) {
        try {
          const listTitleEncoded = encodeURIComponent(listDisplayName);
          const apiUrl =
            `${this.context.pageContext.web.absoluteUrl}` +
            `/_api/web/lists/getByTitle('${listTitleEncoded}')/items(${it.id})/AttachmentFiles` +
            `?$select=FileName,ServerRelativeUrl`;

          const res = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
          if (res.ok) {
            const json: any = await res.json();
            for (const af of (json?.value ?? [])) {
              const fileName: string = af.FileName || '';
              const rel: string = af.ServerRelativeUrl || '';
              if (rel && /\.(png|jpe?g|gif|bmp|webp)$/i.test(fileName)) {
                const abs = new URL(rel, this.context.pageContext.web.absoluteUrl).toString();
                urls.push(abs);
              }
            }
          }
        } catch { /* ignorar */ }
      }

      const arr = groups.get(areaNovedad) ?? [];
      arr.push({
        tituloNovedad: titulo,
        detalleNovedad: detalleHtml,
        imagenesNovedad: Array.from(new Set(urls)),
      });
      groups.set(areaNovedad, arr);
    }

    // SectorGeneral = valor más frecuente (SectorGeneral || Title)
    const sectorGeneral = this.pickMostFrequent(
      items.map((it) => ((it.fields?.SectorGeneral ?? it.fields?.Title) ?? 'General').trim())
    );

    const novedad = Array.from(groups.entries())
      .map(([areaNovedad, items]) => ({ areaNovedad, items }))
      .sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

    return { area: sectorGeneral, novedad };
  }

  /* -------------------- helpers -------------------- */
  private buildGraphSiteId(): string {
    const url = new URL(this.context.pageContext.web.absoluteUrl);
    const hostname = url.hostname;
    const siteId = this.context.pageContext.site.id.toString();
    const webId  = this.context.pageContext.web.id.toString();
    return `${hostname},${siteId},${webId}`;
  }

  private async getList(displayName: string): Promise<{ id: string; name: string } | null> {
    const lists = await this.client.api(`/sites/${this.siteId}/lists`).select('id,name,displayName').get();
    const match = (lists?.value ?? []).find((l: any) =>
      (l.displayName || '').toLowerCase() === displayName.toLowerCase() ||
      (l.name || '').toLowerCase() === displayName.toLowerCase()
    );
    return match ? { id: match.id, name: match.displayName || match.name } : null;
  }

  private async getItems(list: { id: string } | null): Promise<Array<{ id: string; fields: ListaPruebaFields }>> {
    if (!list?.id) return [];
    const res = await this.client
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .expand("fields($select=Title,SectorGeneral,AreaNovedad,TituloNovedad,DetalleNovedad,Resumen)")
      .top(200)
      .get();

    // // DEBUG si lo necesitás:
    // if ((res?.value ?? []).length) {
    //   console.log('fields keys ejemplo:', Object.keys(res.value[0].fields || {}));
    // }

    return (res?.value ?? []) as any[];
  }

  private pickMostFrequent(values: string[]): string {
    const counts: Record<string, number> = {};
    for (const v of values) counts[v] = (counts[v] || 0) + 1;
    let best = 'General', max = -1;
    for (const [k, v] of Object.entries(counts)) if (v > max) { max = v; best = k; }
    return best;
  }
}
