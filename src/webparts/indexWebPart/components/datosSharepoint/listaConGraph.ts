import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

/** Columnas visibles en tu lista (display name) y sus posibles internos */
export type ListaPruebaFields = {
  // OJO: en muchas listas "SectorGeneral" es realmente el campo interno Title renombrado
  SectorGeneral?: string;   // si es una columna nueva con ese internal name
  AreaNovedad?: string;
  TituloNovedad?: string;   // si es una columna nueva con ese internal name
  DetalleNovedad?: string;
  Resumen?: string;

  // SIEMPRE agregamos Title por si "SectorGeneral" o "TituloNovedad" son solo renombres del Title
  Title?: string;
};

export type NovedadItem = {
  tituloNovedad: string;
  detalleNovedad: string;     // HTML
  imagenesNovedad: string[];  // URLs directas de los adjuntos
};

export type NovedadesData = {
  area: string; // SectorGeneral (o Title) más frecuente
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

  /** Trae y mapea: agrupa por AreaNovedad y setea SectorGeneral (o Title) más frecuente */
  public async fetchNovedades(listDisplayName: string): Promise<NovedadesData> {
    const list  = await this.getList(listDisplayName);
    const items = await this.getItems(list);

    const groups = new Map<string, NovedadItem[]>();

    for (const it of items) {
      const f = (it.fields || {}) as ListaPruebaFields;

      // Fallbacks robustos contra renombres:
      const areaNovedad      = (f.AreaNovedad ?? 'General').trim();
      const titulo           = (f.TituloNovedad ?? f.Title ?? '(Sin título)').trim();
      const detalleHtml      = f.DetalleNovedad ?? '';

      const urls: string[] = [];

      // Adjuntos por Graph (carpeta del drive del item)
      if (list?.id) {
        try {
          const resp = await this.client
            .api(`/sites/${this.siteId}/lists/${list.id}/items/${it.id}/driveItem/children`)
            .select('name,@microsoft.graph.downloadUrl,file')
            .get();

          const files = resp?.value ?? [];
          for (const ch of files) {
            const dl = ch['@microsoft.graph.downloadUrl'];
            if (dl) urls.push(dl);
          }
        } catch { /* sin adjuntos o sin permisos */ }
      }

      const arr = groups.get(areaNovedad) ?? [];
      arr.push({
        tituloNovedad: titulo,
        detalleNovedad: detalleHtml,
        imagenesNovedad: Array.from(new Set(urls)),
      });
      groups.set(areaNovedad, arr);
    }

    // SectorGeneral = valor más frecuente entre (SectorGeneral || Title)
    const sectorGeneral = this.pickMostFrequent(
      items.map((it) => ((it.fields?.SectorGeneral ?? it.fields?.Title) ?? 'General').trim())
    );

    const novedad = Array.from(groups.entries()).map(([areaNovedad, items]) => ({
      areaNovedad,
      items,
    })).sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

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
    // Seleccionamos explícitamente los campos (incluyendo Title por los renombres)
    const res = await this.client
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .expand("fields($select=Title,SectorGeneral,AreaNovedad,TituloNovedad,DetalleNovedad,Resumen)")
      .top(200)
      .get();

    const value = (res?.value ?? []) as any[];

    // Debug opcional para ver qué keys devuelve tu lista (si seguís con dudas)
    // console.log('Ejemplo fields keys:', value[0]?.fields && Object.keys(value[0].fields));

    return value;
  }

  private pickMostFrequent(values: string[]): string {
    const counts: Record<string, number> = {};
    for (const v of values) counts[v] = (counts[v] || 0) + 1;
    let best = 'General', max = -1;
    for (const [k, v] of Object.entries(counts)) if (v > max) { max = v; best = k; }
    return best;
  }
}
