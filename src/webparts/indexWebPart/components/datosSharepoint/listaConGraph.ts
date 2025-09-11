import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

/** ← nombres EXACTOS de tus columnas */
export type ListaPruebaFields = {
  SectorGeneral?: string;    // Texto
  AreaNovedad?: string;      // Texto
  TituloNovedad?: string;    // Texto
  DetalleNovedad?: string;   // Texto enriquecido (HTML)
  Resumen?: string;          // (opcional)
  // "Datos adjuntos" se obtiene vía driveItem/children (Graph), no aparece en fields
};

export type NovedadItem = {
  tituloNovedad: string;
  detalleNovedad: string;     // HTML
  imagenesNovedad: string[];  // URLs directas de los adjuntos
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

  /** Trae y mapea: agrupa por AreaNovedad y setea SectorGeneral (más frecuente) */
  public async fetchNovedades(listDisplayName: string): Promise<NovedadesData> {
    const items = await this.getItems(listDisplayName);
    const list  = await this.getList(listDisplayName);

    // Group by AreaNovedad
    const groups = new Map<string, NovedadItem[]>();

    for (const it of items) {
      const f = (it.fields || {}) as ListaPruebaFields;

      const areaNovedad = (f.AreaNovedad || 'General').trim();
      const titulo = (f.TituloNovedad || '(Sin título)').trim();
      const detalleHtml = f.DetalleNovedad || '';

      const urls: string[] = [];

      // Adjuntos del ítem (carpeta del drive del item)
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
        } catch {/* sin adjuntos o sin permisos */}
      }

      const arr = groups.get(areaNovedad) ?? [];
      arr.push({
        tituloNovedad: titulo,
        detalleNovedad: detalleHtml,
        imagenesNovedad: Array.from(new Set(urls)),
      });
      groups.set(areaNovedad, arr);
    }

    // SectorGeneral = más frecuente en la lista
    const sectorGeneral = this.pickMostFrequent(
      items.map((it) => (it.fields?.SectorGeneral ?? 'General').trim())
    );

    // Armar secciones [{ areaNovedad, items }]
    const novedad = Array.from(groups.entries()).map(([areaNovedad, items]) => ({
      areaNovedad,
      items,
    }));

    // Ordenar por nombre de área (opcional)
    novedad.sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

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

  private async getItems(
    listDisplayName: string
  ): Promise<Array<{ id: string; fields: ListaPruebaFields }>> {
    const list = await this.getList(listDisplayName);
    if (!list?.id) return [];
    const res = await this.client
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .expand('fields')
      .top(200)
      .get();
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
