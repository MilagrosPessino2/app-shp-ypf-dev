import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export type ListaPruebaFields = {
  Area?: string;           // ajusta a tus columnas reales
  Titulo?: string;
  DetalleHtml?: string;
  ImagenesJson?: string;        // opcional: JSON con URLs
  ImagenesSeparadas?: string;   // opcional: "url1;url2"
};

export type NovedadItem = {
  tituloNovedad: string;
  detalleNovedad: string;     // HTML
  imagenesNovedad: string[];  // URLs
};

export type NovedadesData = {
  area: string;
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

  public async fetchNovedades(listDisplayName: string): Promise<NovedadesData> {
    const items = await this.getItems(listDisplayName);
    const list  = await this.getList(listDisplayName);

    const salida: NovedadItem[] = [];
    for (const it of items) {
      const f = (it.fields || {}) as ListaPruebaFields;
      const urls: string[] = [];

      if (f.ImagenesJson) {
        try {
          const arr = JSON.parse(f.ImagenesJson);
          if (Array.isArray(arr)) arr.forEach(u => { if (typeof u === 'string' && u) urls.push(u); });
        } catch {}
      }
      if (f.ImagenesSeparadas) {
        f.ImagenesSeparadas.split(';').map(s => s.trim()).filter(Boolean).forEach(u => urls.push(u));
      }

      if (list?.id) {
        try {
          const resp = await this.client
            .api(`/sites/${this.siteId}/lists/${list.id}/items/${it.id}/driveItem/children`)
            .select('name,@microsoft.graph.downloadUrl')
            .get();
          (resp?.value ?? []).forEach((ch: any) => {
            const dl = ch['@microsoft.graph.downloadUrl'];
            if (dl) urls.push(dl);
          });
        } catch {}
      }

      salida.push({
        tituloNovedad: (f.Titulo || '(Sin título)').trim(),
        detalleNovedad: f.DetalleHtml || '',
        imagenesNovedad: Array.from(new Set(urls)),
      });
    }

    const area = this.pickArea(items);
    return { area, novedad: [{ areaNovedad: area || 'General', items: salida }] };
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

  private async getItems(listDisplayName: string): Promise<Array<{ id: string; fields: ListaPruebaFields }>> {
    const list = await this.getList(listDisplayName);
    if (!list?.id) return [];
    const res = await this.client
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .expand('fields')
      .top(200)
      .get();
    return (res?.value ?? []) as any[];
  }

  private pickArea(items: Array<{ fields: ListaPruebaFields }>): string {
    const counts: Record<string, number> = {};
    for (const it of items) {
      const a = (it.fields?.Area ?? 'General').trim();
      counts[a] = (counts[a] || 0) + 1;
    }
    let best = 'General', max = -1;
    for (const [k, v] of Object.entries(counts)) if (v > max) { max = v; best = k; }
    return best;
  }
}
