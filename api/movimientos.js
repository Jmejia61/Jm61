// Backend serverless para Vercel — solo lectura.
// Consulta la base de Notion y devuelve los movimientos a la app.
// El token se lee de variables de entorno, nunca del código.

const NOTION_API = 'https://api.notion.com/v1';
const NOTION_VERSION = '2022-06-28';

const DATABASE_ID = process.env.NOTION_DATABASE_ID;
const TOKEN = process.env.NOTION_TOKEN;

// Convierte una página de Notion al formato que usa la app
function mapPage(page) {
  const p = page.properties || {};
  const titulo = p['Descripción'] && p['Descripción'].title;
  const sel = (nombre) => (p[nombre] && p[nombre].select ? p[nombre].select.name : '');

  return {
    id: page.id,
    descripcion: titulo && titulo[0] ? titulo[0].plain_text : '(sin descripción)',
    tipo: sel('Tipo') || 'Gasto',
    persona: sel('Persona'),
    categoria: sel('Categoría'),
    metodo: sel('Método de pago'),
    monto: p['Monto'] && typeof p['Monto'].number === 'number' ? p['Monto'].number : 0,
    fecha: p['Fecha'] && p['Fecha'].date ? p['Fecha'].date.start : ''
  };
}

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'Solo lectura: método no permitido' });
  }

  if (!TOKEN || !DATABASE_ID) {
    return res.status(500).json({
      error: 'Faltan las variables de entorno NOTION_TOKEN o NOTION_DATABASE_ID'
    });
  }

  try {
    const movimientos = [];
    let cursor = undefined;

    // Notion pagina de a 100; recorremos todo el histórico
    do {
      const body = {
        sorts: [{ property: 'Fecha', direction: 'descending' }],
        page_size: 100
      };
      if (cursor) body.start_cursor = cursor;

      const r = await fetch(`${NOTION_API}/databases/${DATABASE_ID}/query`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${TOKEN}`,
          'Notion-Version': NOTION_VERSION,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
      });

      const data = await r.json();
      if (!r.ok) return res.status(r.status).json(data);

      (data.results || []).forEach((pg) => movimientos.push(mapPage(pg)));
      cursor = data.has_more ? data.next_cursor : undefined;
    } while (cursor);

    // Cache corto: evita golpear Notion en cada recarga
    res.setHeader('Cache-Control', 's-maxage=30, stale-while-revalidate=120');
    return res.status(200).json({
      actualizado: new Date().toISOString(),
      total: movimientos.length,
      movimientos
    });

  } catch (err) {
    return res.status(500).json({ error: String(err && err.message ? err.message : err) });
  }
};
