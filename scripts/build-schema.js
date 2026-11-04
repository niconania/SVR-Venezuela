#!/usr/bin/env node
/* Genera el bloque JSON-LD de productos a partir de products.json y lo
   inserta en tienda/index.html entre los marcadores. El schema vive donde
   el precio se ve: la tienda. products.json es la única
   fuente: si cambia un precio o se agota algo, se corre esto y listo.

   Uso:  node scripts/build-schema.js                                     */

const fs   = require('fs');
const path = require('path');
const RAIZ = path.join(__dirname, '..');
const SITIO = 'https://svr-venezuela.com';

const INICIO = '  <!-- catalogo:inicio (generado por scripts/build-schema.js — no editar a mano) -->';
const FIN    = '  <!-- catalogo:fin -->';

const cat = JSON.parse(fs.readFileSync(path.join(RAIZ, 'products.json'), 'utf8'));

const schema = {
  '@context': 'https://schema.org',
  '@type': 'ItemList',
  name: 'Catálogo SVR Venezuela',
  numberOfItems: cat.productos.length,
  itemListElement: cat.productos.map((p, i) => {
    const portada = (p.media.find(m => m.type === 'img') || {}).src;
    return {
      '@type': 'ListItem',
      position: i + 1,
      item: {
        '@type': 'Product',
        name: `SVR ${p.linea} ${p.nombre} ${p.presentacion}`,
        sku: p.sku,
        description: p.descripcion,
        ...(portada ? { image: `${SITIO}/${portada}` } : {}),
        brand: { '@type': 'Brand', name: 'SVR' },
        category: p.linea,
        offers: {
          '@type': 'Offer',
          url: `${SITIO}/tienda/`,
          priceCurrency: cat.moneda,
          price: p.precio.toFixed(2),
          availability: p.disponible === false
            ? 'https://schema.org/OutOfStock'
            : 'https://schema.org/InStock',
          seller: { '@type': 'Organization', name: 'SVR Venezuela' }
        }
      }
    };
  })
};

const bloque = INICIO + '\n  <script type="application/ld+json">\n'
  + JSON.stringify(schema, null, 2).split('\n').map(l => '  ' + l).join('\n')
  + '\n  </' + 'script>\n' + FIN;

const F = path.join(RAIZ, 'tienda', 'index.html');
let h = fs.readFileSync(F, 'utf8');

if (h.includes(INICIO) && h.includes(FIN)) {
  const a = h.indexOf(INICIO), b = h.indexOf(FIN) + FIN.length;
  h = h.slice(0, a) + bloque + h.slice(b);                 // reemplaza el anterior
} else {
  const ancla = '  <link rel="preconnect" href="https://fonts.googleapis.com">';
  if (!h.includes(ancla)) { console.error('✗ no encontré dónde insertar'); process.exit(1); }
  h = h.replace(ancla, () => bloque + '\n' + ancla);       // primera vez
}

fs.writeFileSync(F, h);
console.log(`✓ schema de ${cat.productos.length} productos escrito en tienda/index.html`);
const agotados = cat.productos.filter(p => p.disponible === false);
console.log(`  · disponibles: ${cat.productos.length - agotados.length}   agotados: ${agotados.length}`);
