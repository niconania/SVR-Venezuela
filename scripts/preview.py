#!/usr/bin/env python3
"""Servidor de previsualización para svr-venezuela.com

  Uso:   python3 scripts/preview.py          (solo esta máquina)
         python3 scripts/preview.py 9000     (otro puerto)
         python3 scripts/preview.py --lan     (visible desde la red local,
                                               para abrirlo desde otro equipo)

Hace dos cosas que `python3 -m http.server` no hace:

  1. Manda todo sin caché, así que al refrescar siempre ves la versión
     actual y no una guardada por el navegador.
  2. Recarga la pestaña sola cuando cambia un archivo del sitio, para no
     tener que estar refrescando a mano.

Es solo para trabajar en local. No se usa en producción: GitHub Pages
sirve los archivos directamente.
"""

import http.server
import os
import socketserver
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --lan lo hace visible desde otros equipos de la misma red.
ARGS   = [a for a in sys.argv[1:] if a != '--lan']
LAN    = '--lan' in sys.argv
PUERTO = int(ARGS[0]) if ARGS else 8080
HOST   = '0.0.0.0' if LAN else '127.0.0.1'

# Se vigilan los archivos que editamos, no los 47 MB de imágenes y video.
VIGILADAS = ('.html', '.json', '.xml', '.css', '.js')
IGNORAR = {'.git', 'node_modules', '.github'}

RECARGA = """
<script>
/* Recarga automática — solo en localhost, no viaja a producción. */
(function () {
  var actual = null;
  var espera = 1000;

  /* Se pregunta de nuevo SOLO cuando la respuesta anterior llegó, y se
     espera al menos lo que tardó la última. Sin esto, en una conexión
     lenta las peticiones se apilan, ocupan las 6 conexiones que el
     navegador permite por sitio, y las imágenes no cargan nunca. */
  function preguntar() {
    if (document.hidden) { setTimeout(preguntar, 2000); return; }
    var t0 = Date.now();
    fetch('/__cambios', { cache: 'no-store' })
      .then(function (r) { return r.text(); })
      .then(function (sello) {
        if (actual !== null && sello !== actual) { location.reload(); return; }
        actual = sello;
        var tardo = Date.now() - t0;
        espera = Math.min(15000, Math.max(1000, tardo * 4));
        setTimeout(preguntar, espera);
      })
      .catch(function () { setTimeout(preguntar, 5000); });
  }

  // Arranca cuando la página terminó de cargar, para no competir
  // con las imágenes por las conexiones disponibles.
  if (document.readyState === 'complete') setTimeout(preguntar, 1500);
  else window.addEventListener('load', function () { setTimeout(preguntar, 1500); });
})();
</script>
"""


def sello_de_cambios():
    """Huella del estado actual: si cambia, algo se editó."""
    marcas = []
    for carpeta, subs, archivos in os.walk(RAIZ):
        subs[:] = [s for s in subs if s not in IGNORAR and not s.startswith('.')]
        for a in archivos:
            if a.endswith(VIGILADAS):
                try:
                    marcas.append(str(os.stat(os.path.join(carpeta, a)).st_mtime))
                except OSError:
                    pass
    return str(hash('|'.join(sorted(marcas))))


class Preview(http.server.SimpleHTTPRequestHandler):

    # HTTP/1.1 mantiene viva la conexión. Con HTTP/1.0 (el de fábrica) el
    # navegador abre una conexión nueva por cada archivo: en local no se
    # nota, pero a través de un túnel remoto la página no termina de cargar.
    protocol_version = 'HTTP/1.1'

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=RAIZ, **kw)

    def do_GET(self):
        if self.path.startswith('/__cambios'):
            cuerpo = sello_de_cambios().encode()
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.send_header('Content-Length', str(len(cuerpo)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(cuerpo)
            return

        ruta = self.translate_path(self.path)
        if os.path.isdir(ruta):
            ruta = os.path.join(ruta, 'index.html')

        # El HTML se sirve con el script de recarga inyectado al final.
        if ruta.endswith('.html') and os.path.exists(ruta):
            with open(ruta, 'rb') as f:
                html = f.read()
            if b'</body>' in html:
                html = html.replace(b'</body>', RECARGA.encode() + b'</body>', 1)
            else:
                html += RECARGA.encode()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(html)
            return

        super().do_GET()

    def end_headers(self):
        # Nada de caché: al refrescar siempre se ve lo último.
        if 'Cache-Control' not in self._headers_buffer_names():
            self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def _headers_buffer_names(self):
        return [h.split(b':')[0].decode('latin-1')
                for h in getattr(self, '_headers_buffer', []) if b':' in h]

    def log_message(self, formato, *args):
        # Silencia el ruido de cada imagen; solo interesan los errores.
        if len(args) > 1 and str(args[1]).startswith(('4', '5')):
            sys.stderr.write('  %s %s\n' % (args[1], args[0]))


class Servidor(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


if __name__ == '__main__':
    os.chdir(RAIZ)
    with Servidor((HOST, PUERTO), Preview) as httpd:
        print('  SVR Venezuela — previsualización')
        print(f'  http://localhost:{PUERTO}')
        if LAN:
            import socket
            try:
                s_ = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                s_.connect(('8.8.8.8', 80)); ip = s_.getsockname()[0]; s_.close()
                print(f'  http://{ip}:{PUERTO}   ← desde otro equipo de la red')
            except OSError:
                pass
        print('  Se recarga sola al guardar cambios. Ctrl+C para detener.\n')
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print('\n  Detenido.')
