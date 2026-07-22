const RESPUESTAS_JSON = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store, no-cache, must-revalidate"
};

export default {
  async fetch(request, env) {
    const urlSolicitud = new URL(request.url);

    if (urlSolicitud.pathname !== "/api/archivos") {
      return respuestaJson({ error: "Recurso no encontrado" }, 404);
    }

    if (request.method !== "GET") {
      return respuestaJson({ error: "Método no permitido" }, 405, { allow: "GET" });
    }

    try {
      const prefijo = normalizarPrefijo(env.PREFIJO_REMOTO || "PROVEEPERU/Sistema Visual/");
      const dominioPublico = (env.DOMINIO_PUBLICO || "https://actualizaciones.fralsetech.com").replace(/\/+$/, "");
      const archivos = [];
      let cursor;

      do {
        const pagina = await env.SISTEMA_VISUAL_BUCKET.list({
          prefix: prefijo,
          cursor,
          limit: 1000
        });

        for (const objeto of pagina.objects) {
          const ruta = objeto.key.slice(prefijo.length);
          if (!ruta || ruta.endsWith("/")) continue;

          archivos.push({
            ruta,
            clave: objeto.key,
            tamano: objeto.size,
            etag: objeto.etag,
            fechaModificacion: objeto.uploaded ? objeto.uploaded.toISOString() : null,
            url: `${dominioPublico}/${codificarRuta(objeto.key)}`
          });
        }

        if (pagina.truncated && !pagina.cursor) {
          throw new Error("R2 indicó más resultados pero no devolvió cursor");
        }
        cursor = pagina.truncated ? pagina.cursor : undefined;
      } while (cursor);

      archivos.sort((a, b) => a.ruta.localeCompare(b.ruta));
      return respuestaJson({
        prefijo,
        fechaConsulta: new Date().toISOString(),
        archivos
      });
    } catch (error) {
      console.error("Error al listar objetos R2", error);
      return respuestaJson({ error: "No se pudo consultar el repositorio de archivos" }, 500);
    }
  }
};

function normalizarPrefijo(prefijo) {
  return prefijo.endsWith("/") ? prefijo : `${prefijo}/`;
}

function codificarRuta(ruta) {
  return ruta.split("/").map(segmento => encodeURIComponent(segmento)).join("/");
}

function respuestaJson(contenido, estado, cabecerasAdicionales) {
  return new Response(JSON.stringify(contenido), {
    status: estado || 200,
    headers: { ...RESPUESTAS_JSON, ...(cabecerasAdicionales || {}) }
  });
}
