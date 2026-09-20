// =====================================================================
//  Ali Closet · Edge Function "asistente-ia"
//  Llama a Google Gemini para responder sobre el inventario.
//
//  CÓMO DESPLEGARLA:
//  1) Instalá la CLI de Supabase y logueate.
//  2) supabase functions new asistente-ia
//     Pegá este archivo en supabase/functions/asistente-ia/index.ts
//  3) Guardá tu clave de Gemini como secreto:
//       supabase secrets set GEMINI_API_KEY=tu_clave_de_google_ai_studio
//  4) supabase functions deploy asistente-ia
//
//  La clave de Gemini vive SOLO aquí (en el servidor), nunca en el navegador.
//
//  NOTA IMPORTANTE (por qué antes "no funcionaba"):
//  El modelo "gemini-2.0-flash" fue retirado por Google (junio 2026) y ahora
//  devuelve 404. Abajo se usa un modelo actual con una lista de respaldo:
//  si uno falla, prueba el siguiente automáticamente.
// =====================================================================

const GEMINI_API_KEY = Deno.env.get("GEMINI_API_KEY");

// Se prueban en orden; el primero que responda gana.
const MODELOS = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-flash-latest"];

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const json = (obj: unknown, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { ...cors, "Content-Type": "application/json" } });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });

  if (!GEMINI_API_KEY) {
    return json({ error: "Falta GEMINI_API_KEY en los secretos de Supabase." }, 500);
  }

  try {
    const { pregunta, datos } = await req.json();
    if (!pregunta) return json({ error: "Falta la pregunta." }, 400);

    const sistema = `Sos el asistente del sistema de inventario de "Ali Closet",
una tienda de ropa de mujer y maquillaje. Respondé en español, claro, breve y amable.
Usá SOLO los datos entregados; si algo no está, decilo con honestidad.
Los montos van en dólares.`;

    const prompt = `${sistema}

DATOS ACTUALES (JSON):
${JSON.stringify(datos)}

PREGUNTA DE LA DUEÑA:
${pregunta}`;

    let ultimoError = "";

    for (const modelo of MODELOS) {
      const r = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": GEMINI_API_KEY },
          body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
        },
      );

      const j = await r.json();

      if (r.ok) {
        const respuesta =
          j?.candidates?.[0]?.content?.parts?.[0]?.text ??
          "No pude generar una respuesta esta vez.";
        return json({ respuesta, modelo });
      }

      // Guardamos el error y probamos el siguiente modelo (404 = modelo retirado).
      ultimoError = j?.error?.message || `HTTP ${r.status}`;
      if (r.status !== 404 && r.status !== 400) break; // 401/403 = clave inválida: no sirve reintentar
    }

    return json({ error: `Gemini no respondió: ${ultimoError}` }, 502);
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
});