import { createClient } from "npm:@supabase/supabase-js@2";

const allowedOrigins = new Set([
  "https://arnaudjmpro.github.io",
  "https://localhost:3000",
  "http://localhost:3000",
]);

const maxBase64Length = 8_000_000;
const sessionPattern = /^[a-f0-9]{64}$/;
const allowedMimeTypes = new Set([
  "image/jpeg",
  "image/png",
  "image/webp",
]);

function corsHeaders(request: Request) {
  const origin = request.headers.get("origin") || "";
  const allowedOrigin = allowedOrigins.has(origin)
    ? origin
    : "https://arnaudjmpro.github.io";

  return {
    "Access-Control-Allow-Origin": allowedOrigin,
    "Access-Control-Allow-Headers":
      "authorization, x-client-info, apikey, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Vary": "Origin",
  };
}

function jsonResponse(
  request: Request,
  body: Record<string, unknown>,
  status = 200,
) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders(request),
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

Deno.serve(async (request: Request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", {
      headers: corsHeaders(request),
    });
  }

  if (request.method !== "POST") {
    return jsonResponse(
      request,
      { error: "Méthode non autorisée." },
      405,
    );
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL");
  const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

  if (!supabaseUrl || !serviceRoleKey) {
    return jsonResponse(
      request,
      { error: "Configuration serveur incomplète." },
      500,
    );
  }

  let body: Record<string, unknown>;

  try {
    body = await request.json();
  } catch (_) {
    return jsonResponse(
      request,
      { error: "Requête JSON invalide." },
      400,
    );
  }

  const action = String(body.action || "");
  const sessionId = String(body.sessionId || "");

  if (!sessionPattern.test(sessionId)) {
    return jsonResponse(
      request,
      { error: "Session photo invalide." },
      400,
    );
  }

  const supabase = createClient(
    supabaseUrl,
    serviceRoleKey,
    {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    },
  );

  const now = new Date();
  const expiresAt = new Date(
    now.getTime() + 10 * 60 * 1000,
  ).toISOString();

  await supabase
    .from("tableia_photo_sessions")
    .delete()
    .lt("expires_at", now.toISOString());

  if (action === "create") {
    const { error } = await supabase
      .from("tableia_photo_sessions")
      .upsert(
        {
          session_id: sessionId,
          expires_at: expiresAt,
        },
        {
          onConflict: "session_id",
          ignoreDuplicates: true,
        },
      );

    if (error) {
      console.error(error);
      return jsonResponse(
        request,
        { error: "Impossible de préparer la session photo." },
        500,
      );
    }

    return jsonResponse(request, { status: "created" });
  }

  if (action === "upload") {
    const imageBase64 = String(body.imageBase64 || "");
    const imageMimeType = String(body.imageMimeType || "");

    if (!allowedMimeTypes.has(imageMimeType)) {
      return jsonResponse(
        request,
        { error: "Format d'image non autorisé." },
        400,
      );
    }

    if (
      !imageBase64 ||
      imageBase64.length > maxBase64Length ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(imageBase64)
    ) {
      return jsonResponse(
        request,
        { error: "Photo invalide ou trop volumineuse." },
        400,
      );
    }

    const { error } = await supabase
      .from("tableia_photo_sessions")
      .upsert(
        {
          session_id: sessionId,
          image_base64: imageBase64,
          image_mime_type: imageMimeType,
          expires_at: expiresAt,
        },
        { onConflict: "session_id" },
      );

    if (error) {
      console.error(error);
      return jsonResponse(
        request,
        { error: "Impossible de transmettre la photo." },
        500,
      );
    }

    return jsonResponse(request, { status: "uploaded" });
  }

  if (action === "poll") {
    const { data, error } = await supabase
      .from("tableia_photo_sessions")
      .select("image_base64, image_mime_type, expires_at")
      .eq("session_id", sessionId)
      .gt("expires_at", now.toISOString())
      .maybeSingle();

    if (error) {
      console.error(error);
      return jsonResponse(
        request,
        { error: "Impossible de récupérer la photo." },
        500,
      );
    }

    if (!data?.image_base64 || !data?.image_mime_type) {
      return jsonResponse(request, { status: "waiting" });
    }

    await supabase
      .from("tableia_photo_sessions")
      .delete()
      .eq("session_id", sessionId);

    return jsonResponse(request, {
      status: "ready",
      imageBase64: data.image_base64,
      imageMimeType: data.image_mime_type,
    });
  }

  if (action === "cancel") {
    await supabase
      .from("tableia_photo_sessions")
      .delete()
      .eq("session_id", sessionId);

    return jsonResponse(request, { status: "cancelled" });
  }

  return jsonResponse(
    request,
    { error: "Action inconnue." },
    400,
  );
});
