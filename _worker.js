// _worker.js
//
// Current Apps Script deployment ID.
// Set GAS_ID as a variable in Cloudflare (Worker → Settings → Variables and Secrets)
// to change it without editing code. This constant is only the fallback.
const DEFAULT_GAS_ID = "AKfycbwoOOf_52JVPTf52qruZZNURaeDLBvFDTarqxC27rMY1nL5dlzwkG2ZOnOMYLXtaOQ";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type"
};

function jsonResponse(obj) {
  return new Response(JSON.stringify(obj), {
    status: 200,
    headers: { "Content-Type": "application/json", ...CORS }
  });
}

// Fetches a GAS URL and returns the right kind of response.
async function proxyToGas(gasUrl, { isCalendarFeed, isCreateZoom }) {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), isCreateZoom ? 115e3 : 55e3);
    const response = await fetch(gasUrl, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal,
      headers: {
        "Accept": isCalendarFeed ? "text/calendar, */*" : "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0"
      }
    });
    clearTimeout(timeoutId);
    const body = await response.text();

    if (isCalendarFeed) {
      if (!body.trim().startsWith("BEGIN:VCALENDAR")) {
        const preview = body.slice(0, 500);
        console.log("GAS calendar feed returned unexpected body (status " + response.status + "): " + preview);
        return jsonResponse({ ok: false, error: "GAS returned unexpected calendar body", status: response.status, preview });
      }
      return new Response(body, {
        status: 200,
        headers: {
          "Content-Type": "text/calendar; charset=utf-8",
          "Cache-Control": "no-cache, no-store, must-revalidate",
          ...CORS
        }
      });
    }

    const isJson = body.trim().startsWith("{") || body.trim().startsWith("[");
    if (!isJson) {
      const preview = body.slice(0, 500);
      console.log("GAS non-JSON response (status " + response.status + "): " + preview);
      return jsonResponse({ ok: false, error: "GAS returned non-JSON", status: response.status, preview });
    }
    return new Response(body, {
      status: 200,
      headers: { "Content-Type": "application/json", ...CORS }
    });
  } catch (e) {
    const timedOut = e.name === "AbortError";
    return jsonResponse({
      ok: false,
      error: timedOut ? "Request timed out \u2014 GAS may still be processing" : e.message
    });
  }
}

var worker_default = {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const gasId = (env && env.GAS_ID) || DEFAULT_GAS_ID;

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: CORS });
    }

    // Stable calendar URL — no deployment ID in it.
    //   /calendar?who=Rosswell            → room bookings feed (action=ical)
    //   /calendar?who=Rosswell&type=zoom  → Zoom feed (action=zoomical)
    if (url.pathname === "/calendar" || url.pathname === "/calendar.ics") {
      const params = new URLSearchParams(url.search);
      const type = params.get("type");
      params.delete("type");
      params.set("action", type === "zoom" ? "zoomical" : "ical");
      params.set("_cb", Date.now().toString()); // stop Google caching the GAS response
      const gasUrl = "https://script.google.com/macros/s/" + gasId + "/exec?" + params.toString();
      return proxyToGas(gasUrl, { isCalendarFeed: true, isCreateZoom: false });
    }

    // Existing route — still works for old links and for index.html.
    if (url.pathname.startsWith("/gas/")) {
      const gasUrl = "https://script.google.com/macros/s/" + url.pathname.slice(5) + url.search;
      const isCreateZoom = url.search.includes("action=createZoom");
      const isCalendarFeed = url.search.includes("action=ical") || url.search.includes("action=zoomical");
      return proxyToGas(gasUrl, { isCalendarFeed, isCreateZoom });
    }

    // Static files (index.html etc.)
    if (!env || !env.ASSETS) {
      console.log("ASSETS binding missing — static files are not attached to this Worker");
      return new Response("Static assets are not configured for this Worker (ASSETS binding missing).", {
        status: 500,
        headers: { "Content-Type": "text/plain; charset=utf-8" }
      });
    }
    return env.ASSETS.fetch(request);
  }
};

export {
  worker_default as default
};
