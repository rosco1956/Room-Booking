// _worker.js
var worker_default = {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") {
      return new Response(null, {
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type"
        }
      });
    }
    if (url.pathname.startsWith("/gas/")) {
      const gasUrl = "https://script.google.com/macros/s/" + url.pathname.slice(5) + url.search;
      const isCreateZoom = url.search.includes("action=createZoom");
      const isCalendarFeed = url.search.includes("action=ical") || url.search.includes("action=zoomical");
      try {
        const controller = new AbortController();
        const timeoutId = isCreateZoom ? setTimeout(() => controller.abort(), 115e3) : setTimeout(() => controller.abort(), 55e3);
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

        // Calendar feed requests: pass through as text/calendar, no JSON validation
        if (isCalendarFeed) {
          if (!body.trim().startsWith("BEGIN:VCALENDAR")) {
            const preview = body.slice(0, 500);
            console.log("GAS calendar feed returned unexpected body (status " + response.status + "): " + preview);
            return new Response(JSON.stringify({ ok: false, error: "GAS returned unexpected calendar body", status: response.status, preview }), {
              status: 200,
              headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
            });
          }
          return new Response(body, {
            status: 200,
            headers: {
              "Content-Type": "text/calendar; charset=utf-8",
              "Access-Control-Allow-Origin": "*"
            }
          });
        }

        const isJson = body.trim().startsWith("{") || body.trim().startsWith("[");
        if (!isJson) {
          const preview = body.slice(0, 500);
          console.log("GAS non-JSON response (status " + response.status + "): " + preview);
          return new Response(JSON.stringify({ ok: false, error: "GAS returned non-JSON", status: response.status, preview }), {
            status: 200,
            headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
          });
        }
        return new Response(body, {
          status: 200,
          headers: {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
          }
        });
      } catch (e) {
        const timedOut = e.name === "AbortError";
        return new Response(JSON.stringify({
          ok: false,
          error: timedOut ? "Request timed out \u2014 GAS may still be processing" : e.message
        }), {
          status: 200,
          headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
        });
      }
    }
    return env.ASSETS.fetch(request);
  }
};
export {
  worker_default as default
};
