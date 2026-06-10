export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type',
        }
      });
    }

    if (url.pathname.startsWith('/gas/')) {
      const gasUrl = 'https://script.google.com/macros/s/' +
        url.pathname.slice(5) + url.search;

      // createZoom takes 60-90s — fire and forget, let polling find the result
      if (url.searchParams.get('action') === 'createZoom') {
        fetch(gasUrl, { method: 'GET', redirect: 'follow' }).catch(() => {});
        return new Response(JSON.stringify({ok:true,queued:true}), {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
        });
      }

      try {
        const response = await fetch(gasUrl, { method: 'GET', redirect: 'follow' });
        const body = await response.text();
        return new Response(body, {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
        });
      } catch(e) {
        return new Response(JSON.stringify({ok:false,error:e.message}), {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
        });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
