export default {
  async fetch(request, env, ctx) {
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
      try {
        const response = await fetch(gasUrl, {
          method: 'GET',
          redirect: 'follow',
          headers: {
            'Accept': 'application/json, text/plain, */*',
            'User-Agent': 'Mozilla/5.0'
          }
        });
        const body = await response.text();
        const isJson = body.trim().startsWith('{') || body.trim().startsWith('[');
        if (!isJson) {
          const preview = body.slice(0, 500);
          console.log('GAS non-JSON response (status '+response.status+'): '+preview);
          return new Response(JSON.stringify({ok:false, error:'GAS returned non-JSON', status:response.status, preview:preview}), {
            status: 200,
            headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
          });
        }
        return new Response(body, {
          status: 200,
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          }
        });
      } catch(e) {
        return new Response(JSON.stringify({ok:false, error:e.message}), {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
        });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
