export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    
    // Proxy /gas/* requests to Google Apps Script to avoid CORS
    if (url.pathname.startsWith('/gas/')) {
      const gasUrl = 'https://script.google.com/macros/s/' + 
        url.pathname.replace('/gas/', '') + url.search;
      const response = await fetch(gasUrl, {
        method: request.method,
        headers: { 'Content-Type': 'application/json' }
      });
      const body = await response.text();
      return new Response(body, {
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        }
      });
    }
    
    // All other requests serve static assets
    return env.ASSETS.fetch(request);
  }
};