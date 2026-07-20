export async function onRequest(context) {
  const { request, env } = context;
  const url = new URL(request.url);
  
  // The target API URL
  const API_BASE = 'https://fm-operations-api.facilitycare-audydental.workers.dev';
  
  // Rewrite the URL to point to the worker
  const targetUrl = new URL(url.pathname + url.search, API_BASE);
  
  // Fetch from the worker
  const headers = new Headers(request.headers);
  headers.delete('Host');
  headers.delete('Origin');
  headers.delete('Referer');
  headers.delete('Content-Length');
  
  // Fetch from the worker directly
  return fetch(targetUrl, {
    method: request.method,
    headers: headers,
    body: request.method !== 'GET' && request.method !== 'HEAD' ? request.body : null,
    redirect: 'manual',
    duplex: 'half'
  });
}
