export async function onRequest(context) {
  const { request, env } = context;
  const url = new URL(request.url);
  
  // The target API URL
  const API_BASE = 'https://fm-operations-api.facilitycare-audydental.workers.dev';
  
  // Rewrite the URL to point to the worker
  const targetUrl = new URL(url.pathname + url.search, API_BASE);
  
  // Create a new request based on the original
  const newRequest = new Request(targetUrl, request);
  
  // Set the original IP and Origin for logging/CORS if needed by the worker
  newRequest.headers.set('X-Forwarded-For', request.headers.get('CF-Connecting-IP') || '');
  newRequest.headers.set('X-Forwarded-Host', url.hostname);
  
  // Fetch from the worker
  return fetch(newRequest);
}
