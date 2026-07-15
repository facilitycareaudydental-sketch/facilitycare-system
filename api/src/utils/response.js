// Standard response helpers

export function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin || '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Max-Age': '86400',
  };
}

export function ok(data, status = 200, origin) {
  return new Response(JSON.stringify({ success: true, data }), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function paginated(data, total, page, limit, origin) {
  return new Response(JSON.stringify({
    success: true,
    data,
    pagination: { total, page, limit, pages: Math.ceil(total / limit) },
  }), {
    status: 200,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function error(message, status = 400, origin) {
  return new Response(JSON.stringify({ success: false, error: message }), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function unauthorized(origin) {
  return error('Unauthorized', 401, origin);
}

export function forbidden(origin) {
  return error('Forbidden - insufficient permissions', 403, origin);
}

export function notFound(origin) {
  return error('Not found', 404, origin);
}

export function options(origin) {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(origin),
  });
}
