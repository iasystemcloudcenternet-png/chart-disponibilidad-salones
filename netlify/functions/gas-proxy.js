const GAS_URL = 'https://script.google.com/macros/s/AKfycbzfuCAvWbTQ9SKFP9aEQsHJ2bl_hhNrK1K4R5jBnrsag8E2BztdiMEckY3NU28EYv96/exec';

exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return json(405, { ok: false, message: 'Método no permitido.' });
  }

  try {
    const payload = JSON.parse(event.body || '{}');

    if (
      !payload ||
      typeof payload.fn !== 'string' ||
      !Array.isArray(payload.args)
    ) {
      return json(400, {
        ok: false,
        message: 'Solicitud no válida.'
      });
    }

    const response = await fetch(GAS_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });

    const text = await response.text();

    let data;

    try {
      data = JSON.parse(text);
    } catch (_) {
      return json(502, {
        ok: false,
        message: 'Apps Script devolvió una respuesta no válida.'
      });
    }

    return json(response.ok ? 200 : 502, data);

  } catch (error) {
    return json(500, {
      ok: false,
      message:
        error && error.message
          ? error.message
          : 'Error de comunicación con Apps Script.'
    });
  }
};

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    },
    body: JSON.stringify(body)
  };
}
