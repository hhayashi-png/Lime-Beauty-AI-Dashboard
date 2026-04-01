addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});

const GAS_URL = 'https://script.google.com/macros/s/AKfycbxUBsEzbrUcEZepMfAgEt34jIK3DQjYhDzXA1VFvSP4xvcM5BuW_u_qt4GClxzjBdp9/exec';

async function handleRequest(request) {
  if (request.method === 'GET') {
    return new Response(JSON.stringify({status: 'ok', message: 'LINE Webhook Relay is running'}), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
  }

  if (request.method === 'POST') {
    try {
      const body = await request.text();

      fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: body,
        redirect: 'follow'
      }).catch(err => console.error('GAS fetch error:', err));

      return new Response(JSON.stringify({status: 'ok'}), {
        status: 200,
        headers: { 'Content-Type': 'application/json' }
      });
    } catch (err) {
      return new Response(JSON.stringify({status: 'ok'}), {
        status: 200,
        headers: { 'Content-Type': 'application/json' }
      });
    }
  }

  return new Response(JSON.stringify({status: 'ok'}), {
    status: 200,
    headers: { 'Content-Type': 'application/json' }
  });
}
