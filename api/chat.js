export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { messages, system, apiKey } = req.body;

  if (!apiKey || !apiKey.startsWith('sk-')) {
    return res.status(401).json({ error: 'Clé API manquante ou invalide' });
  }

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      system,
      messages
    })
  });

  const data = await response.json();
  return res.status(response.status).json(data);
}
