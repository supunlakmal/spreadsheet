export default async (req, context) => {
  // Only allow GET requests
  if (req.method !== 'GET') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' }
    });
  }

  // Get API key from environment variable
  const apiKey = Netlify.env.get('METERED_API_KEY');

  if (!apiKey) {
    console.warn('METERED_API_KEY not configured - returning STUN-only fallback');
    return new Response(JSON.stringify({
      iceServers: [
        { urls: 'stun:stun.l.google.com:19302' },
        { urls: 'stun:stun1.l.google.com:19302' }
      ],
      fallback: true
    }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
  }

  try {
    // Fetch credentials from Metered.ca
    const response = await fetch(
      `https://spreadsheetlive.metered.live/api/v1/turn/credentials?apiKey=${apiKey}`
    );

    if (!response.ok) {
      throw new Error(`Metered API error: ${response.status}`);
    }

    const iceServers = await response.json();

    return new Response(JSON.stringify({ iceServers }), {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'private, max-age=3600'
      }
    });
  } catch (error) {
    console.error('Error fetching TURN credentials:', error);

    // Return fallback STUN-only config
    return new Response(JSON.stringify({
      iceServers: [
        { urls: 'stun:stun.l.google.com:19302' },
        { urls: 'stun:stun1.l.google.com:19302' }
      ],
      fallback: true
    }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
  }
};

export const config = {
  path: '/api/turn-credentials'
};
