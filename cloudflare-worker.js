/**
 * Cloudflare Worker endpoint for eBay Marketplace Account Deletion.
 * Add EBAY_ENDPOINT_URL as a Worker variable and
 * EBAY_VERIFICATION_TOKEN as a Worker secret before deployment.
 */
export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname !== "/ebay/account-deletion") {
      return new Response("Not found", { status: 404 });
    }

    if (request.method === "GET") {
      const challengeCode = url.searchParams.get("challenge_code");
      if (!challengeCode || !env.EBAY_VERIFICATION_TOKEN || !env.EBAY_ENDPOINT_URL) {
        return new Response("Missing challenge or Worker configuration", { status: 400 });
      }

      const input = `${challengeCode}${env.EBAY_VERIFICATION_TOKEN}${env.EBAY_ENDPOINT_URL}`;
      const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(input));
      const challengeResponse = [...new Uint8Array(digest)]
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("");

      return Response.json({ challengeResponse });
    }

    if (request.method === "POST") {
      // Do not log or retain the notification body; it may contain user identifiers.
      await request.arrayBuffer();
      return new Response(null, { status: 204 });
    }

    return new Response("Method not allowed", {
      status: 405,
      headers: { Allow: "GET, POST" },
    });
  },
};
