# Stable eBay notification endpoint

This Worker provides a stable `https://<worker>.<account>.workers.dev/ebay/account-deletion` URL. It answers eBay's endpoint challenge and acknowledges deletion notifications without logging or retaining their bodies.

## Create and deploy

1. Sign in to Cloudflare and open **Workers & Pages**.
2. Create a Worker, choose the dashboard editor, and replace the starter code with `cloudflare-worker.js`.
3. In the Worker's settings, add a text variable named `EBAY_ENDPOINT_URL` with the final public URL, including `/ebay/account-deletion`.
4. Add a secret named `EBAY_VERIFICATION_TOKEN` with the same token entered in eBay. Never put this token in the source file.
5. Deploy. Use the Worker's stable `workers.dev` URL as the Marketplace Account Deletion endpoint in eBay. Save the endpoint to trigger eBay's challenge validation.
6. Use eBay's **Send Test Notification** control and confirm the request succeeds.

The endpoint URL must exactly match `EBAY_ENDPOINT_URL`, including scheme, hostname, and path, because eBay includes it in the challenge hash. Do not enable Cloudflare Access or other authentication on this route; eBay must call it without signing in.

This MVP does not store the notification body. eBay recommends verifying the `X-EBAY-SIGNATURE` on notifications; add that verification before extending the handler to process or retain any notification data. The current project does not retain eBay user identifiers.

Cloudflare Workers has a free tier with request limits. A `workers.dev` URL is stable and works without buying a domain, though Cloudflare describes it as suited to personal or hobby projects rather than business-critical services.
