# Pokémon eBay Deal Finder (MVP)

Search active eBay US listings for Pokémon cards, compare each listing with a benchmark you provide, and export the results to an Excel workbook. The default query is `pikachu grey felt hat`. The run makes two separate **text-only** searches sorted by **newly listed first**: a graded pool (condition ID 2750) and a raw pool (condition ID 4000). Each pool searches up to 200 results.

- **Under Market**: listing total price is below its matching benchmark.
- **Regular**: listing total price is at or above its matching benchmark.
- **Needs Review**: eBay's grader/grade characteristics are incomplete, a matching benchmark is missing, or price/shipping cannot be calculated.

The script uses eBay's text search only. Search results are classified by eBay listing condition characteristics and filtered before deduplication. The workbook summary shows the text result counts and how many passed the characteristic filters.

Quality filter: raw listings are included only when eBay's structured card-condition field says **Near Mint** or **Lightly Played**. Other or missing raw-condition values are filtered out. Graded listings are included only when eBay's structured **Professional Grader** characteristic is **PSA** and its structured **Grade** characteristic is **7 or higher**. Other graders and lower grades are filtered out. If eBay identifies a listing as graded but does not provide its grader or grade characteristic, it goes to **Needs Review**. The script fetches item details and uses those eBay characteristics; it never infers a grade from the title.

Raw cards only match the `raw` benchmark row. PSA cards match the exact grade label (for example, `PSA 9`); PSA 9 does not match PSA 8 or raw. Grader and grade come only from eBay's structured listing characteristics.

Each row in `benchmarks.csv` can optionally set `minimum_price` as a per-grade total-price floor (including shipping). Listings below that floor are sent to **Needs Review** with a possible-counterfeit/pricing-error note, and are not treated as deals or emailed. Leave the field blank to disable the floor for that row. These floors are screening hints rather than proof that a listing is counterfeit.

## Important about the benchmark

The public eBay Browse API returns **active asking prices**, not sold prices. eBay's sold-history Marketplace Insights API is limited access. This MVP therefore reads market benchmarks from `benchmarks.csv`; populate it with your preferred trusted source and record that source and date in each row. It intentionally does not invent a market price or use active asks as a sold-price benchmark.

`benchmarks.csv` contains the configured Pikachu Grey Felt Hat benchmark prices. Fill in the source and `as_of` columns to record where and when each benchmark was obtained.

All amounts should be in the same currency as the listing results (USD by default). Listing total price is item price plus shipping when eBay supplies a shipping cost. Sales tax is excluded because it depends on the buyer's location.
For calculated shipping, set `EBAY_ZIP_CODE` in `.env` or pass `--zip`; the script sends the ZIP as eBay's contextual delivery location and fetches full item details when a search summary omits listing characteristics or shipping details. Shipping can still be missing for some listings.

## Setup

1. Install Python 3.10 or newer.
2. In the eBay Developer Program, create a Production application keyset for the Browse API (Application access token / client credentials).
3. Copy `.env.example` to `.env` and enter the Client ID and Client Secret. Keep `.env` private; do not paste secrets into chat or commit them.
   For Sandbox keys, set `EBAY_ENV=sandbox` in `.env` (Sandbox and Production keys are separate).
4. Install dependencies and run:

   ```sh
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
   pip install -r requirements.txt
   python ebay_mewtwo_finder.py
   ```

5. Fill in `benchmarks.csv` before running if you want listings sorted into the two price tabs. Blank/missing benchmark rows are sent to **Needs Review**.

The workbook is saved as `pokemon_ebay_listings.xlsx` in this project folder. Each card-type pool is capped at 200 text results. Change the text search phrase with `--query` or lower the per-pool cap with `--limit` (1-200). Each listing tab has a **Listed (UTC)** column showing when the listing was created on eBay (from the Browse API's `itemCreationDate`); it is left blank if eBay does not return a date.

## Daily underpriced listing email

The scan can send an HTML email containing only listings in the workbook's **Under Market** tab. Each entry includes its eBay link, the date and time it was listed (UTC), and an embedded listing photo when eBay provides an image URL. The daily scheduled run searches up to 200 graded and 200 raw listings, then sends the results after the workbook is saved. It sends a short no-results message when none qualify; **Regular** and **Needs Review** rows are never emailed.

1. Add SMTP settings and the destination address to `.env` using `.env.example` as a guide. `SMTP_HOST`, `SMTP_FROM`, and `EMAIL_TO` are required. Add username/password if your provider requires SMTP authentication. Keep `.env` private.
2. Run a one-time manual send with `./.venv/bin/python ebay_mewtwo_finder.py --limit 200 --send-email`. This sends a real email; use a test recipient first if you are validating the setup.
3. To schedule it on this Mac, make sure macOS is set to the **America/Denver** time zone, then install and load the included LaunchAgent:

   ```sh
   mkdir -p ~/Library/LaunchAgents
   cp com.mewtwo.ebay-market-listings.plist ~/Library/LaunchAgents/
   launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.mewtwo.ebay-market-listings.plist
   ```

   It runs daily at 7:00 AM local time (Mountain Time when macOS uses America/Denver). Keep the Mac awake and online for the scan and email to run at that time. Logs go to `/tmp/mewtwo-ebay-market-listings.log` and `/tmp/mewtwo-ebay-market-listings.err`.

To unload the schedule, run `launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.mewtwo.ebay-market-listings.plist`. The schedule file uses this project's current absolute path; update its three project paths if you move the folder.

## Free local endpoint (temporary setup only)

`ebay_notification_receiver.py` can answer eBay's verification challenge on your computer. A free Cloudflare Quick Tunnel can give it a temporary public HTTPS URL without an account or domain. This is for setup/testing: the random URL changes when the tunnel stops, and eBay requires the endpoint to remain available for future notifications. A laptop that sleeps or goes offline is not a dependable always-on endpoint.

1. Install `cloudflared` (on macOS with Homebrew: `brew install cloudflared`).
2. Start a tunnel to the local receiver port: `cloudflared tunnel --url http://127.0.0.1:8787`.
3. Copy the `https://....trycloudflare.com` URL printed by the tunnel. Use that URL plus `/ebay/account-deletion` as the Notification Endpoint URL.
4. In `.env`, set `EBAY_NOTIFICATION_TOKEN` to the same verification token entered in eBay, and `EBAY_NOTIFICATION_ENDPOINT` to the exact full URL you entered there.
5. In a second terminal, run `.venv/bin/python ebay_notification_receiver.py` from this folder, then save the endpoint in eBay.

Keep both terminal processes running while eBay verifies the endpoint. If the tunnel stops, its URL is temporary; you would need to create a new one and update the eBay endpoint. eBay recommends verifying notification signatures and requires endpoints to acknowledge notifications promptly. For a long-running compliant setup, use a reliable hosted receiver or request the exemption if the application qualifies.

## Hosted Worker

A stable Cloudflare Worker is deployed at `https://mewtwo-ebay-notifications.nnrotello.workers.dev/ebay/account-deletion`. The endpoint challenge was checked successfully. Its verification token is stored as a Cloudflare secret. In the eBay developer portal, save this exact URL as the Marketplace Account Deletion endpoint and use **Send Test Notification** to complete the subscription check. The Worker acknowledges POST notifications without storing their body; eBay recommends validating each `X-EBAY-SIGNATURE` before processing notifications.

## Example benchmark CSV

```csv
card_type,grade,benchmark_price,currency,source,as_of,minimum_price
raw,,45.00,USD,Recent sold comps; median of 10,2026-11-22,
graded,PSA 8,95.00,USD,Recent sold comps; median of 8,2026-11-22,
graded,PSA 9,180.00,USD,Recent sold comps; median of 8,2026-11-22,
```

The numbers above are examples of formatting only, not market estimates.
