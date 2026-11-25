#!/usr/bin/env python3
"""Find Pokémon eBay listings and compare them with user-supplied benchmarks."""
from __future__ import annotations

import argparse
import base64
import csv
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import re
import smtplib
import sys
from datetime import datetime, timezone
from email.message import EmailMessage
from email.utils import formatdate
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

ROOT = Path(__file__).resolve().parent
API_ROOT = "https://api.ebay.com"


def get_access_token(client_id: str, client_secret: str, api_root: str) -> str:
    encoded = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    response = requests.post(
        f"{api_root}/identity/v1/oauth2/token",
        headers={"Authorization": f"Basic {encoded}", "Content-Type": "application/x-www-form-urlencoded"},
        data={"grant_type": "client_credentials", "scope": "https://api.ebay.com/oauth/api_scope"},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def load_benchmarks(path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    benchmarks: dict[tuple[str, str], dict[str, Any]] = {}
    with path.open(newline="", encoding="utf-8-sig") as stream:
        for row in csv.DictReader(stream):
            card_type = (row.get("card_type") or "").strip().lower()
            grade = normalize_grade(row.get("grade") or "") if card_type == "graded" else ""
            raw_price = (row.get("benchmark_price") or "").strip()
            if card_type not in {"raw", "graded"} or not raw_price:
                continue
            try:
                price = float(raw_price.replace(",", ""))
            except ValueError as exc:
                raise ValueError(f"Invalid benchmark price: {raw_price!r}") from exc
            raw_minimum = (row.get("minimum_price") or "").strip()
            try:
                minimum_price = float(raw_minimum.replace(",", "")) if raw_minimum else None
                if minimum_price is not None and minimum_price < 0:
                    raise ValueError
            except ValueError as exc:
                raise ValueError(f"Invalid minimum price: {raw_minimum!r}") from exc
            benchmarks[(card_type, grade)] = {
                "price": price,
                "minimum_price": minimum_price,
                "currency": (row.get("currency") or "USD").strip().upper(),
                "source": (row.get("source") or "").strip(),
                "as_of": (row.get("as_of") or "").strip(),
            }
    return benchmarks


def normalize_grade(value: str) -> str:
    text = re.sub(r"\s+", " ", value.strip().upper())
    text = re.sub(r"^(PSA|BGS|CGC|SGC|BECKETT)\s*", lambda m: "BGS " if m.group(1) == "BECKETT" else m.group(1) + " ", text)
    return text.strip()


def listing_price(item: dict[str, Any]) -> tuple[float | None, str, float | None]:
    price_obj = item.get("price") or {}
    price = float(price_obj["value"]) if price_obj.get("value") is not None else None
    currency = (price_obj.get("currency") or "USD").upper()
    shipping: float | None = 0.0
    options = item.get("shippingOptions") or []
    if options:
        cost = (options[0].get("shippingCost") or {}).get("value")
        shipping = float(cost) if cost is not None else None
    total = price + shipping if price is not None and shipping is not None else None
    return total, currency, shipping


def search_listings(token: str, query: str, limit: int, condition_id: str, marketplace: str, api_root: str, zip_code: str = "") -> list[dict[str, Any]]:
    headers = {"Authorization": f"Bearer {token}", "X-EBAY-C-MARKETPLACE-ID": marketplace, "Accept": "application/json"}
    if zip_code:
        location = quote(f"country=US,zip={zip_code}", safe="")
        headers["X-EBAY-C-ENDUSERCTX"] = f"contextualLocation={location}"
    items: list[dict[str, Any]] = []
    offset = 0
    while len(items) < limit:
        page_size = min(50, limit - len(items))
        response = requests.get(
            f"{api_root}/buy/browse/v1/item_summary/search",
            headers=headers,
            params={"q": query, "limit": page_size, "offset": offset,
                    "sort": "newlyListed",
                    "filter": f"conditionIds:{{{condition_id}}}"},
            timeout=30,
        )
        response.raise_for_status()
        page = response.json()
        batch = page.get("itemSummaries", [])
        items.extend(batch)
        if not batch or len(items) >= int(page.get("total", 0)):
            break
        offset += len(batch)
    return items[:limit]


def enrich_calculated_shipping(items: list[dict[str, Any]], token: str, marketplace: str, api_root: str, zip_code: str = "") -> int:
    """Fetch full item details when a matching Mewtwo summary omits calculated shipping."""
    headers = {"Authorization": f"Bearer {token}", "X-EBAY-C-MARKETPLACE-ID": marketplace, "Accept": "application/json"}
    if zip_code:
        location = quote(f"country=US,zip={zip_code}", safe="")
        headers["X-EBAY-C-ENDUSERCTX"] = f"contextualLocation={location}"
    enriched = 0
    for item in items:
        options = item.get("shippingOptions") or []
        missing_calculated_cost = any(
            option.get("shippingCostType") == "CALCULATED" and not (option.get("shippingCost") or {}).get("value")
            for option in options
        )
        if not missing_calculated_cost:
            continue
        item_id = item.get("itemId")
        if not item_id:
            continue
        response = requests.get(
            f"{api_root}/buy/browse/v1/item/{quote(item_id, safe='')}",
            headers=headers, timeout=30,
        )
        response.raise_for_status()
        details = response.json()
        item.update({k: v for k, v in details.items() if k not in item or not item[k]})
        # Prefer the detailed location-specific quote over the summary's unpriced CALCULATED option.
        if details.get("shippingOptions"):
            item["shippingOptions"] = details["shippingOptions"]
        enriched += 1
    return enriched


def deduplicate_listings(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deduplicate filtered text results by eBay item ID."""
    merged: dict[str, dict[str, Any]] = {}
    for item in items:
        item_id = item.get("itemId") or item.get("itemWebUrl") or item.get("title", "")
        if item_id not in merged:
            merged[item_id] = dict(item)
    return list(merged.values())


def condition_descriptor_values(item: dict[str, Any], descriptor_name: str) -> list[str]:
    wanted = re.sub(r"[^a-z]+", "", descriptor_name.lower())
    found: list[str] = []
    for descriptor in item.get("_condition_descriptors") or item.get("conditionDescriptors") or []:
        name = re.sub(r"[^a-z]+", "", str(descriptor.get("name", "")).lower())
        if name != wanted:
            continue
        for value in descriptor.get("values") or []:
            content = value.get("content")
            if content:
                found.append(str(content))
    return found


def structured_card_type_and_grade(item: dict[str, Any]) -> tuple[str | None, str]:
    condition_id = str(item.get("conditionId") or "")
    condition = str(item.get("condition") or "").strip().lower()
    grader_values = condition_descriptor_values(item, "Professional Grader")
    grade_values = condition_descriptor_values(item, "Grade")
    if grader_values or grade_values:
        grader_text = " ".join(grader_values).upper()
        grader = None
        for label, pattern in (("PSA", r"\bPSA\b|PROFESSIONAL SPORTS AUTHENTICATOR"),
                               ("BGS", r"\bBGS\b|BECKETT"),
                               ("CGC", r"\bCGC\b|CERTIFIED GUARANTY"),
                               ("SGC", r"\bSGC\b|SPORTSCARD GUARANTY"),
                               ("GG", r"\bGG\b|GET\s*GRADED")):
            if re.search(pattern, grader_text):
                grader = label
                break
        number = re.search(r"\b([0-9]{1,2}(?:\.\d)?)\b", grade_values[0])
        if number:
            # eBay's structured grading descriptors are authoritative. Keep
            # unrecognized providers as their own label (commonly "Other").
            if not grader:
                grader = "UNKNOWN" if not grader_text else grader_text.title()
            return "graded", normalize_grade(f"{grader} {number.group(1)}")
        return "graded", ""
    if condition_id == "2750" or condition == "graded":
        return "graded", ""
    if condition_id == "4000" or condition == "ungraded":
        return "raw", ""
    return None, ""


def enrich_item_details(items: list[dict[str, Any]], token: str, marketplace: str, api_root: str, zip_code: str = "") -> tuple[int, int]:
    """Fetch listing characteristics for search results, without inspecting titles."""
    headers = {"Authorization": f"Bearer {token}", "X-EBAY-C-MARKETPLACE-ID": marketplace, "Accept": "application/json"}
    if zip_code:
        location = quote(f"country=US,zip={zip_code}", safe="")
        headers["X-EBAY-C-ENDUSERCTX"] = f"contextualLocation={location}"
    candidates_by_id: dict[str, list[dict[str, Any]]] = {}
    for item in items:
        item_id = item.get("itemId")
        if item_id:
            candidates_by_id.setdefault(item_id, []).append(item)

    def fetch(item_id: str) -> tuple[str, dict[str, Any] | None]:
        response = requests.get(
            f"{api_root}/buy/browse/v1/item/{quote(item_id, safe='')}",
            headers=headers, timeout=30,
        )
        if response.status_code in {404, 410}:
            return item_id, None
        response.raise_for_status()
        return item_id, response.json()

    fetched = failed = 0
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = [pool.submit(fetch, item_id) for item_id in candidates_by_id]
        for future in as_completed(futures):
            try:
                item_id, details = future.result()
            except requests.HTTPError as exc:
                # A rate limit makes the whole scan incomplete; don't silently
                # treat missing descriptors as condition/grade filter failures.
                if exc.response is not None and exc.response.status_code == 429:
                    raise
                failed += 1
                continue
            except requests.RequestException:
                failed += 1
                continue
            if not details:
                failed += 1
                continue
            fetched += 1
            for item in candidates_by_id[item_id]:
                item["_condition_descriptors"] = details.get("conditionDescriptors") or []
                # Keep the compact search image as a fallback, but prefer the
                # full item response's image set for email and workbook output.
                if item.get("image"):
                    item["_summary_image"] = item["image"]
                for key in ("condition", "conditionId", "price", "shippingOptions", "image", "additionalImages"):
                    if details.get(key):
                        item[key] = details[key]
    return fetched, failed


def filter_by_quality(items: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Keep raw NM/LP and PSA 7+; review graded listings with incomplete attributes."""
    included: list[dict[str, Any]] = []
    stats = {"Raw excluded (condition not NM/LP or missing)": 0,
             "Graded excluded (not PSA or below 7)": 0,
             "Graded sent to review (missing grader/grade characteristic)": 0,
             "Unclassified sent to review (missing eBay condition)": 0}
    for item in items:
        card_type, grade = structured_card_type_and_grade(item)
        if card_type == "graded":
            match = re.search(r"([0-9]{1,2}(?:\.\d)?)\s*$", grade)
            if match and grade.startswith("PSA ") and float(match.group(1)) >= 7:
                item["_effective_grade"] = grade
                included.append(item)
            elif not match or grade.startswith("UNKNOWN "):
                stats["Graded sent to review (missing grader/grade characteristic)"] += 1
                included.append(item)
            else:
                stats["Graded excluded (not PSA or below 7)"] += 1
        elif card_type == "raw":
            values = condition_descriptor_values(item, "Card Condition")
            condition = " ".join(values)
            if re.search(r"\b(?:near\s+mint|nm|lightly\s+played|lp)\b", condition, re.I):
                item["_effective_condition"] = condition
                included.append(item)
            else:
                stats["Raw excluded (condition not NM/LP or missing)"] += 1
        else:
            stats["Unclassified sent to review (missing eBay condition)"] += 1
            included.append(item)
    return included, stats


def make_rows(items: list[dict[str, Any]], benchmarks: dict[tuple[str, str], dict[str, Any]]) -> dict[str, list[list[Any]]]:
    tabs: dict[str, list[list[Any]]] = {"Under Market": [], "Regular": [], "Needs Review": []}
    for item in items:
        title = item.get("title", "")
        card_type, grade = structured_card_type_and_grade(item)
        issue = ""
        if card_type == "graded" and (not grade or grade.startswith("UNKNOWN ")):
            issue = "eBay identifies this as graded, but its structured grader or grade characteristic is missing."
        elif card_type is None:
            issue = "Could not determine graded or ungraded status from eBay listing characteristics."
        total, currency, shipping = listing_price(item)
        benchmark = benchmarks.get((card_type, grade)) if card_type else None
        if not issue and benchmark is None:
            issue = f"No benchmark row for {grade or 'raw'} in benchmarks.csv."
        elif not issue and benchmark and currency != benchmark["currency"]:
            issue = f"Currency mismatch: listing {currency}, benchmark {benchmark['currency']}."
        if not issue and total is None:
            issue = "Could not calculate total because listing price or shipping cost is missing."
        minimum_price = benchmark.get("minimum_price") if benchmark else None
        if not issue and minimum_price is not None and total is not None and total < minimum_price:
            issue = f"Below minimum price {currency} {minimum_price:.2f}; possible counterfeit or pricing error. Verify before considering."

        row = [
            title,
            card_type or "Review",
            grade,
            round(total, 2) if total is not None else "",
            currency,
            round(shipping, 2) if shipping is not None else "Unknown",
            round(float(benchmark["price"]), 2) if benchmark else "",
            benchmark["source"] if benchmark else "",
            benchmark["as_of"] if benchmark else "",
            round(total - float(benchmark["price"]), 2) if total is not None and benchmark else "",
            item.get("_effective_condition") or item.get("condition", ""),
            item.get("itemWebUrl", ""),
            issue,
            listing_image_url(item),
            round(minimum_price, 2) if minimum_price is not None else "",
        ]
        if issue:
            tabs["Needs Review"].append(row)
        elif total < float(benchmark["price"]):
            tabs["Under Market"].append(row)
        else:
            tabs["Regular"].append(row)
    return tabs


def listing_image_url(item: dict[str, Any]) -> str:
    """Use the listing's primary gallery image, never a later photo by size."""
    primary_candidates = [item.get("image"), item.get("_summary_image")]
    for image in primary_candidates:
        if isinstance(image, dict) and image.get("imageUrl"):
            return str(image["imageUrl"])

    # The API's primary image can be absent. Only then use the first available
    # additional listing image, followed by the search thumbnail.
    for image in (item.get("additionalImages") or []) + (item.get("thumbnailImages") or []):
        if isinstance(image, dict) and image.get("imageUrl"):
            return str(image["imageUrl"])
    return ""


HEADERS = ["Title", "Card Type", "Grade", "Total Price", "Currency", "Shipping", "Benchmark", "Benchmark Source", "Benchmark As Of", "Difference", "Condition", "eBay URL", "Review Note", "Image URL", "Minimum Price"]


def send_underpriced_email(rows: list[list[Any]], recipient: str, query: str) -> None:
    """Email only Under Market rows, embedding each eBay listing photo when available."""
    required = ["SMTP_HOST", "SMTP_FROM"]
    missing = [name for name in required if not os.getenv(name)]
    if not recipient:
        missing.append("EMAIL_TO or --email-to")
    if missing:
        raise ValueError("Email is enabled but configuration is missing: " + ", ".join(missing))

    message = EmailMessage()
    message["Subject"] = f"Underpriced Pokémon eBay listings: {query} ({len(rows)})"
    message["From"] = os.environ["SMTP_FROM"]
    message["To"] = recipient
    message["Date"] = formatdate(localtime=True)
    if not rows:
        message.set_content(f"No underpriced Pokémon listings were found for '{query}' in today's scan.")
        message.add_alternative(f"<p>No underpriced Pokémon listings were found for '{html_escape(query)}' in today's scan.</p>", subtype="html")
    else:
        html = [f"<html><body><h2>Underpriced Pokémon listings: {html_escape(query)}</h2>"]
        html.append("<p>These listings passed the project filters and are priced below their configured benchmark.</p>")
        embedded_images: list[tuple[bytes, str, str]] = []
        inline_image_bytes = 0
        inline_image_budget = 16 * 1024 * 1024
        for index, row in enumerate(rows):
            title, _, grade, total, currency, _, benchmark, *_ = row
            url = row[11]
            image_url = row[13]
            html.append("<hr>")
            html.append(f"<h3><a href=\"{html_escape(url)}\">{html_escape(title)}</a></h3>")
            html.append(f"<p>{html_escape(grade or 'Raw')} · {html_escape(currency)} {total:.2f} (benchmark {currency} {benchmark:.2f})</p>")
            if image_url:
                try:
                    response = requests.get(image_url, timeout=20)
                    response.raise_for_status()
                    with Image.open(BytesIO(response.content)) as image:
                        image.thumbnail((700, 700))
                        if image.mode not in {"RGB", "L"}:
                            image = image.convert("RGB")
                        output = BytesIO()
                        image.save(output, format="JPEG", quality=68, optimize=True, progressive=True)
                        content = output.getvalue()
                    cid = f"listing-{index}@mewtwo-finder"
                    if inline_image_bytes + len(content) <= inline_image_budget:
                        html.append(f'<p><img src="cid:{cid}" alt="eBay listing photo" style="max-width:700px;height:auto"></p>')
                        embedded_images.append((content, "jpeg", cid))
                        inline_image_bytes += len(content)
                    else:
                        # Keep every listing pictured without exceeding Gmail's
                        # attachment limit when many listings qualify.
                        html.append(f'<p><img src="{html_escape(image_url)}" alt="eBay listing photo" style="max-width:700px;height:auto"></p>')
                except (requests.RequestException, OSError, ValueError, TypeError):
                    html.append(f'<p><img src="{html_escape(image_url)}" alt="eBay listing photo" style="max-width:700px;height:auto"></p>')
            html.append(f'<p><a href="{html_escape(url)}">View listing on eBay</a></p>')
        html.append("</body></html>")
        message.set_content(f"{len(rows)} underpriced Pokémon listing(s) found for '{query}'. Open this email in an HTML-capable mail client to view photos and links.")
        message.add_alternative("".join(html), subtype="html")
        html_part = message.get_body(preferencelist=("html",))
        if html_part is None:
            raise ValueError("Could not create the HTML email body")
        for content, subtype, cid in embedded_images:
            html_part.add_related(content, maintype="image", subtype=subtype, cid=f"<{cid}>")

    host = os.environ["SMTP_HOST"]
    port = int(os.getenv("SMTP_PORT", "587"))
    # Gmail requires the full account address as the SMTP username. Falling
    # back to SMTP_FROM also supports a minimal .env with only an app password.
    username = os.getenv("SMTP_USERNAME", "").strip() or os.environ["SMTP_FROM"].strip()
    password = os.getenv("SMTP_PASSWORD", "")
    if os.getenv("SMTP_USE_SSL", "false").strip().lower() in {"1", "true", "yes"}:
        with smtplib.SMTP_SSL(host, port, timeout=30) as server:
            if username:
                server.login(username, password)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.starttls()
            if username:
                server.login(username, password)
            server.send_message(message)


def html_escape(value: Any) -> str:
    import html
    return html.escape(str(value or ""), quote=True)


def write_workbook(tabs: dict[str, list[list[Any]]], output: Path, query: str,
                   filter_stats: dict[str, int], search_stats: dict[str, dict[str, int]]) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Pokémon eBay Listing Scan"])
    summary.append(["Search query", query])
    summary.append(["Text result sort", "Newly listed first"])
    summary.append(["Created (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")])
    summary.append(["Benchmark method", "User-supplied benchmarks.csv; see README for sold-data limitation"])
    summary.append([])
    summary.append(["Search pool", "Text results", "Passed attribute filters"])
    for name, counts in search_stats.items():
        summary.append([name, counts["text"], counts["combined"]])
    summary.append([])
    summary.append(["Tab", "Listings"])
    for name, rows in tabs.items():
        summary.append([name, len(rows)])
    summary.append([])
    summary.append(["Quality filter", "Count"])
    for label, count in filter_stats.items():
        summary.append([label, count])
    summary.freeze_panes = "A7"
    for name, rows in tabs.items():
        sheet = wb.create_sheet(name)
        sheet.append(HEADERS)
        for row in rows:
            sheet.append(row)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True)
        widths = [55, 14, 12, 14, 12, 12, 14, 32, 16, 14, 22, 48, 48, 48, 16]
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for row in sheet.iter_rows(min_row=2):
            row[11].hyperlink = row[11].value if row[11].value else None
            if row[11].hyperlink:
                row[11].style = "Hyperlink"
            row[0].alignment = Alignment(wrap_text=True, vertical="top")
            row[12].alignment = Alignment(wrap_text=True, vertical="top")
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def main() -> int:
    load_dotenv(ROOT / ".env")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--query", default="pikachu grey felt hat", help="eBay text search phrase")
    parser.add_argument("--limit", type=int, default=200, help="Maximum text results per card type (1-200)")
    parser.add_argument("--send-email", action="store_true", help="Email only Under Market listings using SMTP settings in .env")
    parser.add_argument("--email-to", default=os.getenv("EMAIL_TO", ""), help="Email recipient (or set EMAIL_TO in .env)")
    parser.add_argument("--benchmarks", type=Path, default=ROOT / "benchmarks.csv")
    parser.add_argument("--output", type=Path, default=ROOT / "pokemon_ebay_listings.xlsx")
    parser.add_argument("--marketplace", default=os.getenv("EBAY_MARKETPLACE_ID", "EBAY_US"))
    parser.add_argument("--environment", choices=["production", "sandbox"], default=os.getenv("EBAY_ENV", "production").lower())
    parser.add_argument("--zip", default=os.getenv("EBAY_ZIP_CODE", ""), help="US destination ZIP for calculated shipping")
    args = parser.parse_args()
    if not 1 <= args.limit <= 200:
        parser.error("--limit must be between 1 and 200")
    client_id, client_secret = os.getenv("EBAY_CLIENT_ID"), os.getenv("EBAY_CLIENT_SECRET")
    if not client_id or not client_secret:
        print("Missing EBAY_CLIENT_ID or EBAY_CLIENT_SECRET. Copy .env.example to .env and fill in your eBay app keys.", file=sys.stderr)
        return 2
    if not args.benchmarks.exists():
        print(f"Benchmark file not found: {args.benchmarks}", file=sys.stderr)
        return 2
    try:
        benchmarks = load_benchmarks(args.benchmarks)
        api_root = "https://api.sandbox.ebay.com" if args.environment == "sandbox" else API_ROOT
        token = get_access_token(client_id, client_secret, api_root)
        search_groups = {
            "Graded (PSA 7+)": "2750",
            "Raw (Near Mint / Lightly Played)": "4000",
        }
        listings: list[dict[str, Any]] = []
        search_stats: dict[str, dict[str, int]] = {}
        eligible_by_group: dict[str, list[dict[str, Any]]] = {}
        filter_stats: dict[str, int] = {}
        for group_name, condition_id in search_groups.items():
            text_items = search_listings(token, args.query, args.limit, condition_id,
                                         args.marketplace, api_root, args.zip)
            for item in text_items:
                item["_search_group"] = group_name
            listings.extend(text_items)
            search_stats[group_name] = {"text": len(text_items), "combined": 0}
        details_fetched, details_failed = enrich_item_details(listings, token, args.marketplace, api_root, args.zip)
        for group_name in search_groups:
            group_items = [item for item in listings if item.get("_search_group") == group_name]
            eligible_items, stats = filter_by_quality(group_items)
            eligible_by_group[group_name] = eligible_items
            for key, count in stats.items():
                filter_stats[key] = filter_stats.get(key, 0) + count
            search_stats[group_name]["combined"] = len(eligible_by_group[group_name])
        eligible_listings = deduplicate_listings([item for group_items in eligible_by_group.values() for item in group_items])
        tabs = make_rows(eligible_listings, benchmarks)
        write_workbook(tabs, args.output, args.query, filter_stats, search_stats)
        if args.send_email:
            send_underpriced_email(tabs["Under Market"], args.email_to, args.query)
    except requests.HTTPError as exc:
        detail = exc.response.text[:1000] if exc.response is not None else str(exc)
        print(f"eBay API request failed: {detail}", file=sys.stderr)
        return 1
    except (requests.RequestException, smtplib.SMTPException, OSError, ValueError, KeyError) as exc:
        print(f"Could not complete scan: {exc}", file=sys.stderr)
        return 1
    print(f"Scanned {len(listings)} text results; kept {len(eligible_listings)} unique listings after attribute filters and deduplication.")
    for name, counts in search_stats.items():
        print(f"{name}: {counts['text']} text results; {counts['combined']} passed characteristic filters.")
    print(f"Retrieved eBay item characteristics for {details_fetched} listings; detail lookups unavailable: {details_failed}.")
    print(f"Kept {len(eligible_listings)} listings after condition/grade filters; filtered out {sum(filter_stats.values())}.")
    for label, count in filter_stats.items():
        print(f"{label}: {count}")
    print(f"Under Market: {len(tabs['Under Market'])}; Regular: {len(tabs['Regular'])}; Needs Review: {len(tabs['Needs Review'])}")
    print(f"Workbook: {args.output.resolve()}")
    if args.send_email:
        print(f"Email sent to {args.email_to}; underpriced listings included: {len(tabs['Under Market'])}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
