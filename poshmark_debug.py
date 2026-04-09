#!/usr/bin/env python3
"""
Quick diagnostic: fetch one Poshmark search page and dump what we can find.
Run: python3 poshmark_debug.py
"""
import json, re, sys
import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

def probe(url, label):
    print(f"\n{'='*60}")
    print(f"Probing: {label}")
    print(f"URL: {url}")
    s = requests.Session()
    s.headers.update(HEADERS)
    try:
        r = s.get(url, timeout=15)
        print(f"Status: {r.status_code}  |  Content-Type: {r.headers.get('Content-Type','?')}")
        print(f"Body length: {len(r.text)} chars")

        # 1. Look for __NEXT_DATA__
        soup = BeautifulSoup(r.text, "html.parser")
        tag = soup.find("script", {"id": "__NEXT_DATA__"})
        if tag and tag.string:
            try:
                data = json.loads(tag.string)
                pp = data.get("props", {}).get("pageProps", {})
                print(f"\n__NEXT_DATA__ found. Top-level pageProps keys: {list(pp.keys())[:20]}")
                # Try to find any list with items
                def find_lists(obj, path="", depth=0):
                    if depth > 6:
                        return
                    if isinstance(obj, list) and len(obj) > 0:
                        print(f"  List at '{path}': {len(obj)} items, first keys: {list(obj[0].keys())[:8] if isinstance(obj[0], dict) else type(obj[0]).__name__}")
                    elif isinstance(obj, dict):
                        for k, v in obj.items():
                            find_lists(v, f"{path}.{k}", depth+1)
                find_lists(pp)
            except json.JSONDecodeError as e:
                print(f"  JSON parse error: {e}")
        else:
            print("\nNo __NEXT_DATA__ script tag found.")

        # 2. Look for any script tags with listing/post data
        scripts = soup.find_all("script")
        print(f"\nTotal <script> tags: {len(scripts)}")
        for sc in scripts:
            txt = sc.string or ""
            if "listing" in txt.lower() and len(txt) > 200:
                print(f"  Script with 'listing': {len(txt)} chars, snippet: {txt[:120].strip()!r}")
                break

        # 3. Poshmark's internal catalog API
        api_url = (
            "https://poshmark.com/vm-rest/posts"
            "?request[search][query]=Brioni%20suit"
            "&request[browse][department]=Men"
            "&count=10"
        )
        print(f"\nTrying internal API: {api_url[:80]}...")
        r2 = s.get(api_url, timeout=15)
        print(f"API status: {r2.status_code}")
        if r2.status_code == 200:
            try:
                j = r2.json()
                print(f"API JSON keys: {list(j.keys())[:10]}")
                if "data" in j:
                    d = j["data"]
                    print(f"  data type: {type(d).__name__}")
                    if isinstance(d, list):
                        print(f"  data length: {len(d)}")
                        if d:
                            print(f"  first item keys: {list(d[0].keys())[:10]}")
            except Exception as e:
                print(f"  API JSON parse error: {e}")
                print(f"  Raw: {r2.text[:300]}")

    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    probe(
        "https://poshmark.com/search?q=Brioni%20suit&department=Men",
        "Poshmark search page (HTML)"
    )
