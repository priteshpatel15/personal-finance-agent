#!/usr/bin/env python3
from __future__ import annotations

import argparse
import asyncio
from pathlib import Path


async def main() -> int:
    parser = argparse.ArgumentParser(description="Render-check a generated budget HTML report with Playwright.")
    parser.add_argument("html_report", type=Path)
    args = parser.parse_args()
    try:
        from playwright.async_api import async_playwright
    except Exception as exc:  # pragma: no cover - depends on optional local browser tooling
        print(f"Playwright unavailable: {exc}")
        return 2

    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch()
        page = await browser.new_page(viewport={"width": 1280, "height": 900})
        await page.goto(args.html_report.resolve().as_uri())
        await page.locator("h1").filter(has_text="Budget Report").wait_for()
        table_rows = await page.locator("tbody tr").count()
        if table_rows < 20:
            raise AssertionError(f"Expected at least 20 category rows, saw {table_rows}")
        await browser.close()
    print("Playwright report QA passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
