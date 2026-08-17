#!/usr/bin/env python3
"""fetch_emails.py — Fetch Outlook research emails (no link crawling).

This is the GENERIC, reusable version. It contains NO hardcoded fund name,
NO hardcoded file paths, and NO credentials. You supply:
  - the target date          (--date YYYY-MM-DD, default: yesterday)
  - the Outlook inbox URL    (--url, or OUTLOOK_URL env, default below)
  - the persistent Edge profile dir containing your saved login session
                             (--profile, or OUTLOOK_PROFILE_DIR env;
                              default: ./.edge_profile — gitignored, never commit)

The persistent profile is created the FIRST time you log in manually. After that,
the script reuses the saved session (no password prompt). Never commit the
profile directory.

Output: a JSON file (today_emails.json or emails_<date>.json) with one entry per
email: {subject, sender, time, body, body_len, report_links_crawled:false}.
"""
import asyncio, json, re, os, sys, argparse
from pathlib import Path
from datetime import datetime, timezone, timedelta
from urllib.parse import urljoin

CST = timezone(timedelta(hours=8))
DEFAULT_INBOX = "https://outlook.office.com/mail/inbox"

TARGET_DATE_STR = None


def _resolve_target_date():
    if TARGET_DATE_STR:
        try:
            return datetime.strptime(TARGET_DATE_STR, "%Y-%m-%d").date()
        except ValueError:
            pass
    return (datetime.now(CST) - timedelta(days=1)).date()


def _output_json_for(date_obj, out_dir):
    yesterday = (datetime.now(CST) - timedelta(days=1)).date()
    if date_obj == yesterday:
        return out_dir / "today_emails.json"
    return out_dir / f"emails_{date_obj.isoformat()}.json"


def _app_root():
    if getattr(sys, "frozen", False):
        root = Path(sys.executable).resolve().parent
    else:
        root = Path(__file__).resolve().parent
    if root.name == "_internal":
        root = root.parent
    return root


APP_ROOT = _app_root()
PROFILE_DIR_ARG = os.environ.get("OUTLOOK_PROFILE_DIR")
INBOX_URL = os.environ.get("OUTLOOK_URL", DEFAULT_INBOX)
OUTPUT_DIR = APP_ROOT.parent / "reports"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

try:
    import sys as _s
    _s.stdout.reconfigure(encoding="utf-8")
    _s.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

L = lambda m: print(m, flush=True)


async def run(profile_dir: Path):
    lk = profile_dir / "SingletonLock"
    if lk.exists():
        lk.unlink(missing_ok=True)

    async with async_playwright() as p:
        L("Launching (nolinks)...")
        L(f"Using profile: {profile_dir}")
        b = await p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir), channel="msedge", headless=False,
            args=["--start-maximized"], viewport={"width": 1440, "height": 900},
        )
        pg = b.pages[0] if b.pages else await b.new_page()
        emails = []

        try:
            await pg.goto(INBOX_URL, wait_until="domcontentloaded")
            await pg.wait_for_timeout(8000)
            if any(k in pg.url.lower() for k in ("login", "oauth")):
                L("Please log in (first run only)...")
                for _ in range(300):
                    await pg.wait_for_timeout(2000)
                    if not any(k in pg.url.lower() for k in ("login", "oauth")):
                        break
            L("Logged in"); await pg.wait_for_timeout(3000)

            try:
                ib = pg.locator('[title*="Inbox"], [title*="收件箱"]').first
                if await ib.count() > 0:
                    await ib.click()
                    await pg.wait_for_timeout(2000)
            except Exception:
                pass

            target_date = _resolve_target_date()
            date_str = target_date.strftime("%Y-%m-%d")
            L(f"Searching: received:{date_str}")
            try:
                search_input = pg.locator('input[aria-label*="Search"], [role="searchbox"]').first
                if await search_input.count() > 0:
                    await search_input.click()
                    await pg.wait_for_timeout(500)
                    await search_input.fill("")
                    await search_input.type(f"received:{date_str}", delay=80)
                    await pg.keyboard.press("Enter")
                    await pg.wait_for_timeout(4000)
                    L("  Search complete")
            except Exception as e:
                L(f"  Search failed ({e})")

            L("Preloading: scroll to render all results...")
            prev_opt_count = 0
            for rnd in range(40):
                await pg.evaluate("""() => {
                    const lb = document.querySelector('[role="listbox"]');
                    if (!lb) return 0;
                    const before = lb.scrollTop;
                    lb.scrollTop = lb.scrollHeight;
                    return lb.scrollTop - before;
                }""")
                await pg.wait_for_timeout(400)
                if rnd >= 5:
                    curr = await pg.locator('[role="option"]').count()
                    if curr == prev_opt_count:
                        L(f"  Preload done: {curr} options")
                        break
                    prev_opt_count = curr

            await pg.evaluate("""() => {
                const lb = document.querySelector('[role="listbox"]');
                if (lb) lb.scrollTop = 0;
            }""")
            await pg.wait_for_timeout(1500)

            opts = pg.locator('[role="option"]')
            oc = await opts.count()
            L(f"Visible options: {oc}")
            if oc < 1:
                L("Empty list"); return 0

            first_idx = 0
            for i in range(oc):
                aria = await opts.nth(i).get_attribute("aria-label") or ""
                if len(aria) > 15 and "Navigation" not in aria and "Folder" not in aria:
                    first_idx = i
                    break

            await opts.nth(first_idx).click()
            await pg.wait_for_timeout(2000)
            L("First email selected; navigating one by one (skip link tracking)...")

            body_sels = ['[role="main"]', '[aria-label="Message body"]',
                         'div[role="document"]', '.allowTextSelection', 'div[role="textbox"]']
            seen_subjects = set()
            no_new_count = 0

            for i in range(500):
                if i > 0:
                    await pg.keyboard.press("ArrowDown")
                    await pg.wait_for_timeout(800)

                sj_fast = ""
                try:
                    idx = first_idx + (i % max(oc, 1))
                    aria = await pg.locator('[role="option"]').nth(idx).get_attribute("aria-label") or ""
                    parts = aria.split(" ")
                    sp = []
                    sk = {"Unread", "Read", "Collapsed", "Expanded", "Has", "attachments", ""}
                    for part in parts:
                        if part in sk:
                            continue
                        if re.match(r'^\d{1,2}:\d{2}', part):
                            break
                        if len(part) > 1:
                            sp.append(part)
                    sj_fast = " ".join(sp) if sp else aria[:80]
                except Exception:
                    pass
                if not sj_fast:
                    sj_fast = "(no subject)"

                dedup = sj_fast[:60]
                if dedup in seen_subjects:
                    no_new_count += 1
                    if no_new_count >= 3:
                        scrolled = await pg.evaluate("""() => {
                            const lb = document.querySelector('[role="listbox"]');
                            if (!lb) return 0;
                            const before = lb.scrollTop;
                            lb.scrollTop += lb.clientHeight;
                            return lb.scrollTop - before;
                        }""")
                        if scrolled > 10:
                            L("    Duplicates found, scrolling down to load more...")
                            await pg.wait_for_timeout(800)
                    if no_new_count >= 15:
                        L(f"{no_new_count} consecutive duplicates ({len(emails)} total); list exhausted")
                        break
                    continue

                body = ""
                for s in body_sels:
                    try:
                        body = await pg.evaluate(f"""
                            (sel) => {{
                                const el = document.querySelector(sel);
                                if (!el) return '';
                                const t = el.textContent || '';
                                return t.length > 12000 ? t.substring(0, 12000) + '...[truncated]' : t;
                            }}
                        """, s)
                        if body and len(body.strip()) > 25:
                            break
                    except Exception:
                        pass

                if not body:
                    await pg.wait_for_timeout(800)
                    for s in body_sels:
                        try:
                            body = await pg.evaluate(f"""
                                (sel) => {{
                                    const el = document.querySelector(sel);
                                    if (!el) return '';
                                    const t = el.textContent || '';
                                    return t.length > 12000 ? t.substring(0, 12000) + '...[truncated]' : t;
                                }}
                            """, s)
                            if body and len(body.strip()) > 25:
                                break
                        except Exception:
                            pass

                sd = ""; ts = ""; sj = ""
                if body:
                    for ln in body.split("\n"):
                        ln = ln.strip()
                        if ln.startswith("From:"):
                            sd = ln[5:].strip()
                        elif ln.startswith("Subject:"):
                            sj = ln[8:].strip()
                        elif ln.startswith("Sent:"):
                            ts = ln[5:].strip()
                if not sj:
                    sj = sj_fast

                seen_subjects.add(dedup)
                no_new_count = 0

                # Skip link tracking — only save email body
                emails.append({"subject": sj, "sender": sd, "time": ts,
                               "body": body[:8000], "body_len": len(body),
                               "report_links_crawled": False})
                L(f"  [{len(emails)}] {sj[:80]}")

            L(f"\nDone: {len(emails)} emails (nolinks)")

        except Exception as e:
            L(f"Error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            try:
                await b.close()
            except Exception:
                pass
            L("Browser closed")

    if emails:
        out_path = _output_json_for(target_date, OUTPUT_DIR)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(emails, f, ensure_ascii=False, indent=2)
        L(f"Saved: {out_path.name}")
    return len(emails)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch Outlook research emails (no link crawling).")
    parser.add_argument("--date", default=None,
                        help="Target date YYYY-MM-DD for received: search (default: yesterday).")
    parser.add_argument("--profile", default=PROFILE_DIR_ARG,
                        help="Path to the persistent Edge profile dir with your saved login session.")
    parser.add_argument("--url", default=INBOX_URL,
                        help="Outlook inbox URL.")
    args = parser.parse_args()

    if args.date:
        TARGET_DATE_STR = args.date
    if args.url:
        INBOX_URL = args.url

    profile_dir = Path(args.profile) if args.profile else (APP_ROOT.parent / ".edge_profile")
    if not profile_dir.exists():
        L(f"Profile not found: {profile_dir}")
        L("Create it once by logging into Outlook with Edge, then re-run.")
        sys.exit(2)

    try:
        from playwright.async_api import async_playwright
    except ImportError:
        L("playwright not installed. Install it in your Python env: pip install playwright && playwright install msedge")
        sys.exit(3)

    c = asyncio.run(run(profile_dir))
    L(f"TOTAL={c}")
