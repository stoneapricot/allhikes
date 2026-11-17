"""Sort each club's Telegram links into person / group / channel / bot, using Telegram's public
t.me pages, and save the best one people can actually message as `chat` in data/clubs.json.

A person's account is best (that's the club's rep), then a group chat. Channels only let people
read, so they're never used for the "Message the club" button.
Run it again after adding or changing Telegram links: python scripts/check_telegram.py
"""
import json, re, time, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CLUBS = ROOT / "data/clubs.json"


def kind(link):
    name = link.split("t.me/")[-1].strip("/").split("?")[0]
    req = urllib.request.Request(f"https://t.me/{name}", headers={"User-Agent": "Mozilla/5.0"})
    page = urllib.request.urlopen(req, timeout=20).read().decode("utf8", "replace")
    extra = (re.search(r'tgme_page_extra">\s*([^<]*)', page) or [None, ""])[1].strip()
    button = (re.search(r'tgme_action_button_new[^>]*>([^<]*)', page) or [None, ""])[1].strip()
    if "subscriber" in extra: return "channel"
    if "member" in extra: return "group"
    if "Start Bot" in button or name.lower().endswith("bot"): return "bot"
    if button == "Send Message" or extra.startswith("@"): return "person"
    return "unknown"


clubs = json.loads(CLUBS.read_text(encoding="utf8"))
rank = {"person": 0, "group": 1}
for c in clubs:
    kinds = {}
    for link in c.get("telegram", []):
        try:
            kinds[link] = kind(link)
        except Exception as e:
            kinds[link] = "unknown"
            print("  could not check", link, e)
        time.sleep(0.5)
    c["telegram_kinds"] = kinds
    chats = sorted((rank[k], l) for l, k in kinds.items() if k in rank)
    c["chat"] = chats[0][1] if chats else ""
    if kinds:
        print(f'{c["name"]}: ' + ", ".join(f"{l} = {k}" for l, k in kinds.items()) + (f'  -> {c["chat"]}' if c["chat"] else "  -> no chat"))

CLUBS.write_text(json.dumps(clubs, ensure_ascii=False, indent=1) + "\n", encoding="utf8")
