"""Club contacts spreadsheet: outreach/club-contacts.xlsx, made from data/clubs.json.

Shows every club, all its contacts, what kind each Telegram link is, and which button the
website shows on that club's hikes ("Telegram chat" or "Instagram DM").
Run: python scripts/contacts_sheet.py   (needs: python -m pip install openpyxl)
"""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
clubs = json.loads((ROOT / "data/clubs.json").read_text(encoding="utf8"))
hikes = json.loads((ROOT / "assets/data.js").read_text(encoding="utf8")[len("window.AH="):-2])["hikes"]
slug = lambda ig: ig.lower().replace("_", "-").replace(".", "-")

wb = Workbook()
ws = wb.active
ws.title = "Clubs"
head = ["Club", "Type", "Active on site", "Upcoming hikes", "Button on the website", "Button goes to",
        "Instagram", "Telegram links (kind)", "WhatsApp", "Phone", "Email", "Website", "Followers"]
ws.append(head)
GREEN, AMBER, GREY = "D8EBDD", "FBE7C6", "EEEEEE"
for c in sorted(clubs, key=lambda c: (not c.get("chat"), c["name"].lower())):
    kinds = c.get("telegram_kinds", {})
    chat = c.get("chat", "")
    if chat:
        button = "Telegram: message the club" if kinds.get(chat) == "person" else "Telegram: club group chat"
        target = "https://" + chat
    else:
        button, target = "Instagram DM", f"https://ig.me/m/{c['ig']}"
    ups = sum(1 for h in hikes if h.get("cs") == slug(c["ig"]) and h["k"] != "big")
    ws.append([
        c["name"], c.get("type", ""), "yes" if c.get("active", True) else "no (hidden)", ups, button, target,
        f"https://instagram.com/{c['ig']}",
        "\n".join(f"{l} ({k})" for l, k in kinds.items()) or "",
        c.get("whatsapp", ""), "\n".join(c.get("phone", [])), c.get("email", ""), c.get("web", ""), c.get("followers", ""),
    ])
    fill = GREEN if chat else AMBER
    ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=fill)
    if not c.get("active", True):
        for col in range(1, len(head) + 1):
            ws.cell(ws.max_row, col).fill = PatternFill("solid", fgColor=GREY)

for col in range(1, len(head) + 1):
    cell = ws.cell(1, col)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="3D6A49")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
widths = [30, 15, 13, 10, 28, 36, 36, 40, 22, 18, 26, 30, 10]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(vertical="top", wrap_text=True)
ws.freeze_panes = "B2"
ws.auto_filter.ref = ws.dimensions

n_chat = sum(1 for c in clubs if c.get("chat"))
info = wb.create_sheet("How to read this")
for line in [
    ["Club contacts for AllHikes"],
    [],
    [f"{len(clubs)} clubs. {n_chat} have a Telegram chat people can message; the rest get an Instagram DM button."],
    ["Green 'Button on the website' = Telegram chat. Amber = Instagram DM (no Telegram chat found)."],
    ["Grey rows are clubs hidden on the site (active = false in data/clubs.json)."],
    ["Telegram kinds: person = a person's account (best: the club's rep); group = a group chat anyone can join;"],
    ["channel = announcements only, people can't write there, so it's never used for the button; bot = a bot."],
    ["When a club gives us its rep's Telegram account, add it to the club's telegram list in data/clubs.json,"],
    ["then run: python scripts/check_telegram.py, python scripts/build.py, python scripts/contacts_sheet.py"],
]:
    info.append(line)
info["A1"].font = Font(bold=True, size=14)
info.column_dimensions["A"].width = 110

out = ROOT / "outreach/club-contacts.xlsx"
wb.save(out)
print(f"{out.relative_to(ROOT)}: {len(clubs)} clubs, {n_chat} with a Telegram chat button")
