# -*- coding: utf-8 -*-
"""Интерактивный отчёт для родителей. Собирается ИЗ журнала кассы, руками не правится."""
import datetime, json
from openpyxl import load_workbook

import os as _os
SRC = _os.environ.get("SRC", "/home/claude/Касса_2В_2026-2027.xlsx")
OUT = _os.environ.get("OUT", "/home/claude/index.html")
AS_OF = datetime.date(2026, 8, 7)
import os
STAMP = os.environ.get("STAMP", "-")
MASK = os.environ.get("MASK") == "1"
PASSWORD = os.environ.get("PASSWORD", "")


def encrypt_payload(plaintext: str, password: str) -> str:
    """AES-256-GCM, ключ из PBKDF2-SHA256. Возвращает base64(salt|iv|ciphertext)."""
    import base64
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    salt = os.urandom(16)               # новые при каждой сборке: повтор nonce ломает GCM
    iv = os.urandom(12)
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=250000)
    key = kdf.derive(password.encode())
    ct = AESGCM(key).encrypt(iv, plaintext.encode(), None)
    return base64.b64encode(salt + iv + ct).decode()


def pub(n):
    """Публичная форма имени: «Елисеев Артур» -> «Артур Е.». Фамилия не раскрывается."""
    if not MASK or not n:
        return n
    parts = str(n).split()
    return f"{parts[1]} {parts[0][0]}." if len(parts) > 1 else n
S_FIRST, S_LAST, B_FIRST, B_LAST, MC = 6, 27, 6, 17, 4

wb = load_workbook(SRC, data_only=True)
sv, sb, uch, dgy, dgs, uc, vz, rs = (wb["Свод"], wb["Сборы"], wb["Ученики"], wb["Долги_год"],
                                     wb["Долги_срок"], wb["Участие"], wb["Взносы"], wb["Расходы"])
totals = {"collected": sv["C6"].value or 0, "spent": sv["C7"].value or 0,
          "rest": sv["C8"].value or 0, "dueNow": sv["C9"].value or 0,
          "dueYear": sv["C10"].value or 0}


def dt(x):
    return x.strftime("%d.%m.%Y") if hasattr(x, "strftime") else ""


def num(x):
    return x if isinstance(x, (int, float)) else 0


kids = []
for r in range(S_FIRST, S_LAST + 1):
    n = uch.cell(row=r, column=3).value
    if not n:
        continue
    kids.append({"row": r, "col": MC + (r - S_FIRST), "no": len(kids) + 1, "raw": n, "name": pub(n),
                 "prev": num(uch.cell(row=r, column=4).value),
                 "paid": num(uch.cell(row=r, column=5).value),
                 "debtY": num(uch.cell(row=r, column=7).value),
                 "debtN": num(uch.cell(row=r, column=8).value),
                 "share": num(uch.cell(row=r, column=9).value),
                 "rest": num(uch.cell(row=r, column=10).value),
                 "bday": str(uch.cell(row=r, column=11).value or "")[:5],
                 "first": (n.split()[1] if len(n.split()) > 1 else n), "by": []})

sbory = []
for r in range(B_FIRST, B_LAST + 1):
    title = sb.cell(row=r, column=3).value
    if not title:
        continue
    per, first = num(sb.cell(row=r, column=4).value), num(sb.cell(row=r, column=5).value)
    share = num(sb.cell(row=r, column=15).value)
    parts = []
    for k in kids:
        if uc.cell(row=r, column=k["col"]).value != 1:
            continue
        p = sum(num(vz.cell(row=vr, column=5).value) for vr in range(6, 86)
                if vz.cell(row=vr, column=3).value == k["raw"]
                and vz.cell(row=vr, column=4).value == title)
        dy, dn = num(dgy.cell(row=r, column=k["col"]).value), num(dgs.cell(row=r, column=k["col"]).value)
        parts.append({"no": k["no"], "first": k["first"], "name": k["name"], "paid": p, "debtY": dy, "debtN": dn})
        k["by"].append({"sbor": title, "code": sb.cell(row=r, column=2).value or "",
                        "plan": per, "first": first, "paid": p,
                        "debtY": dy, "debtN": dn, "share": share, "rest": p - share})
    spends = [{"date": dt(rs.cell(row=e, column=2).value), "what": rs.cell(row=e, column=4).value or "",
               "amount": num(rs.cell(row=e, column=6).value),
               "proof": str(rs.cell(row=e, column=9).value or "со слов").lower()}
              for e in range(6, 66)
              if rs.cell(row=e, column=3).value == title and rs.cell(row=e, column=6).value]
    sbory.append({"code": sb.cell(row=r, column=2).value or "", "title": title, "per": per,
                  "first": first, "due": dt(sb.cell(row=r, column=6).value),
                  "dueFull": dt(sb.cell(row=r, column=7).value),
                  "event": dt(sb.cell(row=r, column=8).value),
                  "n": num(sb.cell(row=r, column=9).value), "plan": num(sb.cell(row=r, column=10).value),
                  "collected": num(sb.cell(row=r, column=11).value),
                  "debtY": num(sb.cell(row=r, column=12).value),
                  "debtN": num(sb.cell(row=r, column=13).value),
                  "spent": num(sb.cell(row=r, column=14).value),
                  "rest": num(sb.cell(row=r, column=16).value), "parts": parts, "spends": spends})

for k in kids:
    k["pays"] = [{"date": dt(vz.cell(row=vr, column=2).value),
                  "sbor": vz.cell(row=vr, column=4).value or "",
                  "amount": num(vz.cell(row=vr, column=5).value),
                  "way": vz.cell(row=vr, column=6).value or ""}
                 for vr in range(6, 86)
                 if vz.cell(row=vr, column=3).value == k["raw"] and vz.cell(row=vr, column=5).value]

all_spends = [{"date": dt(rs.cell(row=e, column=2).value), "sbor": rs.cell(row=e, column=3).value or "не привязан",
               "what": rs.cell(row=e, column=4).value or "", "cat": rs.cell(row=e, column=5).value or "без направления",
               "amount": num(rs.cell(row=e, column=6).value),
               "proof": str(rs.cell(row=e, column=9).value or "со слов").lower()}
              for e in range(6, 66) if rs.cell(row=e, column=6).value]

TEACHER = {"name": "Учитель класса", "date": "30.08", "role": "учитель"}
EVENTS = []
for s in sbory:
    if s["due"]:
        EVENTS.append({"date": s["due"], "kind": "due", "code": s["code"],
                       "title": f'Внести {int(s["first"]) if s["first"] else 0} \u20bd - {s["title"]}'})
    if s.get("dueFull"):
        EVENTS.append({"date": s["dueFull"], "kind": "due", "code": s["code"],
                       "title": f'Полная сумма {int(s["per"]) if s["per"] else 0} \u20bd - {s["title"]}'})
    if s["event"]:
        EVENTS.append({"date": s["event"], "kind": "event", "code": s["code"],
                       "title": s["title"]})
DATA = json.dumps({"t": totals, "teacher": TEACHER, "events": EVENTS, "sbory": sbory, "spends": all_spends,
                   "kids": [{a: b for a, b in k.items() if a not in ("row", "col", "raw")} for k in kids]},
                  ensure_ascii=False)

PAGE = r"""<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow, noarchive">
<meta name="referrer" content="no-referrer">
<title>Касса класса 2В - отчёт на __ASOF__</title>
<style>
 :root{--ink:#16181d;--dim:#6b7280;--line:#e6e8ec;--bg:#fff;--accent:#2f5496;--good:#0f766e;
       --bad:#b3261e;--warn:#8a6100;--soft:#f7f8fa;--track:#eef1f6;}
 *{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}
 body{margin:0;background:var(--soft);color:var(--ink);
      font:16px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;}
 .wrap{max-width:780px;margin:0 auto;padding:22px 14px 60px;}
 #gate{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;}
 .gcard{background:var(--bg);border:1px solid var(--line);border-radius:14px;padding:26px 24px;
        max-width:380px;width:100%;box-shadow:0 4px 24px rgba(22,24,29,.07);}
 .gh{margin:0 0 6px;font-size:19px;} .gp{margin:0 0 16px;color:var(--dim);font-size:14px;}
 #gform{display:flex;gap:8px;} #gform.busy{opacity:.5;pointer-events:none;}
 #gpass{flex:1;min-width:0;padding:11px 13px;border:1px solid var(--line);border-radius:9px;
        font:inherit;font-size:15px;}
 #gform button{border:0;background:var(--accent);color:#fff;font:inherit;font-weight:600;
               padding:11px 18px;border-radius:9px;cursor:pointer;}
 .gerr{color:var(--bad);font-size:13.5px;margin-top:10px;min-height:18px;}
 html{scroll-behavior:smooth;scrollbar-gutter:stable;overflow-y:scroll;}
 header{padding:6px 0 18px;border-bottom:2px solid var(--accent);margin-bottom:18px;}
 h1{font-size:23px;margin:0 0 5px;letter-spacing:-.01em;}
 .sub{color:var(--dim);font-size:13.5px;}
 .tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:18px;}
 .tile{background:var(--bg);border:1px solid var(--line);border-radius:11px;padding:11px 12px;}
 .tile .lab{font-size:10.5px;color:var(--dim);text-transform:uppercase;letter-spacing:.05em;}
 .tile .val{font-size:18px;font-weight:700;margin-top:4px;white-space:nowrap;}
 .tile.rest{border-color:var(--good);} .tile.rest .val{color:var(--good);}
 .tile.owed{border-color:#e3c26b;} .tile.owed .val{color:var(--warn);}
 nav{display:flex;flex-wrap:wrap;gap:6px;background:var(--bg);border:1px solid var(--line);
     border-radius:11px;padding:5px;margin-bottom:16px;position:sticky;top:8px;z-index:5;
     box-shadow:0 2px 10px rgba(22,24,29,.06);}
 nav button{flex:1 1 auto;border:0;background:transparent;font:inherit;font-size:14.5px;padding:9px 6px;
            border-radius:8px;cursor:pointer;color:var(--dim);}
 nav button[aria-selected=true]{background:var(--accent);color:#fff;font-weight:600;}
 .card{background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:13px 15px;
       margin-bottom:9px;}
 .card-top{display:flex;align-items:flex-start;gap:10px;}
 .code.sm{font-size:10px;padding:2px 5px;}
 .code{flex:none;font-size:11px;font-weight:700;color:var(--accent);background:#eef2fa;
       border-radius:6px;padding:3px 7px;margin-top:2px;}
 h3{font-size:16px;margin:0;} .meta{color:var(--dim);font-size:12.5px;margin-top:3px;}
 .bar{height:7px;background:var(--track);border-radius:99px;margin:11px 0 5px;overflow:hidden;}
 .bar i{display:block;height:100%;background:var(--accent);border-radius:99px;}
 .barlab{display:flex;justify-content:space-between;font-size:12.5px;color:var(--dim);}
 .stats{display:flex;flex-wrap:wrap;gap:6px 16px;margin:10px 0 12px;font-size:13px;
        color:var(--dim);}
 .stats b{color:var(--ink);font-size:14px;font-variant-numeric:tabular-nums;white-space:nowrap;}
 .stats b.good{color:var(--good);} .stats b.bad{color:var(--bad);}
 .acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px;}
 .btn{border:1px solid var(--line);background:var(--bg);border-radius:8px;padding:7px 12px;
      font:inherit;font-size:13.5px;color:var(--accent);cursor:pointer;}
 .btn:hover{background:#f3f6fc;} .btn[aria-expanded=true]{background:#eef2fa;font-weight:600;}
 .chev{display:inline-block;transition:transform .15s;margin-left:5px;font-size:11px;flex:none;}
 .btn[aria-expanded=true] .chev{transform:rotate(180deg);}
 .panel{display:none;margin-top:10px;} .panel.open{display:block;}
 ul.list{list-style:none;margin:0;padding:0;font-size:14.5px;}
 ul.list li{display:flex;align-items:center;gap:9px;padding:8px 0;border-bottom:1px solid var(--line);}
 ul.list li:last-child{border-bottom:none;}
 .dot{flex:none;width:8px;height:8px;border-radius:99px;background:var(--good);}
 .idx{flex:none;min-width:20px;font-size:12px;color:var(--dim);
      font-variant-numeric:tabular-nums;text-align:right;}
 .dot.part{background:#d4a017;} .dot.no{background:var(--bad);}
 .amt{margin-left:auto;font-variant-numeric:tabular-nums;white-space:nowrap;font-weight:600;}
 .amt.bad{color:var(--bad);} .amt.ok{color:var(--good);}
 .amt.soft{color:var(--dim);font-weight:400;}
 .tag{font-size:10.5px;color:var(--dim);border:1px solid var(--line);border-radius:20px;
      padding:1px 7px;white-space:nowrap;}
 .search{width:100%;padding:11px 13px;border:1px solid var(--line);border-radius:10px;font:inherit;
         font-size:15px;margin-bottom:10px;background:var(--bg);}
 .kid{background:var(--bg);border:1px solid var(--line);border-radius:11px;margin-bottom:7px;}
 .kid-h{display:flex;align-items:center;gap:10px;padding:12px 14px;cursor:pointer;}
 .kid-h .nm{font-weight:600;font-size:15px;}
 .kid-h .st{margin-left:auto;text-align:right;font-size:12.5px;color:var(--dim);white-space:nowrap;}
 .kid-h .st b{display:block;font-size:15px;color:var(--ink);}
 .kid-body{display:none;padding:0 14px 14px;} .kid.open .kid-body{display:block;}
 .scroller{overflow-x:auto;-webkit-overflow-scrolling:touch;}
 table{width:100%;min-width:640px;border-collapse:collapse;font-size:13px;}
 thead th{white-space:nowrap;vertical-align:bottom;line-height:1.2;}
 th{text-align:left;font-size:10.5px;text-transform:uppercase;color:var(--dim);font-weight:600;
    padding:7px 8px;border-bottom:1px solid var(--line);}
 th.num,td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}
 td{padding:8px;border-bottom:1px solid var(--line);} tr:last-child td{border-bottom:none;}
 .empty{background:var(--bg);border:1px dashed var(--line);border-radius:11px;padding:16px;
        color:var(--dim);font-size:14.5px;}
 .note{background:var(--bg);border:1px solid var(--line);border-left:3px solid var(--accent);
       border-radius:9px;padding:12px 15px;margin-top:10px;font-size:14px;color:#39414f;}
 .note b{color:var(--ink);}
 h2{font-size:16px;margin:26px 0 10px;color:var(--accent);}
 .chips{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px;align-items:center;}
 .sortbox{margin:2px 0 8px;} .sortbox .slab{font-size:11px;color:var(--dim);
  text-transform:uppercase;letter-spacing:.04em;margin-right:2px;}
 .chip.sc{padding:4px 10px;font-size:12.5px;}
 .chip{border:1px solid var(--line);background:var(--bg);border-radius:20px;padding:5px 12px;
       font:inherit;font-size:13px;color:var(--dim);cursor:pointer;}
 .chip[aria-pressed=true]{background:var(--accent);color:#fff;border-color:var(--accent);font-weight:600;}
 .rows{background:var(--bg);border:1px solid var(--line);border-radius:11px;overflow:hidden;}
 .row{border-bottom:1px solid var(--line);}
 .row:last-child{border-bottom:none;}
 .row-h{display:flex;align-items:center;gap:7px;padding:9px 11px;cursor:pointer;font-size:14.5px;}
 .row-h .nm{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
 .row-h .due{font-variant-numeric:tabular-nums;white-space:nowrap;font-weight:700;color:var(--bad);
             font-size:13px;min-width:66px;text-align:right;}
 .row-h .bal{font-variant-numeric:tabular-nums;white-space:nowrap;font-weight:700;min-width:78px;
             text-align:right;}
 .row-h .bal.neg{color:var(--bad);}
 .row-b{display:none;padding:2px 10px 12px;} .row.open .row-b{display:block;}
 .row-head{display:flex;align-items:center;gap:8px;padding:8px 12px;background:#fbfcfe;
           border-bottom:1px solid var(--line);font-size:10.5px;text-transform:uppercase;
           letter-spacing:.04em;color:var(--dim);font-weight:600;}
 .row-head .bal{color:var(--dim);font-weight:600;font-size:10.5px;}
 .pays{margin-top:10px;border-top:1px solid var(--line);padding-top:8px;}
 .pays-h{font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--dim);
         font-weight:700;margin-bottom:4px;}
 .pay{display:flex;align-items:baseline;gap:8px;font-size:12.5px;padding:3px 0;}
 .pay .pd{flex:none;font-variant-numeric:tabular-nums;font-weight:600;}
 .pay .pn{color:var(--dim);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
 .pay .pa{margin-left:auto;font-weight:700;font-variant-numeric:tabular-nums;white-space:nowrap;}
 .sum{font-size:12.5px;color:var(--dim);line-height:1.5;padding:6px 2px 8px;}
 .sum b{color:var(--ink);font-variant-numeric:tabular-nums;} .sum b.bad{color:var(--bad);}
 .row-h .chev{color:var(--dim);}
 .row.open .row-h .chev{transform:rotate(180deg);}
 .row.open{border:2px solid var(--accent);border-radius:10px;margin:6px 0;
           box-shadow:0 2px 10px rgba(47,84,150,.10);}
 .row.open .row-h{background:#eef2fa;border-radius:8px 8px 0 0;}
 .row.open .row-b{background:#fbfcfe;border-radius:0 0 8px 8px;}
 .hero{background:var(--bg);border:1px solid #e3c26b;border-radius:11px;padding:12px 14px;
        margin-bottom:14px;}
 .hero .lab{display:block;font-size:10.5px;color:var(--warn);text-transform:uppercase;
            letter-spacing:.05em;font-weight:700;margin-bottom:6px;}
 .hero-i{display:flex;align-items:baseline;gap:8px;padding:4px 0;font-size:14.5px;flex-wrap:wrap;}
 .hero-i b{white-space:nowrap;} .hero-i .in{margin-left:auto;font-size:12.5px;color:var(--warn);}
 .hero.none{border-color:var(--line);color:var(--dim);font-size:14px;}
 .cal{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;}
 @media(min-width:560px){.cal{grid-template-columns:repeat(3,1fr);}}
 .m{background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:9px 11px 10px;}
 .m.cur{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent) inset;}
 .m h4{margin:0 0 5px;font-size:11px;text-transform:uppercase;letter-spacing:.04em;
       color:var(--accent);font-weight:700;}
 .m.mt h4{color:#b6bcc7;}
 .m ul{list-style:none;margin:0;padding:0;font-size:12.5px;}
 .m li{display:flex;gap:6px;padding:2px 0;align-items:baseline;}
 .m li b{flex:none;min-width:17px;font-variant-numeric:tabular-nums;color:var(--dim);font-weight:600;}
 .m li span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
 .m li.soon b,.m li.soon span{color:var(--warn);font-weight:700;}
 .m li i{font-style:normal;color:var(--dim);font-size:11px;}
 .m li.l-due b,.m li.l-due span{color:var(--bad);} .m li.l-due{font-weight:600;}
 .m li.l-event b,.m li.l-event span{color:var(--accent);}
 .k{font-size:10px;text-transform:uppercase;letter-spacing:.03em;border-radius:20px;
    padding:1px 7px;border:1px solid var(--line);color:var(--dim);}
 .k-due{color:var(--bad);border-color:#f3c9c6;} .k-event{color:var(--accent);border-color:#c9d6ee;}
 .m .none{color:#c9ced7;font-size:12.5px;}
 .mgroup{margin-bottom:12px;}
 .mhead{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:700;
        color:var(--accent);margin:0 0 6px 2px;text-transform:uppercase;letter-spacing:.03em;}
 .mhead span{font-weight:400;color:var(--dim);margin-left:auto;text-transform:none;
              font-variant-numeric:tabular-nums;}
 ul.list.pad12{padding:0 12px;} .sub{color:var(--dim);font-size:12px;}
 .bd{display:flex;align-items:center;gap:10px;padding:9px 12px;border-bottom:1px solid var(--line);
     font-size:14.5px;}
 .bd:last-child{border-bottom:none;}
 .bd .date{flex:none;min-width:52px;font-weight:700;font-variant-numeric:tabular-nums;}
 .bd .in{margin-left:auto;font-size:12.5px;color:var(--dim);white-space:nowrap;}
 .bd.soon{background:#fff8e6;} .bd.soon .in{color:var(--warn);font-weight:700;}
 .hint{margin:8px 2px 0;font-size:12.5px;color:var(--dim);line-height:1.45;}
 .hint b{color:var(--ink);font-weight:600;}
 .expl{margin-top:18px;} .btn.wide{width:100%;text-align:left;padding:10px 14px;}
 .expl dl{margin:10px 0 0;background:var(--bg);border:1px solid var(--line);border-radius:11px;
          padding:4px 15px 12px;font-size:13.5px;}
 .expl dt{font-weight:700;margin-top:11px;} .expl dd{margin:2px 0 0;color:#4b5563;}
 footer{margin-top:30px;padding-top:14px;border-top:1px solid var(--line);color:var(--dim);
        font-size:12.5px;}
 @media(max-width:560px){.tiles{grid-template-columns:1fr 1fr;}}
 @media print{
  body{background:#fff;} .wrap{max-width:none;padding:0;}
  nav,.acts,.search{display:none!important;}
  [hidden]{display:block!important;}
  .panel{display:block!important;} .kid-body{display:block!important;}
  .card,.kid{break-inside:avoid;page-break-inside:avoid;}
  table{min-width:0;} .scroller{overflow:visible;}
  #pane-kids::before{content:"По каждому ребёнку";display:block;font-size:16px;
   color:var(--accent);font-weight:700;margin:24px 0 10px;}
  #pane-spends::before{content:"Все траты";display:block;font-size:16px;
   color:var(--accent);font-weight:700;margin:24px 0 10px;}
  .expl .btn{display:none!important;} .expl dl{break-inside:avoid;}
  .row-b{display:block!important;} .row{break-inside:avoid;} .chips{display:none!important;}
 }
</style></head><body>
 <div id="gate">
  <div class="gcard">
   <h2 class="gh">Доступ по слову</h2>
   <p class="gp">Отчёт закрыт от посторонних. Введите слово, которое дала Аня.</p>
   <form id="gform" autocomplete="on">
    <input id="gpass" type="password" name="password" autocomplete="current-password"
           placeholder="кодовое слово" autocapitalize="off" autocorrect="off" spellcheck="false">
    <button type="submit">Открыть</button>
   </form>
   <div class="gerr" id="gerr"></div>
  </div>
 </div>
<div class="wrap" id="app" hidden>
 <header><h1>Касса класса 2В</h1>
  <div class="sub">Учебный год 2026/2027 · отчёт на __ASOF__ · казначей __TREAS__</div>
 </header>
 <div class="tiles" id="tiles"></div>
 <span id="navtop"></span>
 <nav role="tablist">
  <button role="tab" data-tab="bdays" aria-selected="true">События</button>
  <button role="tab" data-tab="sbory" aria-selected="false">Сборы</button>
  <button role="tab" data-tab="kids" aria-selected="false">По детям</button>
  <button role="tab" data-tab="spends" aria-selected="false">Расходы</button>
 </nav>
 <div id="pane-bdays"></div>
 <div class="expl" id="ex-bdays"></div>
 <div id="pane-sbory" hidden></div>
 <div class="expl" id="ex-sbory" hidden></div>
 <div id="pane-kids" hidden>
  <input class="search" id="q" type="search" placeholder="__SEARCHPH__" autocomplete="off">
  <div class="chips" id="chips"></div>
  <div id="sortk"></div>
  <div id="kidlist"></div>
  <div class="hint" id="kidhint"></div>
  <div class="expl" id="ex-kids"></div>
 </div>
 <div id="pane-spends" hidden></div>
 <div class="expl" id="ex-spends" hidden></div>

 <footer>Сформировано __ASOF__ автоматически из журнала «Касса_2В_2026-2027.xlsx».
  Цифры руками не набираются. Версия __STAMP__.</footer>
 </div>
</div>
<script>
const RAW=__DATA__;

async function decryptPayload(b64, pass){
 const bin=Uint8Array.from(atob(b64), c=>c.charCodeAt(0));
 const salt=bin.slice(0,16), iv=bin.slice(16,28), ct=bin.slice(28);
 const km=await crypto.subtle.importKey('raw',new TextEncoder().encode(pass),'PBKDF2',false,['deriveKey']);
 const key=await crypto.subtle.deriveKey(
   {name:'PBKDF2',salt:salt,iterations:250000,hash:'SHA-256'},km,
   {name:'AES-GCM',length:256},false,['decrypt']);
 const out=await crypto.subtle.decrypt({name:'AES-GCM',iv:iv},key,ct);
 return JSON.parse(new TextDecoder().decode(out));
}

function boot(D){

const rub=n=>{n=Math.round((n||0)*100)/100;const s=Number.isInteger(n)?n.toLocaleString('ru-RU'):
  n.toLocaleString('ru-RU',{minimumFractionDigits:2,maximumFractionDigits:2});
  return s.replace(/ /g,' ')+' ₽';};
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const T=D.t;
const tiles=[['Собрано',T.collected,''],['Потрачено',T.spent,''],['Остаток',T.rest,'rest']];
if(T.dueNow)tiles.push(['Ждём к сроку',T.dueNow,'owed']);
document.getElementById('tiles').innerHTML=tiles.map(([l,v,c])=>
 `<div class="tile ${c}"><div class="lab">${l}</div><div class="val">${rub(v)}</div></div>`).join('');
if(tiles.length===3)document.getElementById('tiles').style.gridTemplateColumns='repeat(3,1fr)';

let SORT='name';
const cmp=(a,b)=>SORT==='name'
 ? String(a.first||a.name).localeCompare(String(b.first||b.name),'ru')
 : (a.no||0)-(b.no||0);
const sortChips=id=>`<div class="chips sortbox"><span class="slab">Сортировка</span>`+
 [['name','по имени'],['last','по фамилии']].map(([k,l])=>
  `<button class="chip sc" data-sort="${k}" data-src="${id}" aria-pressed="${k===SORT}">${l}</button>`).join('')+
 `</div>`;

function renderSbory(){
document.getElementById('pane-sbory').innerHTML=D.sbory.map((s,i)=>{
 const pct=s.plan?Math.min(100,Math.round(s.collected/s.plan*100)):0;
 const done=s.parts.filter(p=>!p.debtN).length;
 const parts=[...s.parts].sort(cmp).map(p=>{
  const cls=!p.debtN?(p.debtY?'':''):(p.paid?'part':'no');
  const right=p.debtN?`<span class="amt bad">нужно сейчас ${rub(p.debtN)}</span>`
    :(p.debtY?`<span class="amt soft">ещё ${rub(p.debtY)} за год</span>`
             :`<span class="amt ok">рассчитался</span>`);
  return `<li><span class="dot ${cls}"></span><span class="idx">${p.no}</span>${esc(p.name)}
    <span class="tag">внёс ${rub(p.paid)}</span>${right}</li>`;}).join('');
 const sp=s.spends.length?s.spends.map(e=>`<li>${e.date?`<span class="tag">${esc(e.date)}</span>`:''}
   ${esc(e.what)}<span class="amt">${rub(e.amount)}</span>
   <span class="tag">${esc(e.proof)}</span></li>`).join('')
  :'<li style="color:var(--dim)">Из этого сбора пока ничего не потрачено.</li>';
 const meta=[`${s.n} участников`];
 if(s.event)meta.push(`событие ${s.event.slice(0,5)}`);
 if(s.first)meta.push(`по ${rub(s.first)}${s.due?' до '+s.due.slice(0,5):''}`);
 if(s.per&&s.per!==s.first)meta.push(`всего ${rub(s.per)}${s.dueFull?' до '+s.dueFull.slice(0,5):' - платежами в течение года'}`);
 return `<section class="card">
  <div class="card-top">${s.code?`<span class="code">${esc(s.code)}</span>`:''}
   <div><h3>${esc(s.title)}</h3><div class="meta">${meta.join(' · ')}</div></div></div>
  ${s.plan?`<div class="bar"><i style="width:${pct}%"></i></div>
   <div class="barlab"><span>собрано ${rub(s.collected)} из ${rub(s.plan)}</span><span>${pct}%</span></div>`:''}
  <div class="stats">
   <span>потрачено <b>${rub(s.spent)}</b></span>
   <span>остаток <b class="good">${rub(s.rest)}</b></span>
   <span>ждём к сроку <b${s.debtN?' class="bad"':''}>${rub(s.debtN)}</b></span>
  </div>
  <div class="acts">
   <button class="btn" data-toggle="pp${i}" aria-expanded="false">Участники · внесли к сроку ${done} из ${s.parts.length}<span class="chev">▾</span></button>
   <button class="btn" data-toggle="ps${i}" aria-expanded="false">Расходы · ${s.spends.length}<span class="chev">▾</span></button>
  </div>
  <div class="panel" id="pp${i}">${sortChips('sbory')}<ul class="list">${parts}</ul></div>
  <div class="panel" id="ps${i}"><ul class="list">${sp}</ul></div></section>`;}).join('')
 ||'<div class="empty">Сборов пока нет.</div>';}
renderSbory();

let KFILTER='all';
function renderKids(f){f=(f||'').trim().toLowerCase();
 let L=D.kids.filter(k=>!f||k.name.toLowerCase().includes(f));
 L=[...L].sort(cmp);
 if(KFILTER==='debt')L=L.filter(k=>k.debtN>0);
 if(KFILTER==='year')L=L.filter(k=>!k.debtN&&k.debtY>0);
 if(KFILTER==='ok')L=L.filter(k=>!k.debtN&&!k.debtY);
 document.getElementById('kidlist').innerHTML=L.length?`<div class="rows">
  <div class="row-head"><span class="idx"></span>
   <span class="nm">Ученик</span><span class="bal">Остаток</span><span class="chev" style="visibility:hidden">▾</span></div>
  ${L.map(k=>{
  const rows=k.by.map(b=>`<tr>
    <td>${b.code?`<span class="code sm">${esc(b.code)}</span> `:''}${esc(b.sbor)}</td>
    <td class="num">${rub(b.first)}</td>
    <td class="num">${rub(b.plan)}</td>
    <td class="num">${rub(b.paid)}</td>
    <td class="num"${b.debtN?' style="color:var(--bad);font-weight:700"':''}>${b.debtN?rub(b.debtN):'-'}</td>
    <td class="num">${b.debtY?rub(b.debtY):'-'}</td>
    <td class="num">${rub(b.share)}</td>
    <td class="num"><b>${rub(b.rest)}</b></td></tr>`).join('')
   ||'<tr><td colspan="8" style="color:var(--dim)">Ни в одном сборе не участвует.</td></tr>';
  return `<div class="row"><div class="row-h">
    <span class="idx">${k.no}</span>
    <span class="nm">${esc(k.name)}</span>
    <span class="bal${k.rest<0?' neg':''}">${rub(k.rest)}</span>
    <span class="chev">▾</span></div>
   <div class="row-b">
    <div class="sum">Внесено <b>${rub(k.paid)}</b> · доля расходов <b>${rub(k.share)}</b>
     · остаток <b${k.rest<0?' class="bad"':''}>${rub(k.rest)}</b>${k.debtN?`
     · <b class="bad">нужно внести сейчас ${rub(k.debtN)}</b>`:(k.debtY?`
     · ещё ${rub(k.debtY)} за год`:' · рассчитался')}</div>
    <div class="scroller"><table>
    <thead><tr>
     <th>Сбор</th>
     <th class="num">Внести<br>к сроку</th>
     <th class="num">Всего<br>за год</th>
     <th class="num">Внесено</th>
     <th class="num">Долг<br>сейчас</th>
     <th class="num">Ещё<br>за год</th>
     <th class="num">Доля<br>расходов</th>
     <th class="num">Остаток</th></tr></thead>
    <tbody>${rows}</tbody></table></div>
    ${k.pays.length?`<div class="pays"><div class="pays-h">Платежи</div>
      ${k.pays.map(pp=>`<div class="pay"><span class="pd">${esc(pp.date)}</span>
        <span class="pn">${esc(pp.sbor)}</span>
        <span class="pa">${rub(pp.amount)}</span>
        <span class="tag">${esc(String(pp.way).toLowerCase())}</span></div>`).join('')}</div>`
     :'<div class="pays"><div class="pays-h">Платежей ещё не было</div></div>'}
    </div></div>`;}).join('')}</div>`
  :'<div class="empty">Никого не нашли.</div>';
 const h=document.getElementById('kidhint');
 if(h)h.innerHTML=`Остаток - сколько денег ребёнка ещё не потрачено: внесено минус своя доля расходов
  (сейчас ${rub(D.t.spent)} ÷ ${D.kids.length} = ${rub(D.t.spent/D.kids.length)} с человека).
  <b>Минус означает, что за ребёнка уже потрачено больше, чем он внёс.</b>`;}

const cnt={all:D.kids.length,debt:D.kids.filter(k=>k.debtN>0).length,
 year:D.kids.filter(k=>!k.debtN&&k.debtY>0).length,ok:D.kids.filter(k=>!k.debtN&&!k.debtY).length};
document.getElementById('chips').innerHTML=[['all','Все'],['debt','Должны к сроку'],
 ['year','Только за год'],['ok','Рассчитались']].map(([k,l])=>
 `<button class="chip" data-chip="${k}" aria-pressed="${k==='all'}">${l} · ${cnt[k]}</button>`).join('');
document.getElementById('chips').addEventListener('click',e=>{
 const c=e.target.closest('.chip'); if(!c)return;
 KFILTER=c.dataset.chip;
 document.querySelectorAll('.chip').forEach(x=>x.setAttribute('aria-pressed',String(x===c)));
 renderKids(document.getElementById('q').value);});
document.getElementById('sortk').innerHTML=sortChips('kids');
renderKids('');
document.getElementById('q').addEventListener('input',e=>renderKids(e.target.value));

// дни рождения
const TODAY=new Date(__YEAR__,__MONTH__-1,__DAY__);
const bd=D.kids.filter(k=>k.bday).map(k=>{
 const [d,m]=k.bday.split('.').map(Number);
 let y=TODAY.getFullYear(), next=new Date(y,m-1,d);
 if(next<TODAY)next=new Date(y+1,m-1,d);
 return {name:k.name,no:k.no,date:k.bday,in:Math.round((next-TODAY)/86400000)};});
if(D.teacher){const [td,tm]=D.teacher.date.split('.').map(Number);
 let ty=TODAY.getFullYear(),tn=new Date(ty,tm-1,td); if(tn<TODAY)tn=new Date(ty+1,tm-1,td);
 bd.push({...D.teacher,no:'',in:Math.round((tn-TODAY)/86400000)});}
bd.sort((a,b)=>a.in-b.in);
const MONT=['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август',
 'Сентябрь','Октябрь','Ноябрь','Декабрь'];
const short=n=>{const p=n.split(' ');return p.length>1?p[0]+' '+p[1][0]+'.':n;};
const curM=TODAY.getMonth();
const daysTo=(d,m)=>{let y=TODAY.getFullYear(),n=new Date(y,m-1,d);
 if(n<TODAY)n=new Date(y+1,m-1,d);return {in:Math.round((n-TODAY)/86400000),y:n.getFullYear()};};
const ALL=[];
bd.forEach(b=>{const [d,m]=b.date.split('.').map(Number);const r=daysTo(d,m);
 ALL.push({kind:b.role?'teacher':'bd',d:d,m:m,in:r.in,
  name:b.role?'Учитель класса':b.name,short:b.role?'Учитель':short(b.name)});});
(D.events||[]).forEach(e=>{const [d,m]=e.date.split('.').map(Number);const r=daysTo(d,m);
 ALL.push({kind:e.kind,d:d,m:m,in:r.in,name:e.title,short:e.title,code:e.code});});
ALL.sort((a,b)=>a.in-b.in);
const KIND={bd:'др',teacher:'др',due:'срок',event:'событие'};
const dd=n=>String(n).padStart(2,'0');
const yr=n=>{const a=n%10,b=n%100;
 if(b>=11&&b<=14)return n+' лет'; if(a===1)return n+' год';
 if(a>=2&&a<=4)return n+' года'; return n+' лет';};
const soonList=ALL.filter(e=>e.in<=30);
const byMonth={};ALL.forEach(e=>{(byMonth[e.m-1]=byMonth[e.m-1]||[]).push(e);});
Object.values(byMonth).forEach(a=>a.sort((x,y)=>x.d-y.d));
const order=[];for(let i=0;i<12;i++)order.push((8+i)%12);
document.getElementById('pane-bdays').innerHTML=ALL.length?`
 ${soonList.length?`<div class="hero"><span class="lab">Ближайший месяц</span>
   ${soonList.map(e=>`<div class="hero-i"><b>${dd(e.d)}.${dd(e.m)}</b>
     <span class="k k-${e.kind}">${KIND[e.kind]}</span> ${esc(e.name)}
     <span class="in">${e.in===0?'сегодня':(e.in===1?'завтра':'через '+e.in+' дн.')}</span></div>`).join('')}
  </div>`:'<div class="hero none">В ближайший месяц событий нет.</div>'}
 <div class="cal">${order.map(m=>{
   const items=byMonth[m]||[];
   return `<div class="m${m===curM?' cur':''}${items.length?'':' mt'}">
    <h4>${MONT[m]}</h4>
    ${items.length?`<ul>${items.map(e=>`<li class="${e.in<=14?'soon':''} l-${e.kind}">
      <b>${dd(e.d)}</b><span>${esc(e.short)}</span></li>`).join('')}</ul>`
     :'<div class="none">-</div>'}</div>`;}).join('')}</div>`
 :'<div class="empty">Событий нет.</div>';

// пояснения по вкладкам
const EXPL={
 sbory:[['Два вида долга','«К сроку» - внести сейчас, это первоначальный взнос, у него есть дата. «За год» - сколько останется до полной суммы; конкретная дата для остатка родкомом не объявлена, платежи разбиты в течение учебного года.'],
  ['Участники','У каждого сбора свой список. Кто не участвует - не платит и в тратах по нему не участвует.'],
  ['Оплата частями','Каждый платёж записывается отдельно, долг уменьшается на внесённую сумму.']],
 kids:[['Остаток','Расходы сбора делятся поровну между участниками. Остаток = внесено минус своя доля трат.'],
  ['Минус в остатке','Значит доля уже потраченного больше внесённого - деньги ждём.'],
  ['Долги по сборам','Считаются раздельно: переплата по одному сбору не закрывает долг по другому. Нажмите на строку - видно по каждому сбору.'],
  ['Ошибка?','Если платежа нет или сумма не совпадает - напишите Ане, поправим и перевыпустим отчёт.']],
 spends:[['Подтверждения','Бумажные чеки не собираются. У каждого расхода указано, чем он подтверждён: скрин оплаты, чек или только со слов.'],
  ['К какому сбору','Каждый расход привязан к сбору и делится только между его участниками.'],
  ['Направление','Это категория расхода: праздник, подарки, канцтовары и так далее. По ней видно, на что уходят деньги класса.']],
 bdays:[['Откуда даты','Из списка класса. Год рождения не хранится - для поздравления нужны только день и месяц.'],
  ['Календарь','Месяцы идут по учебному году, с сентября. Жёлтым - ближайшие две недели, рамкой - текущий месяц. Имена сокращены до инициала, полные - в блоке сверху.']]};
for(const [k,items] of Object.entries(EXPL)){
 const el=document.getElementById('ex-'+k); if(!el)continue;
 el.innerHTML=`<button class="btn wide" data-toggle="exp-${k}" aria-expanded="false">Как это считается<span class="chev">▾</span></button>
  <div class="panel" id="exp-${k}"><dl>${items.map(([a,b])=>`<dt>${a}</dt><dd>${b}</dd>`).join('')}</dl></div>`;}

let SGROUP='date';
function renderSpends(){
 const el=document.getElementById('pane-spends');
 if(!D.spends.length){el.innerHTML='<div class="empty">Пока не потрачено ни рубля.</div>';return;}
 const key=e=>SGROUP==='date'?e.date:(SGROUP==='cat'?e.cat:e.sbor);
 const order=[...D.spends];
 if(SGROUP==='date')order.sort((a,b)=>{
  const p=s=>s.split('.').reverse().join(''); return p(b.date).localeCompare(p(a.date));});
 else order.sort((a,b)=>key(a).localeCompare(key(b))||a.date.localeCompare(b.date));
 const gs=[];
 order.forEach(e=>{let g=gs[gs.length-1];
  if(!g||g.k!==key(e)){g={k:key(e),items:[],sum:0};gs.push(g);}
  g.items.push(e); g.sum+=e.amount;});
 const chips=[['date','По датам'],['cat','По направлениям'],['sbor','По сборам']].map(([k,l])=>
  `<button class="chip" data-sg="${k}" aria-pressed="${k===SGROUP}">${l}</button>`).join('');
 el.innerHTML=`<div class="chips" id="sgroup">${chips}</div>
  ${gs.map(g=>`<div class="mgroup">
   <div class="mhead">${esc(g.k)}<span>${rub(g.sum)}</span></div>
   <div class="rows"><ul class="list pad12">${g.items.map(e=>`<li>
     ${SGROUP!=='date'?`<span class="tag">${esc(e.date)}</span>`:''}
     <span>${esc(e.what)}${SGROUP!=='cat'?`<br><span class="sub">${esc(e.cat)}</span>`:''}</span>
     <span class="amt">${rub(e.amount)}</span><span class="tag">${esc(e.proof)}</span></li>`).join('')}
   </ul></div></div>`).join('')}`;
 document.getElementById('sgroup').addEventListener('click',ev=>{
  const c=ev.target.closest('.chip'); if(!c)return; SGROUP=c.dataset.sg; renderSpends();});
}
renderSpends();

document.addEventListener('click',e=>{
 const tab=e.target.closest('nav button');
 if(tab){document.querySelectorAll('nav button').forEach(b=>b.setAttribute('aria-selected',String(b===tab)));
  ['bdays','sbory','kids','spends'].forEach(n=>{
   document.getElementById('pane-'+n).hidden=(n!==tab.dataset.tab);
   const ex=document.getElementById('ex-'+n); if(ex)ex.hidden=(n!==tab.dataset.tab);});
  const a=document.getElementById('navtop');
  if(window.scrollY>a.offsetTop-8)window.scrollTo({top:a.offsetTop-8});
  return;}
 const btn=e.target.closest('.btn[data-toggle]');
 if(btn){const p=document.getElementById(btn.dataset.toggle);
  btn.setAttribute('aria-expanded',String(p.classList.toggle('open')));return;}
 const sc=e.target.closest('.chip.sc');
 if(sc){SORT=sc.dataset.sort;
  document.querySelectorAll('.chip.sc').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.sort===SORT)));
  const open=[...document.querySelectorAll('#pane-sbory .panel.open')].map(p=>p.id);
  renderSbory(); open.forEach(id=>{const p=document.getElementById(id);
   if(p){p.classList.add('open');const b=document.querySelector(`[data-toggle="${id}"]`);
    if(b)b.setAttribute('aria-expanded','true');}});
  document.getElementById('sortk').innerHTML=sortChips('kids');
  renderKids(document.getElementById('q').value);
  document.querySelectorAll('.chip.sc').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.sort===SORT)));
  return;}
 const rh=e.target.closest('.row-h');
 if(rh)rh.parentElement.classList.toggle('open');});

}

(function(){
 if(typeof RAW!=='string'||!RAW.startsWith('ENC:')){document.getElementById('gate').remove();
  document.getElementById('app').hidden=false; boot(RAW); return;}
 const g=document.getElementById('gate'), f=document.getElementById('gform'),
       inp=document.getElementById('gpass'), err=document.getElementById('gerr');
 f.addEventListener('submit', async ev=>{
  ev.preventDefault(); err.textContent=''; f.classList.add('busy');
  try{
   const D=await decryptPayload(RAW.slice(4), inp.value.trim());
   g.remove(); const app=document.getElementById('app'); app.hidden=false; boot(D);
  }catch(e){
   f.classList.remove('busy'); err.textContent='Неверное слово. Спросите у Ани.';
   inp.select();
  }});
 inp.focus();
})();
</script></body></html>
"""
PAGE = PAGE.replace("\u2014", "-").replace("\u2013", "-")
PAYLOAD = ("ENC:" + encrypt_payload(DATA, PASSWORD)) if PASSWORD else DATA
open(OUT, "w", encoding="utf-8").write(PAGE.replace("__DATA__", json.dumps(PAYLOAD, ensure_ascii=False) if PASSWORD else DATA)
                                        .replace("__ASOF__", AS_OF.strftime("%d.%m.%Y"))
                                        .replace("__TREAS__", "Анна Е." if MASK else "Анна Елисеева")
                                        .replace("__SEARCHPH__", "Найти по имени…" if MASK else "Найти по фамилии…")
                                        .replace("__STAMP__", STAMP)
                                        .replace("__YEAR__", str(AS_OF.year))
                                        .replace("__MONTH__", str(AS_OF.month))
                                        .replace("__DAY__", str(AS_OF.day)))
print("report ok ·", len(sbory), "сборов ·", len(kids), "детей")
