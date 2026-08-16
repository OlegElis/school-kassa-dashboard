#!/usr/bin/env python3
"""Проверка собранной страницы: расшифровать payload, показать итоги, поймать утечку.

Запуск:

    PASSWORD='код' .venv/bin/python tools/verify.py index.html

Смысл в том, чтобы смотреть на то, что реально лежит в файле, а не на то, что
генератор напечатал в конце сборки. Расхождение между журналом и страницей
ловится глазами по итогам, а утечка персональных данных - проверками ниже.

Выход 0 - всё сошлось, 1 - хотя бы одна проверка упала.
"""

import base64
import json
import os
import re
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

HERE = os.path.dirname(os.path.abspath(__file__))
GEN = os.path.join(HERE, "build_report.py")

# Публичная форма имени - «Артур Е.»: имя, пробел, одна заглавная буква с точкой.
# Полная фамилия («Елисеев Артур») этот шаблон не проходит - на том и проверка.
MASKED = re.compile(r"^\S+ [А-ЯЁA-Z]\.$")

# Ключи payload, которые в файле обязаны быть только внутри шифротекста.
# Взяты те, что не встречаются в вёрстке и скриптах страницы: «"spends"» сюда
# не годится - так называется вкладка (data-tab="spends"), проверка ловила бы себя.
PLAIN_MARKERS = ('"debtY"', '"dueYear"', '"bday"', '"proof"', '"collected"')

ok = True


def check(cond, good, bad):
    """Печатает строку отчёта и запоминает провал, не прерывая остальные проверки."""
    global ok
    print(f"  {'ok  ' if cond else 'FAIL'}  {good if cond else bad}")
    if not cond:
        ok = False


def norm_pass(p):
    """Та же нормализация, что в build_report.py: иначе выведется другой ключ."""
    return p.strip().lower()


def gen_iters():
    """Число итераций из генератора - источник правды для страницы."""
    m = re.search(r"^PBKDF2_ITERS\s*=\s*(\d+)", open(GEN, encoding="utf-8").read(), re.M)
    if not m:
        sys.exit("не нашёл PBKDF2_ITERS в tools/build_report.py")
    return int(m.group(1))


def rub(n):
    """Число в том же виде, в каком его видит родитель: 9 810,5."""
    s = f"{round(n, 2):,.2f}".replace(",", " ").replace(".", ",")
    return s[:-3] if s.endswith(",00") else s


def decrypt(blob_b64, password, iters):
    blob = base64.b64decode(blob_b64)
    salt, iv, ct = blob[:16], blob[16:28], blob[28:]
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                     iterations=iters).derive(norm_pass(password).encode())
    return json.loads(AESGCM(key).decrypt(iv, ct, None))


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "index.html"
    password = os.environ.get("PASSWORD", "")
    if not password:
        sys.exit("нужен PASSWORD - без кода payload не расшифровать")

    html = open(path, encoding="utf-8").read()
    iters = gen_iters()

    print(f"\n{path}\n")
    print("Шифрование")
    m = re.search(r'const RAW="ENC:([A-Za-z0-9+/=]+)"', html)
    check(m is not None, "данные зашифрованы (RAW=\"ENC:...\")",
          "данных в зашифрованном виде нет - страницу нельзя публиковать")
    if not m:
        sys.exit(1)

    page_iters = [int(x) for x in re.findall(r"iterations:(\d+)", html)]
    check(page_iters and all(i == iters for i in page_iters),
          f"итерации PBKDF2 совпадают с генератором ({iters})",
          f"итерации разошлись: в генераторе {iters}, на странице {page_iters}")

    data = decrypt(m.group(1), password, iters)
    kids, sbory, spends, t = data["kids"], data["sbory"], data["spends"], data["t"]
    # Внешние плательщики (ext) лежат в том же списке, но детьми класса не являются:
    # расходы делятся не на них, и в счёт детей они не идут.
    own = [k for k in kids if not k.get("ext")]
    ext = [k for k in kids if k.get("ext")]

    print("\nСодержимое")
    print(f"        детей: {len(own)}")
    if ext:
        print(f"        внешних плательщиков: {len(ext)}")
    print(f"        сборов: {len(sbory)}")
    print(f"        расходов: {len(spends)}")
    print(f"        собрано: {rub(t['collected'])} ₽")
    print(f"        потрачено: {rub(t['spent'])} ₽")
    print(f"        остаток: {rub(t['rest'])} ₽")
    if own:
        print(f"        доля расходов: {rub(t['spent'] / len(own))} ₽ на ребёнка")

    print("\nАрифметика")
    check(abs(t["collected"] - t["spent"] - t["rest"]) < 0.005,
          "собрано − потрачено = остаток",
          f"{rub(t['collected'])} − {rub(t['spent'])} ≠ {rub(t['rest'])}")
    total = sum(s["amount"] for s in spends)
    check(abs(total - t["spent"]) < 0.005,
          f"сумма {len(spends)} расходов сходится с «потрачено»",
          f"расходы дают {rub(total)}, а в итогах {rub(t['spent'])}")
    if own:
        share = t["spent"] / len(own)
        bad = [k for k in own if abs(k.get("share", 0) - share) > 0.005]
        check(not bad, f"доля расходов у всех детей одна ({rub(share)} ₽)",
              f"доля расходов разъехалась у {len(bad)} детей")

    print("\nПерсональные данные")
    # Имена внешних плательщиков - не фамилии детей («Соседи»), под шаблон маски
    # они не подходят и проверяются не здесь. Собранное без MASK=1 всё равно
    # ловится: у детей фамилии остались бы полными.
    ext_names = {k["name"] for k in ext}
    names = [k["name"] for k in own]
    names += [p["name"] for s in sbory for p in s.get("parts", []) if p["name"] not in ext_names]
    leaked = sorted({n for n in names if not MASKED.match(n)})
    check(not leaked, f"все имена в payload замаскированы ({len(names)} шт.)",
          f"незамаскированных имён: {len(leaked)} - собрано без MASK=1")

    treas = re.search(r"казначей ([^<]+)</div>", html)
    check(treas is not None and MASKED.match(treas.group(1).strip()) is not None,
          "имя казначея замаскировано",
          f"имя казначея в открытом виде: {treas.group(1).strip() if treas else '—'}")

    found = [mk for mk in PLAIN_MARKERS if mk in html]
    check(not found, "открытых данных payload в файле нет",
          f"payload лежит в открытом виде: {', '.join(found)}")

    print("\n" + ("verify ok" if ok else "verify FAILED") + "\n")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
