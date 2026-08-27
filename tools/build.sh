#!/usr/bin/env bash
#
# Сборка отчёта. Единственная точка входа: и публичная версия, и именная.
#
#   ./tools/build.sh                     публичная версия на сегодня -> index.html
#   ./tools/build.sh 2026-08-17          публичная версия задним числом
#   ./tools/build.sh 2026-08-17 --named  именная версия -> папка Свод на Диске
#
# Зачем скрипт: правило разрешения в .claude/settings.local.json записывается
# один раз как ./tools/build.sh * и не зависит от даты. Правило с зашитой датой
# протухало бы на следующий день, а без AS_OF/STAMP не покрывало пересборку
# задним числом.
#
# AS_OF и STAMP выводятся из одного аргумента, поэтому дата в заголовке и метка
# версии в подвале разойтись не могут.

set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd -P)"
cd "$REPO"

PY="$REPO/.venv/bin/python"
SRC="${SRC:-/Users/oleg/Library/CloudStorage/GoogleDrive-o017ev@gmail.com/Мой диск/For_sorting/Школа/Класс_Артура/2026-2027/Свод/Касса_2В_2026-2027.xlsx}"

# Именная версия живёт только на Диске: рядом с журналом, в той же папке Свод.
NAMED_NAME="Отчёт_именной_2В.html"
NAMED_DIR="$(dirname "$SRC")"
NAMED_OUT="$NAMED_DIR/$NAMED_NAME"

die() { printf '%s\n' "$@" >&2; exit 1; }

# --- аргументы ---------------------------------------------------------------

DATE=""
NAMED=0
for arg in "$@"; do
    case "$arg" in
        --named) NAMED=1 ;;
        -h|--help)
            sed -n '2,12p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
            exit 0 ;;
        -*) die "Неизвестный аргумент: $arg" \
                "Использование: ./tools/build.sh [ГГГГ-ММ-ДД] [--named]" ;;
        *)
            [ -n "$DATE" ] && die "Дата указана дважды: «$DATE» и «$arg»."
            DATE="$arg" ;;
    esac
done

[ -n "$DATE" ] || DATE="$(date +%F)"
[ -x "$PY" ] || die "Нет $PY." \
                    "Поставьте зависимости: python3 -m venv .venv && .venv/bin/pip install openpyxl cryptography"

# Дата проверяется календарно, а не шаблоном: 2026-02-31 подходит под ГГГГ-ММ-ДД,
# но такого дня нет. Отсюда же берётся STAMP, чтобы метки были согласованы.
STAMP="$("$PY" -c '
import datetime, sys
try:
    print(datetime.date.fromisoformat(sys.argv[1]).strftime("%d.%m.%Y"))
except ValueError:
    sys.exit("Дата %r: ожидается существующая дата в формате ГГГГ-ММ-ДД, например 2026-08-17." % sys.argv[1])
' "$DATE")" || exit 1

# --- куда пишется именная версия ---------------------------------------------

# Проверки пути идут до всего остального и не зависят от того, нашёлся ли журнал:
# это защита от публикации незашифрованного файла, а не подсказка про опечатку.
if [ "$NAMED" = 1 ]; then
    # В именной версии нет шифрования - данные лежат в файле открытым текстом.
    # (Маскировки нет и в публичной: с 27.08.2026 имена детей полные в обеих.)
    # Она законна ровно в одном месте: в папке Свод на Диске. Проверки ниже -
    # на случай, если SRC переопределили и путь уехал в рабочую копию.
    NAMED_DIR_REAL="$(cd "$NAMED_DIR" 2>/dev/null && pwd -P || printf '%s' "$NAMED_DIR")"

    case "$NAMED_DIR_REAL/" in
        "$REPO"/*|"$REPO"/)
            die "ОСТАНОВЛЕНО: именная версия писалась бы в репозиторий ($NAMED_DIR_REAL)." \
                "В файле незашифрованные данные детей - его место только на Диске." ;;
    esac
    case "$NAMED_DIR_REAL" in
        */CloudStorage/GoogleDrive-*/*) ;;
        *) die "ОСТАНОВЛЕНО: путь именной версии ведёт не на Google Диск ($NAMED_DIR_REAL)." \
               "В файле незашифрованные данные детей - его место только на Диске." ;;
    esac
    # Страховка от ошибки: если такой файл когда-нибудь окажется в рабочей копии,
    # он не должен попасть в git add -A.
    git check-ignore -q "$NAMED_NAME" \
        || die "ОСТАНОВЛЕНО: «$NAMED_NAME» не закрыт .gitignore." \
               "Добавьте строку «$NAMED_NAME» в .gitignore и повторите."

    [ -d "$NAMED_DIR_REAL" ] || die "Папка для именной версии не найдена: $NAMED_DIR"
fi

[ -f "$SRC" ] || die "Журнал не найден: $SRC" \
                     "Если Диск не примонтирован или терминалу не выдан полный доступ к диску - см. README."

# --- именная версия ----------------------------------------------------------

if [ "$NAMED" = 1 ]; then
    ALLOW_UNSAFE=1 AS_OF="$DATE" STAMP="$STAMP" SRC="$SRC" OUT="$NAMED_OUT" \
        "$PY" tools/build_report.py

    printf 'именная версия: %s\n' "$NAMED_OUT"
    printf 'verify не запускается: шифрования в именной версии нет по замыслу.\n'
    exit 0
fi

# --- публичная версия --------------------------------------------------------

# Код доступа не лежит в скрипте: скрипт коммитится, .env.local - нет.
ENV_FILE="$REPO/.env.local"
[ -f "$ENV_FILE" ] || die "Нет файла .env.local с кодом доступа." \
                          "Создайте $ENV_FILE строкой вида:" \
                          "    PASSWORD=кодовое-слово" \
                          "Файл закрыт .gitignore и в репозиторий не попадает."

PASSWORD="$(sed -n 's/^[[:space:]]*PASSWORD[[:space:]]*=[[:space:]]*//p' "$ENV_FILE" | head -1 | tr -d '\r')"
PASSWORD="${PASSWORD%\"}"; PASSWORD="${PASSWORD#\"}"
PASSWORD="${PASSWORD%\'}"; PASSWORD="${PASSWORD#\'}"
[ -n "$PASSWORD" ] || die "В $ENV_FILE нет непустой строки PASSWORD=..." \
                          "Пустой код собрал бы страницу в открытом виде - сборка остановлена."

PASSWORD="$PASSWORD" AS_OF="$DATE" STAMP="$STAMP" SRC="$SRC" OUT="$REPO/index.html" \
    "$PY" tools/build_report.py

printf '\n'
# SRC передаётся и в verify: одна из проверок ищет на странице имена технических
# строк журнала («Не ученик…» без денег) - без журнала ей неоткуда их взять.
if ! PASSWORD="$PASSWORD" SRC="$SRC" "$PY" tools/verify.py "$REPO/index.html"; then
    die "" "СБОРКА НЕ ПРИНЯТА: verify.py вернул ошибку. Такой файл не публикуют."
fi
