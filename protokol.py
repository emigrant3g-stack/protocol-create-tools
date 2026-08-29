#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
protokol.py — разбор и сборка протоколов совещаний ПО «ФОРЭНЕРГО».

Использование (<протокол.md> — файл любого протокола, имя произвольное):
    python protokol.py show    <протокол.md> [--date ДД.ММ.ГГГГ]
    python protokol.py short   <протокол.md>
    python protokol.py people  <протокол.md>
    python protokol.py md      <протокол.md>  --out <протокол.md>
    python protokol.py print   <протокол.md>  [--tpl шаблон.docx]
    python protokol.py renum   <протокол.md>  --out <протокол.md>

Облако (Google Диск через Apps Script), URL и секрет — из инструкций проекта:
    python protokol.py list  --url … --secret …
    python protokol.py pull  --name <протокол.md> --url … --secret …
    python protokol.py push  <протокол.md>        --url … --secret …

Если на совещании переносили сроки — перечислить эти номера:
    python protokol.py print <протокол.md> --moved 7,9.1

Имя файла ничем не ограничено: скрипт читает структуру из содержимого,
а не из названия. Протоколов может быть сколько угодно.

Шаблон печати скрипт при отсутствии рядом сам тянет с GitHub:
    https://github.com/emigrant3g-stack/protocol-create-tools  (ветка main)
Нет сети — скрипт скажет приложить файл.

Скрипт читает файл протокола в структуру (двумерный массив узлов),
считает просрочку и метки эскалации и выдаёт нужное представление.
Никаких «на глаз»: даты сравниваются арифметикой.
"""
import re, sys, argparse, datetime, copy, os, urllib.request, urllib.parse

# ---------- источники ----------
GH_RAW = "https://raw.githubusercontent.com/emigrant3g-stack/protocol-create-tools/main/"
TPL_NAME = "Шаблон_протокола_на_печать.docx"
TPL_URL = GH_RAW + urllib.parse.quote(TPL_NAME)


def ensure_tpl(path):
    """Вернуть путь к шаблону. Нет локально — попробовать скачать с GitHub.
    Нет сети — вернуть None, вызывающая сторона попросит приложить файл."""
    if path and os.path.exists(path):
        return path
    dst = path or TPL_NAME
    try:
        req = urllib.request.Request(TPL_URL, headers={"User-Agent": "protokol.py"})
        with urllib.request.urlopen(req, timeout=20) as r:
            data = r.read()
        if not data.startswith(b"PK"):
            return None
        open(dst, "wb").write(data)
        sys.stderr.write(f"шаблон загружен с GitHub: {len(data)} байт\n")
        return dst
    except Exception as e:
        sys.stderr.write(f"шаблон недоступен ({e.__class__.__name__}); приложите {TPL_NAME}\n")
        return None

# ---------- облако ----------
def cloud(url, secret, method="GET", payload=None, params=None, timeout=25):
    import json
    if method == "GET":
        q = urllib.parse.urlencode(dict(params or {}, secret=secret))
        req = urllib.request.Request(url + "?" + q)
    else:
        data = json.dumps(dict(payload or {}, secret=secret)).encode("utf-8")
        req = urllib.request.Request(url, data=data,
                                     headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


# ---------- модель ----------
# узел: dict(num, kind, mark, text, otv, srok)
#   kind: 'item'   — поручение (есть строка Ответственный/Срок)
#         'sect'   — раздел-заголовок (номер + название)
# mark: базовый счёт N0 (1 = метки нет)

MARKS = {1: "", 2: "Повторно!", 3: "В третий раз!", 4: "В четвёртый раз!",
         5: "В пятый раз!", 6: "В шестой раз!", 7: "В седьмой раз!",
         8: "В восьмой раз!", 9: "В девятый раз!", 10: "В десятый раз!"}
MARK_RE = re.compile(r"^(Повторно!|В (?:третий|четвёртый|пятый|шестой|седьмой|восьмой|девятый|десятый) раз!)\s*")
MARK_N = {v: k for k, v in MARKS.items() if v}

MONTHS = {"январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
          "июл": 7, "август": 8, "сентябр": 9, "октябр": 10,
          "ноябр": 11, "декабр": 12}
NO_DEADLINE = ("постоянно", "до результата", "до востребования", "пауза", "—", "-", "")


def last_day(y, m):
    return (datetime.date(y + (m == 12), (m % 12) + 1, 1) - datetime.timedelta(days=1)).day


def add_months(d, k):
    y, m = d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1
    return datetime.date(y, m, min(d.day, last_day(y, m)))


def add_workdays(d, k):
    while k > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5:
            k -= 1
    return d


WD = {"понедельн": 0, "вторн": 1, "сред": 2, "четверг": 3, "четвертг": 3,
      "пятниц": 4, "суббот": 5, "воскресен": 6}
NUMWORD = {"один": 1, "одну": 1, "два": 2, "две": 2, "три": 3, "четыре": 4,
           "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9,
           "десять": 10, "пару": 2, "полтора": 1, "полторы": 1}


def norm_srok(srok, base):
    """«через неделю», «в следующую среду» → «01.09.2026». База — дата совещания.
    Абсолютные и бессрочные формулировки возвращаются как есть."""
    raw = (srok or "").strip()
    s = raw.lower().replace("ё", "е")
    if s in [x.replace("ё", "е") for x in NO_DEADLINE]:
        return raw
    within = bool(re.match(r"^(?:в течение|в течении)\s+", s))
    s = re.sub(r"^(?:к|до|на|в течение|в течении)\s+", "", s).strip()
    s = re.sub(r"[.!]+$", "", s).strip()

    def out(d):
        return d.strftime("%d.%m.%Y")

    if s in ("сегодня", "сейчас"):            return out(base)
    if s == "завтра":                          return out(base + datetime.timedelta(days=1))
    if s == "послезавтра":                     return out(base + datetime.timedelta(days=2))

    def qty(txt, default=1):
        m = re.search(r"(\d+)", txt)
        if m:
            return int(m.group(1))
        for w, v in NUMWORD.items():
            if w in txt:
                return v
        return default

    # «через …» и «в течение …»
    rel = re.sub(r"^через\s+", "", s)
    if rel != s or within or re.match(
            r"^\d+\s*(календарн\w*\s*|рабоч\w*\s*)?(дн|недел|месяц|год|лет)", s):
        t = rel
        if re.search(r"рабоч\w*\s*дн", t):     return out(add_workdays(base, qty(t)))
        if re.search(r"\bдн(ей|я|ь)\b|\bсутки\b|\bсуток\b", t):
            return out(base + datetime.timedelta(days=qty(t)))
        if "полторы недел" in t:               return out(base + datetime.timedelta(days=10))
        if re.search(r"недел", t):             return out(base + datetime.timedelta(days=7 * qty(t)))
        if "полгода" in t or "пол года" in t:  return out(add_months(base, 6))
        if re.search(r"квартал", t):           return out(add_months(base, 3))
        if re.search(r"полтора\s*месяц", t):   return out(base + datetime.timedelta(days=45))
        if re.search(r"месяц", t):             return out(add_months(base, qty(t)))
        if re.search(r"\bгод|\bлет\b", t):     return out(add_months(base, 12 * qty(t)))

    # «к концу недели / месяца / квартала / года», «до конца …»
    if re.search(r"конц\w*\s+недел", s):
        return out(base + datetime.timedelta(days=(4 - base.weekday()) % 7))
    if re.search(r"конц\w*\s+месяц", s):
        return out(datetime.date(base.year, base.month, last_day(base.year, base.month)))
    if re.search(r"конц\w*\s+квартал", s):
        mo = ((base.month - 1) // 3 + 1) * 3
        return out(datetime.date(base.year, mo, last_day(base.year, mo)))
    if re.search(r"конц\w*\s+год", s):
        return out(datetime.date(base.year, 12, 31))

    # дни недели: «в следующую среду», «в среду», «к пятнице», «в ближайший вторник»
    for stem, wd in WD.items():
        if stem in s:
            step = (wd - base.weekday()) % 7
            if step == 0:
                step = 7
            d = base + datetime.timedelta(days=step)
            if re.search(r"следующ|через\s+недел", s):
                # понедельник следующей календарной недели + смещение
                nxt = base + datetime.timedelta(days=7 - base.weekday())
                d = nxt + datetime.timedelta(days=wd)
            return out(d)

    return raw


def deadline(srok, year):
    """Крайняя дата срока. None — срока нет (счётчик не двигается)."""
    s = (srok or "").strip().lower().replace("ё", "е")
    if s in [x.replace("ё", "е") for x in NO_DEADLINE]:
        return None
    s = s.replace("до ", "").strip()
    # диапазон: берём конец
    for sep in ("..", "…", "—", "–", "-"):
        if sep in s and not re.match(r"^\d{1,2}[.\-]\d{1,2}", s):
            s = s.split(sep)[-1].strip()
    m = re.match(r"^(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})$", s)          # 20.09.2026
    if m:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{1,2})[.\-](\d{1,2})$", s)                       # 03.09
    if m:
        return datetime.date(year, int(m.group(2)), int(m.group(1)))
    m = re.search(r"(\d)\s*квартал(?:\s*(\d{4}))?", s)                  # 3 квартал 2026
    if m:
        q = int(m.group(1)); y = int(m.group(2) or year); mo = q * 3
        return datetime.date(y, mo, last_day(y, mo))
    m = re.search(r"(\d)\s*полугодие(?:\s*(\d{4}))?", s)                # 2 полугодие
    if m:
        h = int(m.group(1)); y = int(m.group(2) or year); mo = h * 6
        return datetime.date(y, mo, last_day(y, mo))
    if "лето" in s:
        return datetime.date(year, 8, 31)
    for stem, mo in MONTHS.items():                                     # август/2026
        if stem in s:
            m2 = re.search(r"(\d{4})", s); y = int(m2.group(1)) if m2 else year
            return datetime.date(y, mo, last_day(y, mo))
    m = re.search(r"^(\d{4})", s)                                       # 2026 г.
    if m:
        y = int(m.group(1))
        return datetime.date(y, 12, 31)
    return None


def parse(path):
    txt = open(path, encoding="utf-8").read()
    head = {}
    m = re.search(r"^#\s*(.+)$", txt, re.M);            head["title"] = m.group(1).strip() if m else ""
    m = re.search(r"Дата последней редакции:\s*(.+)", txt); head["date"] = m.group(1).strip() if m else ""
    m = re.search(r"Удал[её]нные номера:\s*(.+)", txt)
    head["deleted"] = m.group(1).strip() if m else ""
    m = re.search(r"Использованные номера:.*?удалены:\s*([^)]*)\)", txt)   # старый формат
    if m and not head["deleted"]:
        d = m.group(1).strip()
        head["deleted"] = "" if d in ("—", "-", "") else d
    lines = txt.split("\n")
    place = ""
    for i, l in enumerate(lines):
        if l.startswith("Дата последней редакции:") and i + 1 < len(lines):
            place = lines[i + 1].strip()
    head["place"] = place
    head["block"] = ""

    body = re.search(r"```\n(.*?)```", txt, re.S)
    nodes = []
    if body:
        bl = body.group(1).split("\n")
        i = 0
        while i < len(bl):
            line = bl[i].strip()
            m = re.match(r"^(\d+(?:\.\d+)?)\.\s+(.*)$", line)
            if m:
                num, rest = m.group(1), m.group(2)
                mark = 1
                mm = MARK_RE.match(rest)
                if mm:
                    mark = MARK_N[mm.group(1)]; rest = rest[mm.end():]
                otv = srok = None
                if i + 1 < len(bl) and bl[i + 1].startswith("Ответственный:"):
                    f = bl[i + 1]
                    fm = re.match(r"Ответственный:\s*(.*?)\s{2,}Срок:\s*(.*)$", f)
                    if fm:
                        otv, srok = fm.group(1).strip(), fm.group(2).strip()
                    i += 1
                nodes.append(dict(num=num, kind="item" if otv is not None else "sect",
                                  mark=mark, text=rest.strip(), otv=otv, srok=srok))
            elif line and not line.startswith("Ответственный:"):
                if nodes:
                    nodes.append(dict(num="", kind="head", mark=1, text=line,
                                      otv=None, srok=None))
                else:
                    head["block"] = line
            i += 1
    return head, nodes


def escalate(nodes, today, moved=()):
    """Метка = N0 + 1, если пункт просрочен.
    Просрочен = срок в файле раньше сегодняшней даты ИЛИ номер указан в moved
    (срок перенесли на этом совещании — значит признали просрочку).
    Никакого «состояния на начало совещания» искать не нужно."""
    year = today.year
    moved = set(str(x).strip() for x in moved if str(x).strip())
    out = []
    for n in nodes:
        if n["kind"] != "item":
            out.append(0); continue
        d = deadline(n["srok"], year)
        over = (d is not None and today > d) or (n["num"] in moved)
        out.append(n["mark"] + (1 if over else 0))
    return out


def label(n):
    return MARKS.get(n, f"В {n}-й раз!")


def renumber(nodes):
    """Сквозная перенумерация. Подпункты едут вместе с родителем.
    Возвращает (новые узлы, таблица соответствия старый→новый)."""
    groups, index, pending = [], {}, []
    for n in nodes:
        if n["kind"] == "head":
            pending.append(n); continue
        top = n["num"].split(".")[0]
        if top not in index:
            index[top] = len(groups)
            groups.append({"top": top, "head": None, "subs": [], "pre": pending})
            pending = []
        g = groups[index[top]]
        if "." in n["num"]:
            g["subs"].append(n)
        else:
            g["head"] = n
    out, mapping = [], []
    for i, g in enumerate(groups, start=1):
        out += g.get("pre", [])
        if g["head"]:
            old = g["head"]["num"]
            g["head"] = dict(g["head"], num=str(i))
            out.append(g["head"]); mapping.append((old, str(i)))
        for j, sn in enumerate(g["subs"], start=1):
            old = sn["num"]; new = f"{i}.{j}"
            out.append(dict(sn, num=new)); mapping.append((old, new))
    return out, mapping


# ---------- представления ----------
def wrap(text, width, indent):
    words, lines, cur = text.split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 > width and cur:
            lines.append(cur); cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur: lines.append(cur)
    return [lines[0]] + [" " * indent + l for l in lines[1:]] if lines else [""]


def show(head, nodes, marks, full=True):
    out = [head["title"], f"Дата последней редакции: {head['date']}", head["place"]]
    if head.get("deleted"): out.append(f"Удалённые номера: {head['deleted']}")
    out.append("")
    if head["block"]: out += [head["block"], ""]
    for n, mk in zip(nodes, marks):
        if n["kind"] == "head":
            out += [n["text"], ""]; continue
        pre = (label(mk) + " ") if mk > 1 else ""
        body = pre + n["text"]
        if not full:
            t = (body[:60] + "…") if len(body) > 60 else body
            out.append(f"{n['num']+'.':>6} {t:<62} · {n['otv'] or ''} · {n['srok'] or ''}")
            continue
        ls = wrap(body, 64, 7)
        out.append(f"{n['num']+'.':>6} {ls[0]}")
        out += ls[1:]
        if n["kind"] == "item":
            out.append(" " * 7 + f"Ответственный: {n['otv']}")
            out.append(" " * 7 + f"Срок: {n['srok']}")
        out.append("")
    return "\n".join(out).rstrip()


def people(nodes):
    s = set()
    for n in nodes:
        if n["kind"] != "item" or not n["otv"]: continue
        raw = n["otv"].replace("руководители ДЗО:", ",").replace("(", ",").replace(")", ",")
        for p in raw.split(","):
            p = p.strip(" .,")
            if re.match(r"^[А-ЯЁ][а-яё\-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.$", p + ("" if p.endswith(".") else ".")):
                s.add(re.sub(r"\s+", " ", p if p.endswith(".") else p + "."))
    return sorted(s)


def to_md(head, nodes, marks, today):
    out = [f"# {head['title']}", "", f"Дата последней редакции: {today:%d.%m.%Y}", head["place"]]
    if head.get("deleted"): out.append(f"Удалённые номера: {head['deleted']}")
    out += ["", "```"]
    if head["block"]: out += [head["block"], ""]
    for n, mk in zip(nodes, marks):
        if n["kind"] == "head":
            out += [n["text"], ""]; continue
        pre = (label(mk) + " ") if mk > 1 else ""
        out.append(f"{n['num']}. {pre}{n['text']}")
        if n["kind"] == "item":
            out.append(f"Ответственный: {n['otv']}    Срок: {n['srok']}")
        out.append("")
    out.append("```")
    return "\n".join(out)


# ---------- печать ----------
def to_docx(head, nodes, marks, tpl, out_path, today):
    from docx import Document
    from docx.text.paragraph import Paragraph
    d = Document(tpl); P = d.paragraphs

    def find(sub):
        for p in P:
            if sub in p.text: return p
        raise KeyError(sub)

    def put(p, first, last=None, mark=""):
        """first — текст; mark — метка, печатается полужирным между номером и текстом."""
        txt = [r for r in p.runs if "\t" not in r.text]
        if not txt: return
        r0 = txt[0]
        if mark:
            num, rest = first.split(" ", 1)
            r0.text = num + " "; r0.bold = False
            e1 = copy.deepcopy(r0._element); r0._element.addnext(e1)
            e2 = copy.deepcopy(r0._element); e1.addnext(e2)
            rs = Paragraph(p._element, p._parent).runs
            rs[1].text = mark + " "; rs[1].bold = True
            rs[2].text = rest;       rs[2].bold = False
            rest_runs = [r for r in rs[3:] if "\t" not in r.text]
        else:
            r0.text = first
            rest_runs = txt[1:]
        for r in rest_runs:
            r.text = ""
        tt = [r for r in p.runs if "\t" not in r.text]
        if last is not None and len(tt) > 1:
            tt[-1].text = last

    TITLE = find("[НАЗВАНИЕ]"); DATE = find("[ДД.ММ.ГГГГ]"); PLACE = find("[Место]")
    HEAD = find("[ШАПКА БЛОКА.]"); ITEM = find("1. [Метка!]"); SECT = find("2. [Название раздела]")
    SUB = find("2.1. [Текст поручения.]"); SUB2 = find("2.2. [Текст поручения.]")
    fl = [p for p in P if p.text.startswith("Ответственный:")]
    FIELD, SUBF, SUB2F = fl[0], fl[1], fl[2]
    ISP = find("[Список исполнителей.]")
    par = ITEM._parent

    a = ITEM._element
    for n, mk in zip(nodes, marks):
        if n["kind"] == "head":
            e = copy.deepcopy(HEAD._element); a.addnext(e); a = e
            put(Paragraph(e, par), n["text"]); continue
        lvl = 1 if "." in n["num"] else 0
        tpl_p = SECT if n["kind"] == "sect" else (SUB if lvl else ITEM)
        e = copy.deepcopy(tpl_p._element); a.addnext(e); a = e
        pr = label(mk) if mk > 1 else ""
        put(Paragraph(e, par), f"{n['num']}. {n['text']}", mark=pr)
        if n["kind"] == "item":
            ef = copy.deepcopy((SUBF if lvl else FIELD)._element); a.addnext(ef); a = ef
            put(Paragraph(ef, par), f"Ответственный: {n['otv']}    ", f"Срок: {n['srok']}")

    b = ISP._element
    for name in people(nodes):
        e = copy.deepcopy(ISP._element); b.addnext(e); b = e
        put(Paragraph(e, par), name)

    for p in (ITEM, FIELD, SECT, SUB, SUBF, SUB2, SUB2F, ISP):
        p._element.getparent().remove(p._element)
    if head["block"]:
        put(HEAD, head["block"])
    else:
        HEAD._element.getparent().remove(HEAD._element)

    put(TITLE, re.sub(r"^\d+\.\s*", "", head["title"]).upper())
    put(DATE, f"Дата последней редакции: {today:%d.%m.%Y}")
    put(PLACE, head["place"])
    d.save(out_path)
    return out_path


def check(head, nodes):
    """Структурные аномалии. Список строк; пусто — всё в порядке."""
    out = []
    tops = {}
    nodes = [n for n in nodes if n["kind"] != "head"]
    for n in nodes:
        top = n["num"].split(".")[0]
        tops.setdefault(top, {"head": None, "subs": []})
        (tops[top]["subs"] if "." in n["num"] else tops[top]).__setitem__(
            len(tops[top]["subs"]), n) if False else None
    for n in nodes:
        top = n["num"].split(".")[0]
        if "." in n["num"]:
            tops[top]["subs"].append(n)
        else:
            tops[top]["head"] = n
    for top, g in tops.items():
        h, subs = g["head"], g["subs"]
        if h and h["kind"] == "item" and subs:
            out.append(f"{top}: поручение с ответственным и сроком, но имеет подпункты — заголовок это или пункт?")
        if h is None and subs:
            out.append(f"{top}: номер занят только подпунктами, самой записи «{top}.» нет")
        if h and h["kind"] == "sect" and not subs:
            out.append(f"{top}: раздел-заголовок без подпунктов")
    for n in nodes:
        if n["kind"] == "item":
            if not n["otv"]:
                out.append(f"{n['num']}: не указан ответственный")
            if not n["srok"]:
                out.append(f"{n['num']}: не указан срок")
    nums = [int(n["num"].split(".")[0]) for n in nodes]
    if nums:
        missing = sorted(set(range(1, max(nums) + 1)) - set(nums))
        if missing and not head.get("deleted"):
            out.append("дыры в нумерации: " + ", ".join(map(str, missing)) +
                       " — строки «Удалённые номера» нет, происхождение неизвестно")
    return out


def new_protocol(title, place, grif, today):
    t = title.strip()
    if grif and grif.strip() and grif.strip() != "-":
        t = f"{t} ({grif.strip()})"
    return "\n".join([f"# {t}", "", f"Дата последней редакции: {today:%d.%m.%Y}",
                       place.strip(), "", "```", "```", ""])


def merge(files, today):
    out = ["# МАСТЕР-ФАЙЛ ПРОТОКОЛОВ СОВЕЩАНИЙ", "",
           "**ПО «ФОРЭНЕРГО»** · производственное объединение",
           f"**Актуально на: {today:%d.%m.%Y}**", "", "## ОГЛАВЛЕНИЕ", "",
           "| № | Протокол | Дата |", "|---|---|---|"]
    bodies = []
    for i, p in enumerate(files, start=1):
        h, ns = parse(p)
        out.append(f"| {i} | {h['title']} | {h['date']} |")
        bodies.append(open(p, encoding="utf-8").read().strip())
    out.append("")
    for b in bodies:
        out += ["---", "", b, ""]
    return "\n".join(out)


def used_numbers(head, nodes):
    used = set()
    for n in nodes:
        if n["kind"] == "head": continue
        used.add(int(n["num"].split(".")[0]))
    for x in re.findall(r"\d+", head.get("deleted", "") or ""):
        used.add(int(x))
    return used


def next_top(head, nodes):
    u = used_numbers(head, nodes)
    return (max(u) + 1) if u else 1


def next_sub(nodes, parent):
    subs = [int(n["num"].split(".")[1]) for n in nodes
            if n["num"].startswith(parent + ".") and n["num"].count(".") == 1]
    return (max(subs) + 1) if subs else 1


BAK = ".protokol_bak"
JRN = ".protokol_journal"


def journal(path, line):
    with open(JRN, "a", encoding="utf-8") as j:
        j.write(line.rstrip() + "\n")


def backup(path):
    """Сохранить текущую версию файла перед изменением (для undo)."""
    if os.path.exists(path):
        open(BAK, "w", encoding="utf-8").write(open(path, encoding="utf-8").read())


def write(path, head, nodes, today, keep_date=True):
    """Записать файл, не трогая дату последней редакции (её ставит finish/md)."""
    backup(path)
    d = head.get("date") or f"{today:%d.%m.%Y}"
    out = [f"# {head['title']}", "", f"Дата последней редакции: {d}", head["place"]]
    if head.get("deleted"): out.append(f"Удалённые номера: {head['deleted']}")
    out += ["", "```"]
    if head.get("block"): out += [head["block"], ""]
    for n in nodes:
        if n["kind"] == "head":
            out += [n["text"], ""]; continue
        pre = (label(n["mark"]) + " ") if n["mark"] > 1 else ""
        out.append(f"{n['num']}. {pre}{n['text']}")
        if n["kind"] == "item":
            out.append(f"Ответственный: {n['otv']}    Срок: {n['srok']}")
        out.append("")
    out.append("```")
    open(path, "w", encoding="utf-8").write("\n".join(out) + "\n")


def pos_after(nodes, parent):
    """Индекс, куда вставить новый подпункт: после последнего подпункта parent."""
    last = None
    for i, n in enumerate(nodes):
        if n["num"] == parent or n["num"].startswith(parent + "."):
            last = i
    return (last + 1) if last is not None else len(nodes)


# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["show", "short", "people", "md", "print",
                                    "renum", "check", "finish", "new", "merge",
                                    "add", "set", "del", "head", "undo", "changes", "fix",
                                    "push", "pull", "list"])
    ap.add_argument("file", nargs="?", help="файл протокола .md (любое имя)")
    ap.add_argument("--files", nargs="*", default=[], help="для merge: список файлов протоколов")
    ap.add_argument("--title", default=""); ap.add_argument("--place", default="")
    ap.add_argument("--grif", default="")
    ap.add_argument("--num", default="", help="номер узла для set/del")
    ap.add_argument("--parent", default="", help="номер родителя для add: подпункт")
    ap.add_argument("--text", default=""); ap.add_argument("--otv", default="")
    ap.add_argument("--srok", default=""); ap.add_argument("--section", default="")
    ap.add_argument("--promote", default="", help="fix: вынести подпункт N.M в самостоятельный пункт")
    ap.add_argument("--to-section", dest="to_section", default="", help="fix: сделать пункт разделом-заголовком")
    ap.add_argument("--to-item", dest="to_item", default="", help="fix: сделать раздел пунктом (нужны --otv --srok)")
    ap.add_argument("--url", default=os.environ.get("PROTOKOL_URL", ""), help="URL веб-приложения Google Apps Script")
    ap.add_argument("--secret", default=os.environ.get("PROTOKOL_SECRET", ""), help="секрет шлюза")
    ap.add_argument("--name", default="", help="pull: имя файла в облаке")
    ap.add_argument("--tpl", default=TPL_NAME)
    ap.add_argument("--out")
    ap.add_argument("--date")
    ap.add_argument("--moved", default="", help="номера пунктов, у которых срок перенесён на этом совещании: --moved 7,9.1")
    a = ap.parse_args()
    today = (datetime.datetime.strptime(a.date, "%d.%m.%Y").date() if a.date
             else datetime.date.today())
    if a.cmd in ("push", "pull", "list"):
        if not a.url or not a.secret:
            print("НЕТ ДОСТУПА К ОБЛАКУ: нужны --url и --secret из инструкций проекта"); return
        try:
            if a.cmd == "list":
                r = cloud(a.url, a.secret)
                if not r.get("ok"): print("ОШИБКА:", r.get("error")); return
                for x in r["files"]:
                    print(f"{x['name']}   изменён {x['updated'][:10]}")
                return
            if a.cmd == "pull":
                name = a.name or a.file
                r = cloud(a.url, a.secret, params={"name": name})
                if not r.get("ok"): print("ОШИБКА:", r.get("error")); return
                open(name, "w", encoding="utf-8").write(r["content"])
                print(f"✔ загружен из облака: {name}")
                return
            body = open(a.file, encoding="utf-8").read()
            r = cloud(a.url, a.secret, "POST",
                      {"name": os.path.basename(a.file), "content": body})
            if not r.get("ok"): print("ОШИБКА:", r.get("error")); return
            print(f"✔ сохранён в облако: {r['name']} "
                  f"({'перезаписан' if r['replaced'] else 'создан'}, {r['bytes']} знаков)")
        except Exception as e:
            print(f"ОБЛАКО НЕДОСТУПНО ({e.__class__.__name__}: {e}). "
                  f"Сохраните файл вручную.")
        return

    if a.cmd == "new":
        dst = a.out or "Протокол.md"
        open(dst, "w", encoding="utf-8").write(new_protocol(a.title, a.place, a.grif, today))
        print("saved:", dst); return
    if a.cmd == "merge":
        dst = a.out or "Мастерфайл_Протоколы_ФОРЭНЕРГО.md"
        open(dst, "w", encoding="utf-8").write(merge(a.files or [a.file], today))
        print("saved:", dst); return

    head, nodes = parse(a.file)
    moved = [x for x in a.moved.split(",") if x.strip()]
    marks = escalate(nodes, today, moved)
    if a.cmd == "add":
        if a.parent:
            num = f"{a.parent}.{next_sub(nodes, a.parent)}"
            idx = pos_after(nodes, a.parent)
        else:
            num = str(next_top(head, nodes)); idx = len(nodes)
        if a.section:
            node = dict(num=num, kind="sect", mark=1, text=a.section.strip(), otv=None, srok=None)
        else:
            node = dict(num=num, kind="item", mark=1, text=a.text.strip(),
                        otv=(a.otv.strip() or "—"),
                        srok=(norm_srok(a.srok, today) or "—"))
        nodes.insert(idx, node)
        write(a.file, head, nodes, today)
        if node["kind"] == "sect":
            msg = f"+ {num} создан раздел «{node['text']}»"
        else:
            msg = f"+ {num} {node['text']} · {node['otv']} · {node['srok']}"
        journal(a.file, msg)
        print("✔ " + msg)
        return
    if a.cmd == "set":
        tgt = [n for n in nodes if n["num"] == a.num]
        if not tgt:
            print(f"НЕТ ПУНКТА {a.num}"); return
        n = tgt[0]; ch = []
        if a.text: n["text"] = a.text.strip(); ch.append("текст изменён")
        if a.otv:  n["otv"] = a.otv.strip();  ch.append(f"ответственный: {n['otv']}")
        if a.srok:
            old = n["srok"]; n["srok"] = norm_srok(a.srok, today)
            ch.append(f"срок {old} → {n['srok']}")
        write(a.file, head, nodes, today)
        msg = f"~ {a.num} " + ", ".join(ch)
        journal(a.file, msg)
        print("✔ " + msg)
        return
    if a.cmd == "del":
        gone = [n for n in nodes if n["num"] == a.num or n["num"].startswith(a.num + ".")]
        if not gone:
            print(f"НЕТ ПУНКТА {a.num}"); return
        nodes = [n for n in nodes if n not in gone]
        top = int(a.num.split(".")[0])
        if "." not in a.num and top == max(used_numbers(head, nodes + gone)):
            d = [x for x in re.findall(r"\d+", head.get("deleted", "") or "")]
            head["deleted"] = ", ".join(sorted(set(d + [str(top)]), key=int))
        write(a.file, head, nodes, today)
        left = ", ".join(n["num"] for n in nodes if n["kind"] != "head")
        msg = f"− {a.num} удалён · без изменений: {left}"
        journal(a.file, msg)
        print("✔ " + msg)
        return

    if a.cmd == "head":
        ch = []
        if a.title: head["title"] = a.title.strip(); ch.append("название")
        if a.place: head["place"] = a.place.strip(); ch.append("место")
        if a.grif:
            base = re.sub(r"\s*\([^)]*\)\s*$", "", head["title"]).strip()
            g = a.grif.strip()
            head["title"] = base if g in ("-", "нет", "") else f"{base} ({g})"
            ch.append("гриф")
        write(a.file, head, nodes, today)
        msg = "шапка: " + ", ".join(ch) + f" → {head['title']}, {head['place']}"
        journal(a.file, msg); print("✔ " + msg)
        return
    if a.cmd == "undo":
        if not os.path.exists(BAK):
            print("НЕЧЕГО ОТМЕНЯТЬ"); return
        cur = open(a.file, encoding="utf-8").read()
        open(a.file, "w", encoding="utf-8").write(open(BAK, encoding="utf-8").read())
        open(BAK, "w", encoding="utf-8").write(cur)      # повторный undo вернёт обратно
        lines = []
        if os.path.exists(JRN):
            lines = open(JRN, encoding="utf-8").read().splitlines()
        last = lines.pop() if lines else ""
        open(JRN, "w", encoding="utf-8").write("\n".join(lines) + ("\n" if lines else ""))
        print("✔ отменено: " + (last or "последнее изменение"))
        return
    if a.cmd == "changes":
        if not os.path.exists(JRN) or not open(JRN, encoding="utf-8").read().strip():
            print("Изменений за это совещание нет"); return
        print(open(JRN, encoding="utf-8").read().rstrip())
        return

    if a.cmd == "show" and a.num:
        sel = [(n, m) for n, m in zip(nodes, marks)
               if n["num"] == a.num or n["num"].startswith(a.num + ".")]
        if not sel:
            print(f"НЕТ ПУНКТА {a.num}"); return
        print(show(head, [n for n, _ in sel], [m for _, m in sel], True))
        return
    if a.cmd == "fix":
        if a.promote:
            tgt = [n for n in nodes if n["num"] == a.promote]
            if not tgt: print(f"НЕТ ПУНКТА {a.promote}"); return
            n = tgt[0]; old = n["num"]; n["num"] = str(next_top(head, nodes))
            nodes.remove(n); nodes.append(n)
            write(a.file, head, nodes, today)
            msg = f"структура: {old} → {n['num']} (вынесен в самостоятельный пункт)"
        elif a.to_section:
            tgt = [n for n in nodes if n["num"] == a.to_section]
            if not tgt: print(f"НЕТ ПУНКТА {a.to_section}"); return
            n = tgt[0]; n["kind"] = "sect"; n["otv"] = None; n["srok"] = None
            write(a.file, head, nodes, today)
            msg = f"структура: {n['num']} → раздел-заголовок (поля убраны)"
        elif a.to_item:
            tgt = [n for n in nodes if n["num"] == a.to_item]
            if not tgt: print(f"НЕТ ПУНКТА {a.to_item}"); return
            n = tgt[0]; n["kind"] = "item"
            n["otv"] = a.otv.strip() or "—"; n["srok"] = a.srok.strip() or "—"
            write(a.file, head, nodes, today)
            msg = f"структура: {n['num']} → пункт · {n['otv']} · {n['srok']}"
        else:
            print("УКАЖИТЕ --promote N.M | --to-section N | --to-item N"); return
        journal(a.file, msg); print("✔ " + msg)
        print("Проверьте нумерацию: при необходимости выполните renum.")
        return

    if a.cmd == "show":   print(show(head, nodes, marks, True))
    elif a.cmd == "short":print(show(head, nodes, marks, False))
    elif a.cmd == "people":print("\n".join(people(nodes)))
    elif a.cmd == "md":
        t = to_md(head, nodes, marks, today)
        if a.out: open(a.out, "w", encoding="utf-8").write(t + "\n"); print("saved:", a.out)
        else: print(t)
    elif a.cmd == "renum":
        new_nodes, mapping = renumber(nodes)
        head["deleted"] = ""   # после перенумерации дыр нет
        marks2 = escalate(new_nodes, today, moved)
        changed = [(o, n) for o, n in mapping if o != n]
        print("ПЕРЕНУМЕРАЦИЯ: изменено номеров — %d из %d" % (len(changed), len(mapping)))
        for o, n in changed:
            print(f"  {o} → {n}")
        if not changed:
            print("  нечего менять, нумерация уже сплошная")
        t = to_md(head, new_nodes, marks2, today)
        if a.out:
            open(a.out, "w", encoding="utf-8").write(t + "\n")
            print("saved:", a.out)
        if changed:
            print("Нумерация изменена — распечатайте протокол заново.")
    elif a.cmd == "check":
        issues = check(head, nodes)
        print("\n".join("• " + x for x in issues) if issues else "Структура в порядке")
    elif a.cmd == "finish":
        md_path = a.file
        open(md_path, "w", encoding="utf-8").write(to_md(head, nodes, marks, today) + "\n")
        print("saved md:", md_path)
        tpl = ensure_tpl(a.tpl)
        if not tpl:
            print(f"НЕТ ШАБЛОНА. Приложите {TPL_NAME}."); return
        head2, nodes2 = parse(md_path)
        marks2 = escalate(nodes2, today, moved)
        out = a.out or f"Протокол_{head2['title'].split('.')[0]}_{today:%d.%m.%Y}.docx"
        print("saved docx:", to_docx(head2, nodes2, marks2, tpl, out, today))
        print("ОТДАТЬ РУКОВОДИТЕЛЮ ОБА ФАЙЛА:", md_path, "и", out)
    elif a.cmd == "print":
        tpl = ensure_tpl(a.tpl)
        if not tpl:
            print(f"НЕТ ШАБЛОНА. Приложите {TPL_NAME} к чату.")
            return
        out = a.out or f"Протокол_{head['title'].split('.')[0]}_{today:%d.%m.%Y}.docx"
        print("saved:", to_docx(head, nodes, marks, tpl, out, today))


if __name__ == "__main__":
    main()
