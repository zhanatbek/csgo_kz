#!/usr/bin/env python3
"""Выгрузка базы нотариусов с enis.kz (раздел «Нотариусы по палатам»).

Примеры:
    python enis_notary.py                      # палата 4 -> notaries_chamber_4.xlsx/.csv
    python enis_notary.py --chambers 1-20      # несколько палат в один файл
    python enis_notary.py --chambers 4 --details   # заходить в карточку каждого нотариуса
    python enis_notary.py --chambers 4 --debug     # сохранить сырой HTML в ./debug_html

Скрипт не привязан к жёсткой вёрстке: берёт данные из HTML-таблиц, а если их нет,
то из карточек (блоков с телефоном/адресом). Пагинацию находит по ссылкам с page=.
"""
import argparse
import csv
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

BASE = "https://enis.kz"
LIST_URL = BASE + "/Notary/NotaryByChamber/{chamber}"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9",
}
PHONE_RE = re.compile(r"(?:\+?7|8)[\s\-()]*\d{3}[\s\-()]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
DETAIL_HINT = re.compile(r"/Notary/(?!NotaryByChamber)[^?#]*\d", re.I)


class Client:
    def __init__(self, delay, debug_dir=None):
        self.s = requests.Session()
        self.s.headers.update(HEADERS)
        self.delay = delay
        self.debug_dir = debug_dir
        self.n = 0

    def get(self, url, tries=4):
        for i in range(tries):
            try:
                r = self.s.get(url, timeout=30)
                if r.status_code == 404:
                    return None
                r.raise_for_status()
                r.encoding = r.apparent_encoding or "utf-8"
                time.sleep(self.delay)
                if self.debug_dir:
                    self.n += 1
                    (self.debug_dir / f"{self.n:04d}.html").write_text(
                        f"<!-- {url} -->\n" + r.text, encoding="utf-8")
                return r.text
            except requests.RequestException as e:
                wait = 2 ** (i + 1)
                print(f"  ! {url}: {e}; повтор через {wait}s", file=sys.stderr)
                time.sleep(wait)
        print(f"  ! не удалось загрузить {url}", file=sys.stderr)
        return None


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def set_page(url, page):
    p = urlparse(url)
    q = parse_qs(p.query)
    q["page"] = [str(page)]
    return urlunparse(p._replace(query=urlencode(q, doseq=True)))


def max_page(soup):
    """Максимальный номер страницы из ссылок пагинации (page=N или /N в .pagination)."""
    pages = [1]
    for a in soup.find_all("a", href=True):
        m = re.search(r"[?&](?:page|p|pageNumber)=(\d+)", a["href"], re.I)
        if m:
            pages.append(int(m.group(1)))
        elif a.find_parent(class_=re.compile("pag", re.I)) and clean(a.get_text()).isdigit():
            pages.append(int(clean(a.get_text())))
    return max(pages)


def parse_tables(soup, page_url):
    rows = []
    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if len(trs) < 2:
            continue
        head = [clean(c.get_text()) for c in trs[0].find_all(["th", "td"])]
        if not any(head):
            continue
        for tr in trs[1:]:
            cells = tr.find_all(["td", "th"])
            if not cells:
                continue
            rec = {}
            for i, c in enumerate(cells):
                key = head[i] if i < len(head) and head[i] else f"col{i + 1}"
                rec[key] = clean(c.get_text(" "))
            link = tr.find("a", href=True)
            if link:
                rec["Ссылка"] = urljoin(page_url, link["href"])
            if any(rec.values()):
                rows.append(rec)
    return rows


def label_values(node):
    """Пары «Метка: значение» из dl/dt, th/td и строк вида 'Телефон: ...'."""
    rec = {}
    for dt in node.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            rec[clean(dt.get_text()).rstrip(":")] = clean(dd.get_text(" "))
    for tr in node.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if len(cells) == 2:
            rec[clean(cells[0].get_text()).rstrip(":")] = clean(cells[1].get_text(" "))
    for line in node.get_text("\n").split("\n"):
        m = re.match(r"\s*([А-ЯЁA-Zа-яёa-zӘәҒғҚқҢңӨөҰұҮүҺһІі .№/-]{2,40}):\s*(.+)", line)
        if m:
            rec.setdefault(clean(m.group(1)), clean(m.group(2)))
    return rec


def enrich(rec, text):
    phones = sorted(set(clean(p) for p in PHONE_RE.findall(text)))
    emails = sorted(set(EMAIL_RE.findall(text)))
    if phones:
        rec.setdefault("Телефоны", ", ".join(phones))
    if emails:
        rec.setdefault("Email", ", ".join(emails))
    return rec


def parse_cards(soup, page_url):
    """Если таблиц нет: ищем повторяющиеся блоки с телефоном или ссылкой на нотариуса."""
    candidates = []
    for el in soup.find_all(["div", "li", "article"]):
        text = el.get_text(" ")
        has_link = el.find("a", href=DETAIL_HINT)
        if not (PHONE_RE.search(text) or has_link):
            continue
        # берём самый «внутренний» блок, содержащий ровно одну запись
        inner = [c for c in el.find_all(["div", "li", "article"])
                 if PHONE_RE.search(c.get_text(" ")) or c.find("a", href=DETAIL_HINT)]
        if inner:
            continue
        candidates.append(el)
    rows = []
    for el in candidates:
        rec = label_values(el)
        title = el.find(["h1", "h2", "h3", "h4", "h5", "strong", "b"])
        if title:
            rec.setdefault("ФИО", clean(title.get_text()))
        link = el.find("a", href=DETAIL_HINT) or el.find("a", href=True)
        if link:
            rec.setdefault("ФИО", clean(link.get_text()))
            rec["Ссылка"] = urljoin(page_url, link["href"])
        rec.setdefault("Текст", clean(el.get_text(" ")))
        rows.append(enrich(rec, el.get_text(" ")))
    return rows


def parse_list(html, page_url):
    soup = BeautifulSoup(html, "html.parser")
    rows = parse_tables(soup, page_url) or parse_cards(soup, page_url)
    return rows, soup


def scrape_chamber(client, chamber, details):
    url = LIST_URL.format(chamber=chamber)
    html = client.get(url)
    if not html:
        return []
    rows, soup = parse_list(html, url)
    chamber_name = clean((soup.find("h1") or soup.find("h2") or soup.new_tag("x")).get_text())
    last = max_page(soup)
    seen = {tuple(sorted(r.items())) for r in rows}
    page = 2
    # идём по страницам, пока появляются новые записи (last — подсказка, не ограничение)
    while page <= max(last, 2):
        page_html = client.get(set_page(url, page))
        if not page_html:
            break
        new, psoup = parse_list(page_html, url)
        new = [r for r in new if tuple(sorted(r.items())) not in seen]
        if not new:
            break
        seen.update(tuple(sorted(r.items())) for r in new)
        rows += new
        last = max(last, max_page(psoup), page + 1 if page >= last else last)
        print(f"  палата {chamber}: стр. {page}, всего {len(rows)}")
        page += 1

    if details:
        for i, r in enumerate(rows, 1):
            link = r.get("Ссылка")
            if not link or urlparse(link).netloc not in ("", urlparse(BASE).netloc):
                continue
            dh = client.get(link)
            if dh:
                dsoup = BeautifulSoup(dh, "html.parser")
                main = dsoup.find("main") or dsoup.find(id=re.compile("content|main", re.I)) or dsoup.body or dsoup
                for k, v in enrich(label_values(main), main.get_text(" ")).items():
                    r.setdefault(k, v)
            if i % 20 == 0:
                print(f"  палата {chamber}: карточки {i}/{len(rows)}")

    for r in rows:
        r["Палата ID"] = chamber
        r["Палата"] = chamber_name
        if len(r) > 4:
            r.pop("Текст", None)
    print(f"Палата {chamber} ({chamber_name}): {len(rows)} записей")
    return rows


def parse_chambers(spec):
    ids = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            ids += range(int(a), int(b) + 1)
        else:
            ids.append(int(part))
    return ids


def save(rows, out):
    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    front = ["Палата ID", "Палата"]
    cols = front + [c for c in cols if c not in front]
    csv_path = out.with_suffix(".csv")
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, delimiter=";")
        w.writeheader()
        w.writerows(rows)
    print(f"CSV:  {csv_path}")
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Нотариусы"
        ws.append(cols)
        for r in rows:
            ws.append([r.get(c, "") for c in cols])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                60, max(10, max(len(str(c.value or "")) for c in col[:200]) + 2))
        xlsx_path = out.with_suffix(".xlsx")
        wb.save(xlsx_path)
        print(f"XLSX: {xlsx_path}")
    except ImportError:
        print("openpyxl не установлен — сохранён только CSV")


def main():
    ap = argparse.ArgumentParser(description="Выгрузка нотариусов с enis.kz")
    ap.add_argument("--chambers", default="4", help="ID палат: 4 | 1,4,7 | 1-20")
    ap.add_argument("--details", action="store_true", help="заходить в карточку каждого нотариуса")
    ap.add_argument("--delay", type=float, default=1.0, help="пауза между запросами, сек")
    ap.add_argument("--out", default=None, help="имя файла без расширения")
    ap.add_argument("--debug", action="store_true", help="сохранять HTML в ./debug_html")
    a = ap.parse_args()

    debug_dir = None
    if a.debug:
        debug_dir = Path("debug_html")
        debug_dir.mkdir(exist_ok=True)
    client = Client(a.delay, debug_dir)
    ids = parse_chambers(a.chambers)
    rows = []
    for cid in ids:
        rows += scrape_chamber(client, cid, a.details)
    if not rows:
        print("Ничего не найдено. Запустите с --debug и посмотрите HTML в ./debug_html", file=sys.stderr)
        sys.exit(1)
    out = Path(a.out or f"notaries_chamber_{a.chambers.replace(',', '_')}")
    save(rows, out)
    print(f"Готово: {len(rows)} записей")


if __name__ == "__main__":
    main()
