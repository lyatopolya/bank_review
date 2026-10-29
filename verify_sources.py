#!/usr/bin/env python3
"""Ежедневная сверка ставок программ банков со страницами-источниками.

Скрипт НИЧЕГО не подставляет в базу. Он только отвечает на вопрос «ставка из базы всё ещё
стоит на странице-источнике?» и записывает ответ в data/verification.json:

  ok          ставка на странице совпала с базой  -> программа считается проверенной сегодня
  changed     на странице другая ставка          -> программа помечается как устаревшая
  unreadable  страницу не удалось прочитать или ставки в тексте нет (например, её рисует
              JavaScript) -> дата проверки не обновляется
  blocked     robots.txt запрещает чтение страницы -> не запрашиваем
              (если robots.txt вообще не прочитался, статус unreadable: без него страницы не читаем)
  unsupported источник не из поддерживаемых (читаем только ru.myfin.by)

Правила вежливости: пауза между запросами, понятный User-Agent, проверка robots.txt.
"""
import argparse
import datetime as dt
import json
import re
import sys
import time
from html.parser import HTMLParser
from pathlib import Path
from urllib import error, parse, request, robotparser

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import update_data  # noqa: E402

USER_AGENT = "RealtyRateChecker/1.0 (rate freshness check for a real estate calculator site)"
SUPPORTED_HOSTS = {"ru.myfin.by"}
DELAY_SECONDS = 3
MIN_TEXT_LEN = 1000
TOLERANCE = 0.005


class _Text(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._skip += 1

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript") and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        if not self._skip and data.strip():
            self.parts.append(data.strip())


def visible_text(html: str) -> str:
    p = _Text()
    p.feed(html)
    return re.sub(r"\s+", " ", " ".join(p.parts))


_RATE_RE = re.compile(r"ставк\w*[^%\d]{0,40}?(\d{1,2}(?:[.,]\d{1,3})?)\s*%", re.I)


def rate_candidates(text: str):
    """Все проценты, стоящие сразу после слова «ставка» (до 40 символов без цифр между ними)."""
    out = []
    for m in _RATE_RE.finditer(text):
        v = round(float(m.group(1).replace(",", ".")), 3)
        if v not in out:
            out.append(v)
    return out


def judge(text: str, base_rate_pct: float):
    """Вернуть (status, found_rate_pct) по тексту страницы."""
    if len(text) < MIN_TEXT_LEN:
        return "unreadable", None
    cands = rate_candidates(text)
    if not cands:
        return "unreadable", None
    if any(abs(c - base_rate_pct) <= TOLERANCE for c in cands):
        return "ok", base_rate_pct
    return "changed", cands[0]


def _fetch(url: str, timeout: int = 30) -> str:
    req = request.Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "ru"})
    with request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="replace")


def robots_allows(url: str, fetch=_fetch, cache=None) -> bool:
    host = parse.urlsplit(url)
    root = f"{host.scheme}://{host.netloc}"
    cache = cache if cache is not None else {}
    if root not in cache:
        rp = robotparser.RobotFileParser()
        try:
            rp.parse(fetch(root + "/robots.txt").splitlines())
            cache[root] = rp
        except error.HTTPError as e:
            # нет файла (404) = ограничений нет; остальные коды читаем как «нельзя»
            if e.code == 404:
                rp.parse([])
                cache[root] = rp
            else:
                cache[root] = None
        except Exception:
            cache[root] = None
    rp = cache[root]
    if rp is None:
        return None  # robots.txt не удалось прочитать: не знаем, можно ли
    return rp.can_fetch(USER_AGENT, url)


def verify_one(url: str, base_rate_pct, today: dt.date, fetch=_fetch, robots_cache=None):
    entry = {"checked_at": today.isoformat(), "status": "unreadable", "found_rate_pct": None}
    host = parse.urlsplit(url or "").netloc
    if host not in SUPPORTED_HOSTS:
        entry["status"] = "unsupported"
        return entry
    if base_rate_pct is None:
        return entry
    allowed = robots_allows(url, fetch, robots_cache)
    if allowed is None:
        entry["error"] = "robots.txt недоступен, страницу не запрашиваем"
        return entry
    if not allowed:
        entry["status"] = "blocked"
        return entry
    try:
        text = visible_text(fetch(url))
    except Exception as e:  # сеть, 4xx/5xx, кодировка
        entry["error"] = str(e)[:120]
        return entry
    entry["status"], entry["found_rate_pct"] = judge(text, base_rate_pct)
    return entry


def run(xlsx: Path, out: Path, today: dt.date, fetch=_fetch, delay: float = DELAY_SECONDS):
    programs, _ = update_data.load_programs(xlsx, today)
    result, cache, first = {}, {}, True
    for p in programs:
        url = p.get("source_url")
        if not url or url in result:
            continue
        if parse.urlsplit(url).netloc in SUPPORTED_HOSTS and not first:
            time.sleep(delay)
        first = False
        result[url] = verify_one(url, p.get("rate_min_pct"), today, fetch, cache)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps({"verified_on": today.isoformat(), "sources": result},
                              ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default=str(HERE / "baza-bankov.xlsx"))
    ap.add_argument("--out", default=str(HERE / "data" / "verification.json"))
    a = ap.parse_args(argv)
    res = run(Path(a.xlsx), Path(a.out), dt.date.today())
    counts = {}
    for e in res.values():
        counts[e["status"]] = counts.get(e["status"], 0) + 1
    print("Сверка источников:", ", ".join(f"{k}: {v}" for k, v in sorted(counts.items())) or "нечего проверять")
    for url, e in res.items():
        if e["status"] == "changed":
            print(f"  ставка изменилась: {url} -> {e['found_rate_pct']}%")
    return 0


if __name__ == "__main__":
    sys.exit(main())
