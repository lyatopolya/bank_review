import datetime as dt
import sys
import unittest
from pathlib import Path
from urllib import error

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import verify_sources as vs
import update_data as ud

TODAY = dt.date(2026, 10, 29)
URL = "https://ru.myfin.by/bank/alfabank/ipoteka/kredit-na-kvartiru"
FILLER = "<p>" + "Описание программы ипотеки. " * 60 + "</p>"


def page(rate_text):
    return f"<html><head><script>var x='Ставка 99%'</script></head><body><h1>Ипотека</h1>{FILLER}<div>{rate_text}</div></body></html>"


def fetcher(body, robots="User-agent: *\nAllow: /"):
    def f(url, timeout=30):
        if url.endswith("/robots.txt"):
            if robots is None:
                raise error.HTTPError(url, 404, "nf", {}, None)
            return robots
        if isinstance(body, Exception):
            raise body
        return body
    return f


class JudgeTests(unittest.TestCase):
    def test_same_rate_is_ok(self):
        t = vs.visible_text(page("Ставка от 17.99% Условия ставки: скидка -0,4%"))
        self.assertEqual(vs.judge(t, 17.99), ("ok", 17.99))

    def test_comma_format_is_ok(self):
        t = vs.visible_text(page("Ставка от: 17,99 %"))
        self.assertEqual(vs.judge(t, 17.99)[0], "ok")

    def test_range_contains_base_is_ok(self):
        t = vs.visible_text(page("Ставка 16.9% – 22%"))
        self.assertEqual(vs.judge(t, 16.9)[0], "ok")

    def test_changed_rate(self):
        t = vs.visible_text(page("Ставка от 18.49%"))
        self.assertEqual(vs.judge(t, 17.99), ("changed", 18.49))

    def test_script_text_is_ignored(self):
        # в <script> стоит «Ставка 99%», но она не должна считаться
        t = vs.visible_text(page("Ставка от 17.99%"))
        self.assertNotIn(99.0, vs.rate_candidates(t))

    def test_no_rate_on_page_is_unreadable(self):
        t = vs.visible_text(page("Подберите ипотеку"))
        self.assertEqual(vs.judge(t, 17.99), ("unreadable", None))

    def test_short_page_is_unreadable(self):
        self.assertEqual(vs.judge("Ставка 17.99%", 17.99)[0], "unreadable")


class VerifyOneTests(unittest.TestCase):
    def test_ok(self):
        e = vs.verify_one(URL, 17.99, TODAY, fetcher(page("Ставка от 17.99%")), {})
        self.assertEqual((e["status"], e["checked_at"]), ("ok", "2026-10-29"))

    def test_robots_disallow(self):
        f = fetcher(page("Ставка от 17.99%"), robots="User-agent: *\nDisallow: /bank/")
        self.assertEqual(vs.verify_one(URL, 17.99, TODAY, f, {})["status"], "blocked")

    def test_robots_missing_means_allowed(self):
        f = fetcher(page("Ставка от 17.99%"), robots=None)
        self.assertEqual(vs.verify_one(URL, 17.99, TODAY, f, {})["status"], "ok")

    def test_robots_unreachable_is_unreadable_not_blocked(self):
        def f(url, timeout=30):
            raise OSError("proxy 403")
        e = vs.verify_one(URL, 17.99, TODAY, f, {})
        self.assertEqual(e["status"], "unreadable")

    def test_user_agent_is_latin1_safe(self):
        vs.USER_AGENT.encode("latin-1")

    def test_fetch_error_is_unreadable(self):
        f = fetcher(OSError("timeout"))
        self.assertEqual(vs.verify_one(URL, 17.99, TODAY, f, {})["status"], "unreadable")

    def test_foreign_host_unsupported(self):
        e = vs.verify_one("https://www.sravni.ru/ipoteka/kommercheskaja", 18.49, TODAY, fetcher(""), {})
        self.assertEqual(e["status"], "unsupported")

    def test_no_url_unsupported(self):
        self.assertEqual(vs.verify_one("", 18.49, TODAY, fetcher(""), {})["status"], "unsupported")


def prog(checked, url=URL, rate=17.99):
    d = dt.date.fromisoformat(checked)
    age = (TODAY - d).days
    return {"bank": "Б", "name": "П", "source_url": url, "rate_min_pct": rate,
            "checked": checked, "age_days": age, "stale": age > ud.STALE_DAYS}


class ApplyTests(unittest.TestCase):
    def ver(self, status, day="2026-10-29", found=None):
        return {"sources": {URL: {"status": status, "checked_at": day, "found_rate_pct": found}}}

    def test_ok_refreshes_old_row(self):
        p = ud.apply_verification([prog("2026-09-01")], self.ver("ok"), TODAY)[0]
        self.assertEqual((p["checked"], p["stale"], p["age_days"]), ("2026-10-29", False, 0))
        self.assertEqual(p["checked_in_base"], "2026-09-01")

    def test_changed_marks_stale_even_if_fresh(self):
        p = ud.apply_verification([prog("2026-10-28")], self.ver("changed", found=18.5), TODAY)[0]
        self.assertTrue(p["stale"])
        self.assertEqual(p["source_rate_pct"], 18.5)

    def test_unreadable_changes_nothing(self):
        p = ud.apply_verification([prog("2026-09-01")], self.ver("unreadable"), TODAY)[0]
        self.assertEqual((p["checked"], p["stale"]), ("2026-09-01", True))

    def test_future_verification_ignored(self):
        p = ud.apply_verification([prog("2026-09-01")], self.ver("ok", day="2026-11-30"), TODAY)[0]
        self.assertEqual(p["checked"], "2026-09-01")

    def test_old_ok_does_not_move_date_back(self):
        p = ud.apply_verification([prog("2026-10-28")], self.ver("ok", day="2026-10-20"), TODAY)[0]
        self.assertEqual(p["checked"], "2026-10-28")

    def test_missing_verification_file(self):
        p = ud.apply_verification([prog("2026-09-01")], {}, TODAY)[0]
        self.assertTrue(p["stale"])


if __name__ == "__main__":
    unittest.main()
