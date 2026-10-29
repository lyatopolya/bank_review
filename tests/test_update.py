import datetime as dt
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import update_data as u

XLSX = Path(__file__).resolve().parent / "fixture-baza.xlsx"  # неизменная копия, чтобы правки базы не ломали тесты
TODAY = dt.date(2026, 10, 27)
D = dt.date

NESTED = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body>
<KeyRateXMLResponse xmlns="http://web.cbr.ru/"><KeyRateXMLResult>
<KeyRate xmlns=""><KR><DT>2026-06-22T00:00:00+03:00</DT><Rate>14.25</Rate></KR>
<KR><DT>2026-07-27T00:00:00+03:00</DT><Rate>14</Rate></KR></KeyRate>
</KeyRateXMLResult></KeyRateXMLResponse></soap:Body></soap:Envelope>"""

ESCAPED = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body>
<KeyRateXMLResponse xmlns="http://web.cbr.ru/"><KeyRateXMLResult>
&lt;KeyRate&gt;&lt;KR&gt;&lt;DT&gt;2026-09-11T00:00:00&lt;/DT&gt;&lt;Rate&gt;14,0&lt;/Rate&gt;&lt;/KR&gt;&lt;/KeyRate&gt;
</KeyRateXMLResult></KeyRateXMLResponse></soap:Body></soap:Envelope>"""

EMPTY = """<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
<soap:Body><KeyRateXMLResponse xmlns="http://web.cbr.ru/"><KeyRateXMLResult/></KeyRateXMLResponse></soap:Body></soap:Envelope>"""

# Образцы страниц ЦБ составлены по их описанию, это не снимки реальных страниц.
TABLE_HTML = """<html><body><h1>Ключевая ставка Банка России</h1><p>с 01.01.2026 по 27.10.2026</p>
<table class="data"><tr><th>Дата</th><th>Ставка</th></tr>
<tr><td>27.07.2026</td><td>14,00</td></tr>
<tr><td>22.06.2026</td><td>14,25</td></tr>
<tr><td>27.04.2026</td><td>14,50</td></tr></table></body></html>"""

PRESS_SAVE = """<div class="news"><div class="date">23 октября 2026</div>
<a href="/press/keypr/1">Банк России принял решение сохранить ключевую ставку на уровне 14,00% годовых</a></div>
<div class="news"><div class="date">11 сентября 2026</div>
<a href="/press/keypr/2">Банк России принял решение сохранить ключевую ставку на уровне 14,00% годовых</a></div>"""

PRESS_CUT = """<div class="date">24 июля 2026</div>
<a>Банк России принял решение снизить ключевую ставку на 25 б.п., до 14,00% годовых</a>"""

PRESS_HIKE = """<span>20.03.2026</span><a>Банк России принял решение повысить ключевую ставку на 100 б. п. до 16,00% годовых</a>"""


def src(name, value, since=None, meeting=None):
    return {"source": name, "value_pct": value, "effective_since": since, "meeting": meeting}


class ParseTests(unittest.TestCase):
    def test_soap_nested_takes_latest(self):
        self.assertEqual(u.parse_key_rate(NESTED), (D(2026, 7, 27), 14.0))

    def test_soap_escaped_and_comma(self):
        self.assertEqual(u.parse_key_rate(ESCAPED), (D(2026, 9, 11), 14.0))

    def test_soap_empty_raises(self):
        with self.assertRaises(ValueError):
            u.parse_key_rate(EMPTY)

    def test_table_takes_latest_and_ignores_header_dates(self):
        self.assertEqual(u.parse_key_rate_table(TABLE_HTML, TODAY), (D(2026, 7, 27), 14.0))

    def test_table_ignores_future_dates(self):
        html = TABLE_HTML + "<tr><td>18.12.2026</td><td>9,00</td></tr>"
        self.assertEqual(u.parse_key_rate_table(html, TODAY), (D(2026, 7, 27), 14.0))

    def test_table_without_data_raises(self):
        with self.assertRaises(ValueError):
            u.parse_key_rate_table("<html>техработы</html>", TODAY)

    def test_press_saved_takes_newest_item(self):
        rate, meeting = u.parse_press(PRESS_SAVE)
        self.assertEqual(rate, 14.0)
        self.assertEqual(meeting, {"date": D(2026, 10, 23), "decision": "сохранена"})

    def test_press_cut_and_hike_wording(self):
        rate, meeting = u.parse_press(PRESS_CUT)
        self.assertEqual((rate, meeting["decision"], meeting["date"]), (14.0, "снижена", D(2026, 7, 24)))
        rate, meeting = u.parse_press(PRESS_HIKE)
        self.assertEqual((rate, meeting["decision"], meeting["date"]), (16.0, "повышена", D(2026, 3, 20)))

    def test_press_without_decision_raises(self):
        with self.assertRaises(ValueError):
            u.parse_press("<html>ничего нет</html>")


class ReconcileTests(unittest.TestCase):
    def test_rate_unchanged_for_months_is_fresh(self):
        # Регрессия: ставка не менялась с 27.07, но подтверждена сегодня. Это «ok», а не «устарело».
        prev = {}
        res = [src("soap", 14.0, D(2026, 7, 27)),
               src("press", 14.0, None, {"date": D(2026, 10, 23), "decision": "сохранена"})]
        kr, warns = u.reconcile_key_rate(res, [], prev, TODAY)
        self.assertEqual(kr["status"], "ok")
        self.assertEqual(kr["effective_since"], "2026-07-27")
        self.assertEqual(kr["confirmed_at"], "2026-10-27")
        self.assertEqual(kr["last_meeting"], {"date": "2026-10-23", "decision": "сохранена"})
        self.assertEqual(kr["sources"], ["soap", "press"])
        self.assertEqual(warns, [])

    def test_fallback_to_table_when_soap_fails(self):
        kr, warns = u.reconcile_key_rate([src("table", 14.0, D(2026, 7, 27))], ["источник «soap»: сеть"], {}, TODAY)
        self.assertEqual(kr["status"], "ok")
        self.assertEqual(kr["sources"], ["table"])
        self.assertEqual(len(warns), 1)

    def test_press_only_keeps_previous_change_date_if_value_same(self):
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-26"}
        kr, _ = u.reconcile_key_rate([src("press", 14.0, None, {"date": D(2026, 10, 23), "decision": "сохранена"})], [], prev, TODAY)
        self.assertEqual(kr["effective_since"], "2026-07-27")

    def test_press_only_value_changed_does_not_invent_change_date(self):
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-26"}
        kr, _ = u.reconcile_key_rate([src("press", 13.5, None, {"date": D(2026, 10, 23), "decision": "снижена"})], [], prev, TODAY)
        self.assertIsNone(kr["effective_since"])
        self.assertEqual(kr["value_pct"], 13.5)

    def test_disagreement_is_flagged(self):
        kr, warns = u.reconcile_key_rate([src("soap", 14.0, D(2026, 7, 27)), src("press", 15.0)], [], {}, TODAY)
        self.assertEqual(kr["status"], "conflict")
        self.assertEqual(kr["value_pct"], 14.0)
        self.assertTrue(any("разные значения" in w for w in warns))

    def test_all_fail_recent_previous_still_ok_with_warning(self):
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-26"}
        kr, warns = u.reconcile_key_rate([], ["источник «soap»: сеть"], prev, TODAY)
        self.assertEqual(kr["status"], "ok")
        self.assertEqual(kr["confirmed_at"], "2026-10-26")
        self.assertIsNone(kr["last_meeting"])
        self.assertTrue(any("прошлого запуска" in w for w in warns))

    def test_all_fail_old_previous_is_stale(self):
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-20"}
        kr, _ = u.reconcile_key_rate([], ["сеть"], prev, TODAY)
        self.assertEqual(kr["status"], "stale")

    def test_old_file_format_without_confirmation_is_stale(self):
        kr, _ = u.reconcile_key_rate([], [], {"value_pct": 14.0, "date": "2026-07-27"}, TODAY)
        self.assertEqual(kr["status"], "stale")

    def test_nothing_at_all_is_missing(self):
        kr, _ = u.reconcile_key_rate([], ["сеть"], {}, TODAY)
        self.assertEqual(kr["status"], "missing")
        self.assertIsNone(kr["value_pct"])

    def test_old_meeting_is_not_shown(self):
        old = {"date": D(2026, 3, 20), "decision": "повышена"}
        with mock.patch.object(u, "_http", return_value=PRESS_HIKE):
            r = u.src_press(TODAY)
        self.assertEqual(r["value_pct"], 16.0)
        self.assertIsNone(r["meeting"])
        self.assertEqual(old["decision"], "повышена")


class CollectTests(unittest.TestCase):
    def test_failed_source_does_not_stop_the_others(self):
        def bad(_):
            raise RuntimeError("нет сети")
        with mock.patch.object(u, "SOURCES", [("soap", bad), ("table", lambda t: src("table", 14.0, D(2026, 7, 27)))]):
            results, errors = u.collect_key_rate(TODAY)
        self.assertEqual([r["source"] for r in results], ["table"])
        self.assertEqual(len(errors), 1)
        self.assertIn("soap", errors[0])

    def test_request_window_covers_long_gaps(self):
        body = u._soap_body(TODAY - dt.timedelta(days=u.KEY_RATE_WINDOW_DAYS), TODAY).decode()
        self.assertIn("<fromDate>2025-09-22</fromDate>", body)
        self.assertIn("<ToDate>2026-10-27</ToDate>", body)


class BuildTests(unittest.TestCase):
    def _build(self, sources, prev=None, offline=False):
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / "m.json"
            if prev is not None:
                out.write_text(json.dumps({"key_rate": prev}), encoding="utf-8")
            with mock.patch.object(u, "SOURCES", sources):
                return u.build(XLSX, out, TODAY, offline), out

    def test_programs_and_staleness(self):
        progs, warns = u.load_programs(XLSX, TODAY)
        by = {p["bank"]: p for p in progs}
        self.assertEqual(len(progs), 3)
        self.assertAlmostEqual(by["Сбербанк"]["rate_min_pct"], 21.2)
        self.assertEqual(by["Сбербанк"]["age_days"], 5)
        self.assertFalse(by["Сбербанк"]["stale"])
        self.assertTrue(by["Совкомбанк"]["stale"])
        self.assertIsNone(by["Альфа-Банк"]["down_min_pct"])
        self.assertEqual(warns, [])

    def test_build_full_scenario(self):
        sources = [("soap", lambda t: src("soap", 14.0, D(2026, 7, 27))),
                   ("press", lambda t: src("press", 14.0, None, {"date": D(2026, 10, 23), "decision": "сохранена"}))]
        res, _ = self._build(sources)
        self.assertEqual(res["key_rate"]["status"], "ok")
        self.assertEqual(res["summary"]["lowest_rate_spread_over_key_pct"], 5.59)
        self.assertEqual(res["summary"]["stale_programs"], 1)

    def test_offline_keeps_fresh_previous_without_marking_it_stale(self):
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-27",
                "last_meeting": {"date": "2026-10-23", "decision": "сохранена"}, "sources": ["soap"], "status": "ok"}
        res, _ = self._build([], prev, offline=True)
        self.assertEqual(res["key_rate"]["status"], "ok")
        self.assertEqual(res["key_rate"]["value_pct"], 14.0)

    def test_offline_without_any_data(self):
        res, _ = self._build([], None, offline=True)
        self.assertEqual(res["key_rate"]["status"], "missing")
        self.assertIsNone(res["summary"]["lowest_rate_spread_over_key_pct"])


class CheckTests(unittest.TestCase):
    def _check(self, kr):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "m.json"
            p.write_text(json.dumps({"key_rate": kr}), encoding="utf-8")
            return u.check(p, TODAY)

    def test_fresh_passes(self):
        self.assertEqual(self._check({"value_pct": 14.0, "confirmed_at": "2026-10-27", "status": "ok"}), [])

    def test_old_confirmation_fails_even_if_file_says_ok(self):
        # Регрессия: если ежедневный запуск встал, файл не должен вечно говорить «ok».
        problems = self._check({"value_pct": 14.0, "confirmed_at": "2026-10-10", "status": "ok"})
        self.assertTrue(problems)

    def test_missing_conflict_and_unknown_fail(self):
        self.assertTrue(self._check({}))
        self.assertTrue(self._check({"value_pct": 14.0, "confirmed_at": "2026-10-27", "status": "conflict"}))
        self.assertTrue(self._check({"value_pct": 14.0}))



class DailySeriesTests(unittest.TestCase):
    """ЦБ может отдавать ставку на каждый день: дата изменения это начало последнего отрезка."""

    def test_daily_rows_give_start_of_run(self):
        days = [(D(2026, 7, 20) + dt.timedelta(days=i)) for i in range(0, 100)]
        pairs = [(d, 15.0 if d < D(2026, 7, 27) else 14.0) for d in days]
        self.assertEqual(u._latest_run(pairs), (D(2026, 7, 27), 14.0))

    def test_unchanged_over_whole_window_is_unknown(self):
        pairs = [(D(2025, 9, 22) + dt.timedelta(days=i), 14.0) for i in range(400)]
        self.assertEqual(u._latest_run(pairs), (None, 14.0))

    def test_single_row_keeps_its_date(self):
        self.assertEqual(u._latest_run([(D(2026, 7, 27), 14.0)]), (D(2026, 7, 27), 14.0))

    def test_unknown_start_falls_back_to_previous_date(self):
        res = [{"source": "soap", "value_pct": 14.0, "effective_since": None, "meeting": None}]
        prev = {"value_pct": 14.0, "effective_since": "2026-07-27", "confirmed_at": "2026-10-28"}
        kr, _ = u.reconcile_key_rate(res, [], prev, TODAY)
        self.assertEqual(kr["effective_since"], "2026-07-27")


if __name__ == "__main__":
    unittest.main(verbosity=2)
