import copy
import tempfile
import unittest
from contextlib import ExitStack
from io import StringIO
from pathlib import Path
from unittest.mock import Mock, patch

import main


class Clock:
    def __init__(self):
        self.now = 0.0

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        # Round away floating point residue at the configured deadline.
        self.now = round(self.now + seconds, 9)


def state(loading=False, ready=True, document='complete'):
    return {'loading': loading, 'search_ready': ready, 'ready_state': document}


class QueryWaitTests(unittest.TestCase):
    def wait(self, clock, probe, timeout=20):
        with (
            patch.object(main.time, 'monotonic', side_effect=clock.monotonic),
            patch.object(main.time, 'sleep', side_effect=clock.sleep),
            patch.object(main.time, 'time', side_effect=AssertionError('wall clock used')),
            patch.object(main, '_fr_read_query_state', side_effect=probe),
        ):
            return main._fr_wait_query_done(timeout)

    def test_finishes_after_busy_state_stabilizes_without_waiting_full_timeout(self):
        clock = Clock()
        result = self.wait(clock, lambda: state(clock.now < 1))
        self.assertTrue(result['ok'])
        self.assertTrue(result['completion_confirmed'])
        self.assertFalse(result['timeout'])
        self.assertGreaterEqual(clock.now, 1.8)
        self.assertLess(clock.now, 2.1)

    def test_delayed_loading_does_not_export_during_initial_idle_period(self):
        clock = Clock()
        result = self.wait(clock, lambda: state(3 <= clock.now < 6))
        self.assertTrue(result['saw_loading'])
        self.assertTrue(result['completion_confirmed'])
        self.assertGreaterEqual(clock.now, 6.8)
        self.assertLess(clock.now, 7.1)

    def test_normal_slow_query_can_use_the_configured_wait(self):
        clock = Clock()
        result = self.wait(clock, lambda: state(clock.now < 12), timeout=15)
        self.assertTrue(result['ok'])
        self.assertFalse(result['timeout'])
        self.assertGreaterEqual(clock.now, 12.8)

    def test_no_animation_waits_full_timeout_and_marks_unconfirmed_fallback(self):
        for ready in (True, None):
            with self.subTest(search_ready=ready):
                clock = Clock()
                result = self.wait(clock, lambda: state(ready=ready), timeout=8)
                self.assertEqual(clock.now, 8)
                self.assertTrue(result['ok'])
                self.assertTrue(result['timeout'])
                self.assertFalse(result['completion_confirmed'])
                self.assertEqual(result['completion_reason'], 'no_activity_timeout_fallback')
                self.assertFalse(main._query_wait_indicates_stall(result))

    def test_busy_button_is_activity_even_without_loading_animation(self):
        clock = Clock()
        result = self.wait(clock, lambda: state(ready=clock.now >= 2))
        self.assertTrue(result['saw_busy'])
        self.assertTrue(result['completion_confirmed'])
        self.assertFalse(result['saw_loading'])
        self.assertLess(clock.now, 3.1)

    def test_previously_observed_loading_can_finish_without_waiting_full_limit(self):
        clock = Clock()
        with (
            patch.object(main.time, 'monotonic', side_effect=clock.monotonic),
            patch.object(main.time, 'sleep', side_effect=clock.sleep),
            patch.object(main, '_fr_read_query_state', return_value=state()),
        ):
            result = main._fr_wait_query_done(20, observed_loading=True)
        self.assertTrue(result['completion_confirmed'])
        self.assertLess(clock.now, 1.1)

    def test_busy_at_limit_or_not_yet_stable_cannot_export(self):
        probes = (
            lambda t: state(loading=True),
            lambda t: state(ready=False),
            lambda t: state(document='loading'),
            lambda t: state(loading=t < 3.7),
        )
        for probe in probes:
            with self.subTest(probe=probe):
                clock = Clock()
                result = self.wait(clock, lambda: probe(clock.now), timeout=4)
                self.assertFalse(result['ok'])
                self.assertTrue(result['timeout'])
                self.assertTrue(main._query_wait_indicates_stall(result))
                self.assertEqual(clock.now, 4)

    def test_state_read_uses_one_browser_call(self):
        driver = Mock()
        driver.execute_script.return_value = state()
        with patch.object(main, '_driver', driver):
            self.assertEqual(main._fr_read_query_state(), state())
        driver.execute_script.assert_called_once()
        script = driver.execute_script.call_args.args[0]
        self.assertIn('document.readyState', script)
        self.assertIn('SEARCH', script)
        self.assertIn('loading-whole-indicator', script)

    def test_invalid_state_and_driver_error_fail_closed_without_leaking_url(self):
        for reply in (None, {}, {'loading': 'false'}, state(document='unknown')):
            with self.subTest(reply=reply):
                driver = Mock()
                driver.execute_script.return_value = reply
                with patch.object(main, '_driver', driver), self.assertRaises(RuntimeError):
                    main._fr_read_query_state()
        driver = Mock()
        driver.execute_script.side_effect = RuntimeError('https://private.invalid/?password=secret')
        with patch.object(main, '_driver', driver), self.assertRaises(RuntimeError) as error:
            main._fr_wait_query_done(8)
        self.assertNotIn('private.invalid', str(error.exception))
        self.assertNotIn('secret', str(error.exception))


class PerformanceHelpersTests(unittest.TestCase):
    def test_header_inspection_never_reads_entire_file(self):
        with tempfile.TemporaryDirectory() as temp:
            p = Path(temp) / 'synthetic.xlsx'
            p.write_bytes(b'PK' + b'0' * 1024)
            with patch.object(Path, 'read_bytes', side_effect=AssertionError('full file read')):
                result = main._inspect_download_file(p)
            self.assertTrue(result['valid_excel'])
            self.assertEqual(len(result['head_hex']), 32)

    def test_nested_recovery_time_is_excluded_from_query_time(self):
        clock = Clock()
        with patch.object(main.time, 'monotonic', side_effect=clock.monotonic):
            timer = main._StageTimer()
            timer.switch('query')
            clock.sleep(2)
            with timer.measure('recovery'):
                clock.sleep(5)
            clock.sleep(3)
            timer.switch()
        self.assertEqual(timer.seconds['query'], 5)
        self.assertEqual(timer.seconds['recovery'], 5)
        self.assertEqual(sum(timer.seconds.values()), 10)


class MatrixFlowTests(unittest.TestCase):
    def run_flow(self, scenario, values=('A',)):
        clock = Clock()
        events, records = [], []
        fields = {'institution_name': {'mode': 'list', 'values': list(values), 'method': 'sendkey'}}
        plan = {'fields': fields, 'query': True, 'export': True, 'query_wait_seconds': 8,
                'field_settle_seconds': 0, 'refresh_interval': 0, 'max_items': 0}
        counters = {'field': 0, 'query': 0, 'download': 0, 'wait': 0}

        def fill(*args, **kwargs):
            counters['field'] += 1
            events.append('fields')
            clock.sleep(1)
            if scenario == 'interrupt' and counters['field'] == 2:
                raise KeyboardInterrupt
            return {'ok': True, 'actual': 'synthetic', 'set_mode': 'sendkey'}

        def query():
            counters['query'] += 1
            events.append('query')
            return not (scenario == 'missing_button' and counters['query'] == 1)

        def wait_query(*args, **kwargs):
            counters['wait'] += 1
            incomplete = scenario == 'timeout' or (scenario == 'previous_query_busy' and counters['wait'] == 1)
            clock.sleep(8 if incomplete else 2)
            if incomplete:
                self.assertEqual(args[0], 8)
                return {'ok': False, 'timeout': True, 'loading': True, 'elapsed_seconds': 8}
            return {'ok': True, 'timeout': False, 'loading': False, 'elapsed_seconds': 2,
                    'completion_confirmed': True}

        def recover(*args, **kwargs):
            events.append('recover')
            clock.sleep(3)
            return True

        def wait_download(*args, **kwargs):
            counters['download'] += 1
            clock.sleep(1)
            if scenario == 'export_failure' and counters['download'] == 1:
                return {'ok': False, 'files': [], 'timeout': True}
            return {'ok': True, 'files': [{'path': 'synthetic.xlsx', 'size': 100}]}

        patches = {
            '_driver': Mock(), 'driver_alive': lambda: True,
            'load_task_config': lambda scene: {'synthetic': True},
            'parse_matrix_config': lambda config, scene: plan,
            '_switch_to_report_frame': lambda *a, **kw: True, '_save_iframe_src': lambda: None,
            '_check_page_alive': lambda: True,
            '_fr_is_loading': lambda: scenario == 'previous_query_busy' and counters['field'] == 0,
            '_fr_get_field_value': lambda *a, **kw: '', '_fr_set_field': fill,
            '_fr_click_query': query, '_fr_wait_query_done': wait_query,
            '_fr_accept_possible_dialog': lambda: {'accepted': False},
            '_recover_report': recover,
            '_fr_click_export': lambda: events.append('export'),
            '_wait_download_delta': wait_download,
            '_rename_download_files': lambda result, *args: result,
            '_append_jsonl': lambda path, record: records.append(copy.deepcopy(record)),
        }
        with tempfile.TemporaryDirectory() as temp, ExitStack() as stack:
            stack.enter_context(patch.object(main, 'OUTPUT', Path(temp)))
            stack.enter_context(patch.object(main, '_download_snapshot', return_value=(Path(temp), {})))
            stack.enter_context(patch.object(main.time, 'monotonic', side_effect=clock.monotonic))
            stack.enter_context(patch.object(main.time, 'sleep', side_effect=clock.sleep))
            stack.enter_context(patch('selenium.webdriver.common.action_chains.ActionChains'))
            stack.enter_context(patch('sys.stdout', new_callable=StringIO))
            for name, value in patches.items():
                stack.enter_context(patch.object(main, name, value))
            main.run_matrix_export('org')
        return events, records, clock.now

    def test_export_recovery_refills_and_requeries_before_export(self):
        events, records, elapsed = self.run_flow('export_failure')
        self.assertEqual(events, ['fields', 'query', 'export', 'recover', 'fields', 'query', 'export'])
        summary = next(r for r in records if r['type'] == 'run_summary')
        self.assertEqual((summary['ok'], summary['fail'], summary['attempt_count']), (1, 0, 2))
        self.assertEqual(summary['stage_seconds']['recovery'], 3)
        self.assertEqual(summary['stage_seconds']['query'], 4)
        self.assertAlmostEqual(sum(summary['stage_seconds'].values()), elapsed, places=2)

    def test_missing_query_button_recovery_does_not_query_reset_page_directly(self):
        events, records, _ = self.run_flow('missing_button')
        self.assertEqual(events, ['fields', 'query', 'recover', 'fields', 'query', 'export'])
        self.assertEqual(next(r for r in records if r['type'] == 'run_summary')['ok'], 1)

    def test_previous_query_waits_before_recovery_and_refilling(self):
        events, records, _ = self.run_flow('previous_query_busy')
        self.assertEqual(events, ['recover', 'fields', 'query', 'export'])
        summary = next(r for r in records if r['type'] == 'run_summary')
        self.assertEqual(summary['stage_seconds']['query'], 10)
        self.assertEqual(summary['stage_seconds']['recovery'], 3)

    def test_query_at_limit_has_bounded_retries_and_never_exports(self):
        events, records, _ = self.run_flow('timeout')
        self.assertEqual(events.count('recover'), 3)
        self.assertEqual(events.count('query'), 4)
        self.assertNotIn('export', events)
        summary = next(r for r in records if r['type'] == 'run_summary')
        self.assertEqual((summary['ok'], summary['fail'], summary['processed'], summary['attempt_count']), (0, 1, 1, 4))
        self.assertEqual(len([r for r in records if r['type'] == 'item_summary']), 4)

    def test_interrupt_average_uses_processed_items_and_preserves_partial_timing(self):
        _, records, elapsed = self.run_flow('interrupt', values=('A', 'B', 'C'))
        summary = next(r for r in records if r['type'] == 'run_summary')
        self.assertTrue(summary['interrupted'])
        self.assertEqual((summary['total'], summary['processed'], summary['attempt_count']), (3, 1, 2))
        self.assertEqual(summary['avg_seconds_per_item'], round(elapsed, 3))
        self.assertEqual(len([r for r in records if r['type'] == 'item_summary']), 2)
        self.assertAlmostEqual(sum(summary['stage_seconds'].values()), elapsed, places=2)


if __name__ == '__main__':
    unittest.main()
