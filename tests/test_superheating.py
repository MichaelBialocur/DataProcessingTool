import unittest
from pathlib import Path
from unittest.mock import patch
import numpy as np
import pandas as pd
import process_php_csvs as tool


def fixture():
    path = Path('Test_LTS_EVAP-demo_COND-demo_R1336mzzE_FR60_Water_TW25_VFR2p7_SS.csv')
    metadata = tool.parse_file_name(path)
    metadata["source_file"] = path.name
    steps = []
    for power, pressure, evap in [(100, 2., 20.5), (200, 3., 31.), (300, 4., 42.)]:
        steps.append(pd.DataFrame({'RelTime': np.arange(100), 'W_PSU': power,
            'T_CU_1': 55., 'T_WATER_IN': 25., 'T_WATER_OUT': 26., 'VFR_WATER': 2.7,
            'T_EVAP_IN': 22., 'T_EVAP_OUT': evap, 'T_COND_IN': 30., 'T_COND_OUT': 26.,
            'Psat': pressure}))
    rows = [tool.calculate_step_result(step, metadata, ['T_CU_1']) for step in steps]
    detail = {'source_file': path.name, 'metadata': metadata, 'data': pd.concat(steps), 'steps': steps}
    return rows, detail


def saturation_stub(pressure, fluid):
    # Deterministic property stub: tests pressure conversion and averaging independently of a licensed backend.
    return np.asarray(pressure) / 10000., 'test property backend'


class SuperheatingTests(unittest.TestCase):
    @patch.object(tool, 'saturation_temperatures', side_effect=saturation_stub)
    def test_lts_calculation_and_optional_summary(self, _):
        rows, detail = fixture()
        output = tool.add_lts_psat_results(rows, [detail], {})
        self.assertEqual([r['Superheating [K]'] for r in output], [.5, 1., 2.])
        self.assertEqual([r['Subcooling [K]'] for r in output], [4., 4., 4.])
        available, defaults = tool.summary_value_options(output, 'LTS', [detail])
        self.assertIn('Superheating [K]', available)
        self.assertNotIn('Superheating [K]', defaults)
        before, _ = tool.summary_value_options(rows, 'LTS', [detail])
        self.assertIn('Superheating [K]', before)

    @patch.object(tool, 'saturation_temperatures', side_effect=saturation_stub)
    def test_shift_schedule_and_multiple_pressures(self, _):
        rows, detail = fixture()
        data = detail['steps'][0].copy()
        data['RelTime'] = np.arange(100)
        data['Psat'] = np.linspace(2, 4, 100)
        data['Psat_2'] = data['Psat'] + 1
        detail.update(data=data, steps=[data], schedule=[{'Average from [s]': 10.5, 'Average to [s]': 99.5}])
        expected = tool.shift2dc_auxiliary_averages(detail, data['Psat'].to_numpy()*10)[0]
        output = tool.add_lts_psat_results(rows[:1], [detail], {})[0]
        self.assertAlmostEqual(output['Superheating [K]'], rows[0]['T_EVAP_OUT [°C]']-expected)
        self.assertAlmostEqual(output['Superheating_2 [K]'], output['Superheating [K]']-10)
        settings = {'pressure_settings': {'Psat': {'unit': 'bar', 'basis': 'Gauge', 'atmospheric_pressure_pa': 100000}}}
        adjusted = tool.add_lts_psat_results(rows[:1], [detail], settings)[0]
        self.assertAlmostEqual(adjusted['Superheating [K]'], output['Superheating [K]']-10)

    def test_threshold_is_strict_and_configurable(self):
        for value, expected in [(1., False), (1.01, True), (-2., False), (np.nan, False)]:
            self.assertEqual(tool.summary_cell_exceeds_limit('Superheating [K]', value, pd.DataFrame(), {}), expected)
        self.assertFalse(tool.summary_cell_exceeds_limit('Superheating_2 [K]', 1.1, pd.DataFrame(), {'summary_superheating_limit': 2}))
        with self.assertRaises(ValueError):
            tool.summary_red_thresholds({'summary_superheating_limit': 'nan'})

    def test_missing_properties_remain_blank(self):
        rows, detail = fixture()
        with patch.object(tool, 'saturation_temperatures', return_value=(np.full(100, np.nan), 'unavailable')):
            output = tool.add_lts_psat_results(rows, [detail], {})
        self.assertTrue(all(np.isnan(r['Superheating [K]']) for r in output))
        self.assertIn('Superheating [K]', tool.summary_value_options(output, 'LTS')[0])
        self.assertEqual(len(list(tool.superheating_subcooling_groups(pd.DataFrame(rows)))), 1)


if __name__ == '__main__':
    unittest.main()
