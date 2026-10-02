import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import numpy as np
import pandas as pd
import process_php_csvs as tool


def raw_data():
    return pd.DataFrame({'RelTime': np.arange(600, dtype=float),
        'T_CPU_1': 55., 'T_CPU_3': 56., 'T_CPU_5': 50.,
        'T_WATER_IN': 25., 'T_WATER_OUT': 26., 'VFR_WATER': 2.7,
        'T_EVAP_IN': 25., 'T_EVAP_OUT': 31., 'T_COND_IN': 30., 'T_COND_OUT': 26.,
        'W_PSU': 999., 'I_PSU_1': 9., 'V_PSU_2_SP': 48., 'T_PSU_1': 45.})


def config():
    return dict(board_count=3, board_columns=['T_CPU_1', 'T_CPU_3', 'T_CPU_5'],
                start_s=10., durations_s=[200., 200.], cpu_powers=[[10., 20., 0.], [20., 10., 0.]])


def saved_data():
    data = raw_data()
    boards, _, schedule = tool.build_shift2dc_steps(data, config())
    return tool.save_shift2dc_schedule_columns(data, boards, schedule)


class CPUCsvTests(unittest.TestCase):
    def test_zero_off_intervals_and_vector_plateaus(self):
        saved = saved_data()
        self.assertFalse(set(tool.SHIFT2DC_LEGACY_SCHEDULE_COLUMNS) & set(saved))
        self.assertEqual(set(tool.shift2dc_cpu_power_columns(saved)), {1, 3, 5})
        self.assertTrue((saved.loc[(saved.RelTime < 10) | (saved.RelTime >= 410), ['W_CPU_1', 'W_CPU_3', 'W_CPU_5']] == 0).all().all())
        self.assertTrue((saved['W_CPU_5'] == 0).all())
        for column in ['W_PSU', 'I_PSU_1', 'V_PSU_2_SP']:
            self.assertNotIn(column, saved)
        self.assertTrue(saved['T_PSU_1'].equals(raw_data()['T_PSU_1']))
        _, _, recovered = tool.build_shift2dc_saved_steps(saved)
        self.assertEqual(len(recovered), 2)
        self.assertEqual([r['Scheduled total power [W]'] for r in recovered], [30., 30.])
        self.assertEqual([r['W_CPU_1 [W]'] for r in recovered], [10., 20.])
        self.assertEqual([r['Average from [s]'] for r in recovered], [110., 310.])

    def test_legacy_migration_and_validation(self):
        saved = saved_data()
        saved.loc[saved.RelTime < 10, ['W_CPU_1', 'W_CPU_3', 'W_CPU_5']] = np.nan
        saved['Heat_Load_Step_Start_s'] = 'obsolete'
        saved['Total_Heat_Load_W'] = 99999
        saved['Scheduled_Power_Per_Board_W'] = 999
        cleaned = tool.clean_shift2dc_power_channels(saved)
        self.assertTrue((cleaned.loc[cleaned.RelTime < 10, 'W_CPU_1'] == 0).all())
        self.assertFalse(set(tool.SHIFT2DC_LEGACY_SCHEDULE_COLUMNS) & set(cleaned))
        self.assertEqual(len(tool.build_shift2dc_saved_steps(cleaned)[2]), 2)
        cleaned.loc[20, 'W_CPU_3'] = np.nan
        with self.assertRaisesRegex(ValueError, 'incomplete'):
            tool.build_shift2dc_saved_steps(tool.clean_shift2dc_power_channels(cleaned))
        cleaned.loc[20, 'W_CPU_3'] = -1
        with self.assertRaisesRegex(ValueError, 'negative'):
            tool.build_shift2dc_saved_steps(cleaned)

    def test_off_gap_and_transient_total(self):
        data = raw_data()
        data['W_CPU_1'] = 0.
        data['W_CPU_3'] = 0.
        data.loc[10:209, ['W_CPU_1', 'W_CPU_3']] = [10., 20.]
        data.loc[220:419, ['W_CPU_1', 'W_CPU_3']] = [10., 20.]
        cleaned = tool.clean_shift2dc_power_channels(data)
        self.assertEqual(len(tool.build_shift2dc_saved_steps(cleaned)[2]), 2)
        metadata = tool.parse_file_name(Path('Shift2DC_LTS_EVAP-demo_COND-demo_R1336mzzE_FR60_Water_TW25_VFR2p7_TR.csv'))
        detail = dict(data=cleaned, metadata=metadata, boards=['T_CPU_1', 'T_CPU_3'], t_cu_columns=['T_CPU_1', 'T_CPU_3'])
        panels = tool.transient_panels(detail, {})
        total = next(series['Total heat load'] for _, _, series in panels if 'Total heat load' in series)
        self.assertEqual(total.iloc[0], 0.)
        self.assertEqual(total.iloc[100], 30.)

    def test_raw_and_postprocessed_roundtrip_without_dialog(self):
        name = 'Shift2DC_LTS_EVAP-demo_COND-demo_R1336mzzE_FR60_Water_TW25_VFR2p7_SS.csv'
        with tempfile.TemporaryDirectory() as tmp, \
             patch.object(tool, 'configure_report', return_value=({}, {})), \
             patch.object(tool, 'write_shift2dc_excel'), \
             patch.object(tool, 'write_shift2dc_pdf'), \
             patch.object(tool, 'show_shift2dc_dialog', side_effect=AssertionError('Unexpected schedule prompt')):
            root = Path(tmp)
            source = root/name
            raw_data().to_csv(source, index=False)
            first = tool.process_shift2dc_files([source], root/'out', configurations={name: config()}, open_pdf_when_done=False)
            raw = pd.read_csv(source)
            pd.testing.assert_frame_equal(raw, pd.read_csv(root/'out'/name))
            self.assertFalse(set(tool.SHIFT2DC_LEGACY_SCHEDULE_COLUMNS) & set(raw))
            second = tool.process_shift2dc_files([source], root/'out', configurations={}, open_pdf_when_done=False)
            pd.testing.assert_frame_equal(pd.DataFrame(first['results']), pd.DataFrame(second['results']))
            pd.testing.assert_frame_equal(raw, pd.read_csv(source))


if __name__ == '__main__':
    unittest.main()
