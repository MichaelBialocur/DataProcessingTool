# DataProcessingTool

Desktop Python tool for PHP and LTS thermal-test processing, Shift2DC CPU tests, and simulation-to-experiment correlation.

## Run

Use Python 3.12 with Tkinter available (included with the standard Windows Python installer).

```bash
python -m pip install -r requirements.txt
python process_php_csvs.py
```

For refrigerant saturation properties and pressure–enthalpy diagrams, install CoolProp or configure an installed NIST REFPROP library with its Python bridge:

```bash
python -m pip install CoolProp
# Optional bridge for an existing NIST REFPROP installation:
python -m pip install ctREFPROP
```

NIST REFPROP itself is a separate installation. Property availability depends on the selected fluid and backend.

## Workflows

- **Test analysis:** select raw CSV data and configure report sections, summary values, individual test details, and colormap limits. Outputs include cleaned data, vertically stacked Excel averages, and a PDF report.
- **Shift2DC:** supports T_CPU channels, saved per-CPU heat-load plateaus, uniform or heterogeneous CPU loads entered directly or imported from CSV, and selected pressure/PSU/p–h pages. Initial schedule entry writes W_CPU columns into the selected raw CSV and removes obsolete W_PSU/I_PSU/V_PSU channels; later runs reuse the saved schedule.
- **Correlation:** select a reference XLSX workbook and simulation CSV/XLSX, map conditions, and identify the physics/solver configuration. Matching configurations update existing result columns; new configurations add columns. Failed and unmatched runs preserve existing results. Charts retain solver colors, project marker shapes, and parity/deviation lines; project labels remain merged by block.

The source combines the report-processing and correlation-analysis changes developed across the related project conversations. Experimental workbooks and raw measurements are not bundled.

## Superheating and subcooling (LTS / Shift2DC)

Select **Superheating [K]** in the report summary values to include
`T_EVAP_OUT - T_SAT`. Saturation temperature uses the measured Psat,
its selected pressure units/reference, and the filename refrigerant with the
existing CoolProp/REFPROP calculation. Both values use the same plateau averaging
windows. Missing pressure, temperature, or fluid properties leave the result blank.
Separate pressure sensors retain separate superheating columns.

The GUI offers a combined **superheating and subcooling comparison** page,
including single-test campaigns. Superheating above **1 K (1°C difference)** is
red in the summary by default; its threshold is editable alongside subcooling
and CPU temperature thresholds. Exactly 1 K is not highlighted. Subcooling keeps
the established definition `T_COND_IN - T_COND_OUT`.

Run the focused calculation checks with:

```bash
python -m unittest discover -s tests -v
```

## Saved CPU powers in CSV files

Raw and postprocessed Shift2DC CSVs retain the measured channels and only add
`W_CPU_<ID>` for the selected CPUs. Off periods and individually unpowered CPUs
contain `0`. Existing all-blank CPU rows from older schedule exports become zero.
The five legacy schedule/total columns (`Heat_Load_Step`,
`Heat_Load_Step_Start_s`, `Heat_Load_Step_End_s`, `Total_Heat_Load_W`, and
`Scheduled_Power_Per_Board_W`) are removed when the file is processed.

Reopening a test detects plateaus directly from the full set of CPU powers,
including a change in load distribution with unchanged total power. Total power
is calculated internally for reports, including transient plots. Boundaries use
the recorded sample timestamps; consecutive intervals with identical per-CPU
loads form one plateau. The final 100 seconds of each accepted plateau are used
for averaging. Partial missing or invalid CPU loads still require correction.

For steady-state Shift2DC tests, processing also trims the rewritten raw CSV
and postprocessed CSV to **100 seconds after the final heat-load step**. The
sample exactly at that cutoff is retained; later rows are deleted. This uses
elapsed seconds, not a fixed number of samples, and works for both new schedules
and existing `W_CPU` plateaus. Earlier off gaps and all powered intervals remain;
recordings with less than 100 seconds of cooldown are not padded. Reprocessing
keeps the same cutoff and plateau averages. Transient (`_TR`) recordings retain
their full duration.

## Excel CPU columns and filling-ratio colors

Each individual `T_CPU_<ID> [°C]` column is immediately followed by its matching
`W_CPU_<ID> [W]`. Physical CPU IDs determine the pairing, including gaps in the
numbering and heterogeneous loads. Missing channels stay blank; recorded zero
loads stay zero. Results from different CSVs continue to stack downwards.

Excel filling-ratio cells and report FR curves use the same color gradient.
The established anchors remain 40% green, 50% yellow/orange, 60% red, and 70%
purple. Ratios between anchors blend in RGB (54% is 40% of the way from the
50% color to the 60% color); ratios outside the anchor range use the nearest
endpoint color.
