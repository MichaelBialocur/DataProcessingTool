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
