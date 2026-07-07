# Digital Crewchief

Excel-based maintenance forecasting and crew planning tool for Apache aircraft. The workbook imports SITS maintenance data, applies rotation cycles and flying-hour rates, and produces per-aircraft forecast charts plus a fleet dashboard.

**Workbook:** `20260603-Digital Crewchief_BETA_MASTER.xlsm`

## Repository layout

```
Digital_Crewchief/
├── 20260603-Digital Crewchief_BETA_MASTER.xlsm   # Main workbook (binary)
├── vba/
│   ├── Modules/          # Standard modules (.bas)
│   ├── ExcelObjects/     # ThisWorkbook and worksheet code (.cls)
│   └── Forms/            # UserForm code (.frm)
└── tools/
    └── export_vba.py     # Re-export VBA source from the workbook
```

## Workbook sheets

| Sheet | Purpose |
|---|---|
| **TODO** | Development backlog |
| **Task_Rules** | Task alignment and grouping rules |
| **Settings** | Forecast dates, thresholds, flying rates (named ranges) |
| **Rotation** | Aircraft roster, cycle position, down-maintenance windows |
| **Dashboard** | Fleet-wide maintenance summary |
| **{TAIL}_IMPORT** | Imported SITS task data per aircraft (e.g. ZZ382_IMPORT) |
| **{TAIL}_CHART** | Gantt-style maintenance forecast chart per aircraft |

## VBA architecture

| Area | Key modules |
|---|---|
| Entry / refresh | `modRefresh`, `modAircraftRefresh`, `ThisWorkbook` |
| Data import | `modImport`, `modImportTaskReader`, `modImportLayout` |
| Aircraft roster | `modAircraftRoster`, `modAircraftAdmin`, `modAircraftStatus` |
| Planning logic | `modRotationSchedule`, `modPlannerRates`, `modPlannerSettings`, `modPlannerCalendar` |
| Task engine | `modTaskOccurrenceBuilder`, `modTaskGrouping`, `modTaskMovement`, `modTaskRules` |
| Charts & dashboard | `modChartOutput`, `modChartLayout`, `modDashboardLayout`, `modDashboardBuckets` |
| User interfaces | `frmAircraftManager`, `frmPlannerSettings`, `frmTaskRules` |
| Export | `modWordExport` |

**Main refresh flow:** `Workbook_Open` → `DeferredRefreshAll` → `RefreshAll`, which rebuilds forecasts for every aircraft on the roster.

## Working with VBA in this repo

Edit files under `vba/`, then import changed modules into the workbook via the VBA editor (Alt+F11 → File → Import File). Re-export after workbook changes:

```bash
pip install pyopenvba
python tools/export_vba.py
```

Run **RefreshAll** (or reopen the workbook) to rebuild charts and the dashboard after logic changes.

## UserForms

| Form | Purpose |
|---|---|
| **frmAircraftManager** | Add/edit aircraft, planning mode, cycle position, down-maintenance dates |
| **frmPlannerSettings** | Adjust forecast dates, flying rates, grouping thresholds without editing sheets |
| **frmTaskRules** | Configure task prefix highlight / pull / extension rules |

## Requirements

- Microsoft Excel with macros enabled (`.xlsm`)
- SITS maintenance data imported per aircraft before running a full refresh

## Development notes

- Aircraft-specific sheets (`ZZ382_IMPORT`, `ZZ388_IMPORT`, etc.) are created dynamically when aircraft are added to the roster.
- The workbook also has Python-in-Excel environment metadata configured, but the core planner logic is implemented in VBA.
- See the **TODO** sheet in the workbook for the current feature backlog.
