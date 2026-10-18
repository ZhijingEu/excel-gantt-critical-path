# Excel Gantt Chart with Dependencies and Critical Path

[![Three project team members discussing a house plan, with a connected Gantt schedule behind them](assets/medium-article-hero.png)](https://zhijingeu.medium.com/building-a-spreadsheet-based-gantt-chart-template-with-python-in-excel-aa57b3947c4e)

**[Read the Medium article: Building a Spreadsheet-Based Gantt Chart Template with Python in Excel](https://zhijingeu.medium.com/building-a-spreadsheet-based-gantt-chart-template-with-python-in-excel-aa57b3947c4e)**

*Why I built it, how the scheduling logic works, and how to get started. Illustration generated with OpenAI.*

Edit a familiar spreadsheet and let Python in Excel calculate linked activity dates and finish float. The Gantt uses worksheet formulas and conditional formatting, so you can inspect the schedule without running a separate Python program or enabling VBA.

By Zhijing Eu.

**Development disclosure:** I conceived this project and defined its requirements. It was co-developed with support from ChatGPT 6.1 Sol Light through an iterative process, with me reviewing changes, testing the workbook and notebook, and making the key design decisions. AI also assisted with documentation and illustrations.

## Why I built this

I wanted a small linked schedule that a team could work through in Excel during a planning discussion. Change a duration or predecessor, and the dates, float and critical highlighting update in the same workbook. Python in Excel handles the dependency calculations, while worksheet formulas and conditional formatting keep the Gantt visible in the cells. There is no VBA or separate application to generate the chart. The aim is to make broad planning assumptions and trade-offs easier to explore together, within the limits of a lightweight tool.

The template is provided free of charge and supports up to 100 task rows, FS/SS/FF relationships with lags, visible early/late dates and total/free float. The predecessor parser has no fixed six-predecessor cap; normal cell and template limits still apply. The companion notebook is optional. Excel users still need an eligible Microsoft 365 environment and internet access; the template being free does not remove Microsoft's licensing requirements.

Other spreadsheets already provide critical-path analysis. For example, [Vertex42's CPM template](https://www.vertex42.com/ExcelTemplates/critical-path-method.html) is also macro-free: its published free-version limits are 25 tasks and six predecessors per task, while its paid Pro version supports 500 tasks and ten predecessors. Its documented dependency model uses basic predecessor/successor links. This project's emphasis is the combination of a free 100-row template, multiple relationship types and an inspectable Python-in-Excel calculation. Vertex42 also supports older Excel versions without Python-in-Excel cloud compute. These are different trade-offs, rather than a claim that this is the first or universally better scheduling spreadsheet. Comparison checked 18 October 2026; see the linked publisher page for changes.

## Start here

**Current release: v1.3.** Free template for up to 100 task rows, with dependency scheduling, late dates, total/free float and critical highlighting. The workbook and notebook have been reviewed in desktop Excel and local Jupyter.

Download [the desktop-checked v1.3 workbook](CritPath_GanttChart_Template_v1.3.xlsx), save a working copy, and open it in Microsoft Excel with Python in Excel available. Keep the supplied example unchanged for reference.

You need internet access and a qualifying Microsoft 365 subscription. Python calculations run in Microsoft's cloud. This workbook does not provide offline Python calculation or local-only data processing. If you already use Microsoft 365 services such as OneDrive, assess this feature within your existing organisation's cloud policies.

See Microsoft's [Python in Excel introduction](https://support.microsoft.com/en-us/excel/python/introduction-to-python-in-excel) and [availability information](https://support.microsoft.com/en-gb/excel/python/python-in-excel-availability).

**Standard versus premium compute:** qualifying subscriptions include standard compute plus a limited monthly allowance of faster premium compute. Microsoft says calculations switch to standard speed when that allowance is exhausted, with premium allowance resetting at the start of each month. Premium compute is not required by this template; standard-speed responsiveness has not been benchmarked. The optional Python in Excel add-on provides premium compute and additional calculation modes. Excel shows notifications for premium use, nearing the allowance, and switching to standard speed; Microsoft's linked guidance does not specify a numeric allowance or remaining-usage counter. The local notebook and local tests do not consume Python-in-Excel compute. Check the linked availability page for current plan/platform and preview conditions.

## A quick walkthrough

1. In **Workplan**, review the house-construction example and identify the editable input columns A:M. Detailed instructions are in the workbook's **README** tab.
2. Enter up to 100 task rows consecutively, beginning at row 3. Leave unused rows at the end. Do not leave blank rows between tasks.
3. Keep TaskIDs unique and use text IDs to preserve leading zeros. Enter the TaskType, duration in weeks, predecessors, and any fixed start constraints. At least one valid constraint date is required to establish the scheduling origin.
4. Allow Python calculation to finish. Check the status banner, calculated dates, finish float, and weekly Gantt. Detailed fixed-date conflict messages appear in column CL, after the chart.
5. Try changing **PREPARE SITE** from 4 to 5 weeks. In the supplied example, its finish and the overall house finish each move one week. Restore the duration when finished.

Sort using the input table only. Do not include calculated columns Q:Z, the Gantt, or warning output in a manual sort. Native testing confirmed that sorting the input table recalculates results into the new row order.

## Inputs and dependencies

| Field | Meaning |
| --- | --- |
| TaskID | Unique text identifier, used by dependency references |
| TaskType | `Activity`, `Milestone`, or `Summary` |
| Duration | Elapsed weeks; milestones are zero duration and hammock spans are calculated |
| Predecessors | Comma-separated TaskID, relationship, and optional lag |
| Constraints | Optional fixed start date; at least one date is required in the schedule |
| Progress | Informational percentage; it does not change schedule calculations |

TaskTitle, descriptions, Stage, and grouping columns help explain or organise the plan. They do not drive the current scheduling calculation.

Dependency examples:

| Entry | Meaning |
| --- | --- |
| `0030 FS` | Start after 0030 finishes |
| `0030 SS+2` | Start at least two weeks after 0030 starts |
| `0030 FF-1` | Finish no earlier than one week before 0030 finishes, subject to other logic |
| `0030 FS, 0040 FS` | Respect both predecessors |
| `0030 SS, 0080 FF` on a Summary | Display a hammock spanning those start/finish anchors |

Negative lags are supported. Blank durations retain the original zero-duration convention, so review them deliberately. Invalid nonblank durations, invalid dates, duplicate IDs, cycles, missing IDs, and gaps produce blocking errors. Summary anchors must refer to Activities or Milestones; nested Summaries are not supported.

## Reading the results

**LateStart/LateFinish** are week offsets, matching EarlyStart/EarlyFinish. **TotalFinishFloat** is late finish minus early finish (equivalently, late start minus early start in this fixed-duration model). **FreeFloat** measures how far a task can move without moving successor early dates or extending the current project finish, respecting dependency types and lags. Negative values remain visible for conflicts. Summary late dates and floats are blank. Activities and Milestones participate in criticality calculations. Summaries are display-only hammocks and are excluded.

| Symbol | Meaning |
| --- | --- |
| X | Activity with effectively zero or negative float |
| O | Activity with positive float |
| ◆ | Milestone, regardless of its critical flag |
| = | Summary/hammock span |

Fixed starts override dependency-derived early starts. If that violates a relationship, the schedule remains visible with a warning. **The backward pass does not pin fixed dates: positive float is not permission to miss a fixed date.** Resolve conflicts before relying on the plan.

If Python reports an error, correct the input and wait for successful recalculation. Do not rely on previously displayed or saved schedule values while an error is present.

## Limitations

Use at your own risk. This is a lightweight planning aid for small schedules, not a comprehensive scheduling or assurance system.

- The presentation is functional and uses text symbols rather than polished interactive bars.
- The template supports 100 task rows and a fixed 63-week visible chart window. Longer plans may calculate dates beyond the displayed timeline.
- Time is expressed in elapsed weeks, without custom work calendars, holidays, or selectable hourly/daily/monthly views.
- Progress is informational. There is no comprehensive actuals, remaining-duration, baseline, or status-date treatment.
- Fixed starts affect early scheduling only; float has the limitation described above.
- Hammocks are anchored spans, not a full hierarchical work-breakdown structure.
- No resource loading/levelling, schedule risk simulation, or formal DCMA compliance assessment is provided.
- Python-in-Excel availability and calculation speed depend on the Microsoft environment.

For larger or operationally complex schedules, evaluate a professional scheduling tool rather than extending this template indefinitely.

## What has been checked

Desktop Excel checks cover live duration changes, constraint warnings, negative-duration and cycle errors, recovery, saving/reopening, 50/100-task examples, and input-table sorting. The author has also confirmed browser co-authoring works. Local automated tests complement those checks. They are not certification of every possible network.

## Optional notebook: a local interactive Gantt

[Critical_Path_Gantt_v1.3.ipynb](Critical_Path_Gantt_v1.3.ipynb) is an explanatory document with runnable Python cells. It runs immediately with a built-in sample table, or reads a CSV or the saved XLSX inputs. It contains the same v1.3 scheduling statements, split into sections explaining validation, dependency parsing, forward/backward passes, hammocks and float. It displays an interactive Plotly chart inside the notebook; standalone HTML export is optional and off by default. The built-in sample, supplied CSV and XLSX all contain the same 23-task house plan, including descriptive fields, so their dates, float and critical flags agree. Sample mode remains self-contained and needs no workbook file.

Start with [sample_schedule.csv](sample_schedule.csv). Required, case-sensitive headers are `TaskID`, `TaskType`, `Duration`, `Predecessors`, `Constraints`, in any order. `TaskTitle` is recommended and defaults to TaskID when absent; other descriptive columns are optional. Blank constraint cells mean no override. Enter at least one date, preferably `YYYY-MM-DD`; durations and lags are elapsed weeks. Keep IDs as text. Predecessor lists containing commas must be quoted as one CSV field. When HTML export is enabled, the notebook also saves its built-in sample as a CSV template.

| Use the XLSX when… | Use the notebook when… |
| --- | --- |
| You want to type into Excel and have Python in Excel recalculate | You want an annotated view of the calculation or a more readable interactive chart |
| You have an eligible Microsoft 365 environment and internet access | You can run local Python/Jupyter; no Python-in-Excel licence is needed |
| You want the worksheet's weekly symbol Gantt | You want hover details, pan/zoom, milestone diamonds and optional dependency connectors |
| You want changes to recalculate within Excel | You can save workbook edits and rerun the notebook to refresh the chart |

The notebook retains the same 100-task limit and scheduling limitations, including forward-only fixed starts and informational progress. Its continuous date axis does not introduce working-day calendars or a different duration model. It reads literal inputs from **Workplan / Table1**, never saves over the source workbook, and recalculates from inputs instead of trusting saved Excel outputs. Formula inputs in A:M are rejected because openpyxl cannot evaluate them; enter literal values in a working copy. Keep IDs stored as text to preserve leading zeros.

### Run locally in an isolated environment

Open a terminal in the downloaded repository folder. These commands create a local `.venv` and keep Jupyter/IPython settings inside the repository. They do not change Anaconda or install packages globally. The tested setup is Python 3.13.6 on Windows.

**Windows PowerShell:**

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:PYTHONNOUSERSITE = "1"
$env:JUPYTER_CONFIG_DIR = "$PWD/.jupyter/config"
$env:JUPYTER_DATA_DIR = "$PWD/.jupyter/data"
$env:JUPYTER_RUNTIME_DIR = "$PWD/.jupyter/runtime"
$env:IPYTHONDIR = "$PWD/.ipython"
$env:JUPYTER_PATH = "$PWD/.venv/share/jupyter"
.\.venv\Scripts\python.exe -m ipykernel install --prefix .venv --name python3 --display-name "Python (Excel Gantt)"
.\.venv\Scripts\python.exe -m notebook Critical_Path_Gantt_v1.3.ipynb --ServerApp.ip=127.0.0.1
```

**macOS/Linux terminal** (instructions provided; not independently tested):

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
export PYTHONNOUSERSITE=1
export JUPYTER_CONFIG_DIR="$PWD/.jupyter/config"
export JUPYTER_DATA_DIR="$PWD/.jupyter/data"
export JUPYTER_RUNTIME_DIR="$PWD/.jupyter/runtime"
export IPYTHONDIR="$PWD/.ipython"
export JUPYTER_PATH="$PWD/.venv/share/jupyter"
.venv/bin/python -m ipykernel install --prefix .venv --name python3 --display-name "Python (Excel Gantt)"
.venv/bin/python -m notebook Critical_Path_Gantt_v1.3.ipynb --ServerApp.ip=127.0.0.1
```

Download packages once with internet access. Subsequent local calculation and HTML viewing do not require Microsoft's Python-in-Excel service. Use this environment's **Python (Excel Gantt)** kernel. Registering with `--prefix .venv` keeps the kernel local; do not substitute `--user`.

1. Leave `INPUT_MODE = "sample"` for the built-in house plan, or choose `"csv"` / `"xlsx"` and set the corresponding file path. Keep the default relative paths if your inputs are alongside the notebook.
2. Choose **Kernel → Restart Kernel and Run All Cells**. Resolve errors and fixed-start warnings before relying on results. Saved output can remain visible after an error.
3. Explore the inline chart; toggle **Dependencies (toggle)** in its legend to show links.
4. Near the chart, **step 13** has `EXPORT_HTML = False` by default. Set it to `True`, run that setting cell, then the export cell to save `outputs/critical_path_gantt.html` and a sample CSV. The output folder is created only when needed; it is not distributed. The exported CSV intentionally repeats the built-in sample so it can serve as a template.
5. After changing Excel or CSV inputs, save them and rerun the calculation. To stop Jupyter, use Ctrl+C in its terminal. Closing a browser tab alone may leave the server running.

The HTML embeds the chart library and works without a hosted app or CDN. It contains schedule details and hover information, so treat it as sensitively as the workbook. Hover/zoom is interactive; the exported HTML does not edit inputs or recalculate the schedule. A workbook in OneDrive may still be synchronised by OneDrive even though the notebook calculation itself is local.

The [Plotly export documentation](https://plotly.com/python/interactive-html-export/) explains standalone HTML; the [IPython kernel guide](https://ipython.readthedocs.io/en/stable/install/kernel_install.html) describes environment-scoped kernels.

## Files in this repository

| File | Purpose |
| --- | --- |
| `CritPath_GanttChart_Template_v1.3.xlsx` | Excel template, instructions and readable Python code copies |
| `Critical_Path_Gantt_v1.3.ipynb` | Self-contained annotated calculation and interactive chart |
| `sample_schedule.csv` | The same 23-task house example for CSV mode |
| `requirements.txt` | Pinned dependencies for a local notebook environment |
| `LICENSE` | MIT licence |

The notebook embeds its input readers, scheduling code and chart helpers. It does not require a `src/` folder, a development checkout or the spreadsheet in sample mode. Local environments, configuration and exports are ignored by Git.

## Verification and limitations

The workbook's **Python Formulas In WorkplanTab** copies match its actual Python cells. The calculation has been checked against saved Excel results and independent dependency examples, including lags, milestones, fixed-date warnings, invalid inputs, cycles and 100-task schedules. A separate clean-folder execution check covers sample, CSV and XLSX notebook input, inline chart output and HTML export. Development helpers and tests are maintained privately and are not required to use this release.

These checks do not certify every possible network or formal schedule-quality compliance. See the scope and limitations above.

## Attribution and licence

The house-construction sample is adapted from the **Example House Construct Plan** supplied with Oracle Primavera Risk Analysis. The adapted example is redistributed with permission. This project is not affiliated with or endorsed by Oracle.

The project code and original workbook/documentation content are released under the [MIT licence](LICENSE). Attribution does not imply ownership of Oracle's original material. Python in Excel, Microsoft 365 and the notebook libraries are supplied under their respective providers' terms.
