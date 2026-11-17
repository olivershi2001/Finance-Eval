# Customer Lab

A local product around the existing customer-cube evaluation engine. The first
workflow is deliberately narrow: monthly customer/product revenue to a reviewable
quarterly and annual analysis, with original source evidence and saved human judgment.

```text
Excel upload / synthetic example
              ↓
Validate columns, amounts and file limits
              ↓
Calculate reference customer cube
              ↓
Choose calculated observations, AI explanation, or independent model calculation
              ↓
Source checks / bridge reconciliation / model comparison
              ↓
Optional separate model judge
              ↓
Review output → inspect source → save an issue or resolution
              ↓
Download Excel / CSV / complete run with notes
```

## Working with the app

1. Run `python3 app.py` from the project directory and open
   http://127.0.0.1:8765.
2. Start with the synthetic example. Its names and financials are fabricated;
   no employer or deal data is used.
3. Confirm what Amount means. Monthly revenue is summed; month-end balances
   use the calendar period end. Years start in January.
4. Choose a mode. The calculated mode is usable without an API key. AI explanation
   uses the configured candidate model. Model testing asks the model to perform
   the calculations and compares its output with the engine for the uploaded data.
5. On Overview, inspect revenue, retention, and the bridge. Click an insight's
   evidence link or an Inspect button in Customer cube to view relevant source rows.
6. Source evidence includes all months for the selected customer/product because
   opening balances, churn, and rolling TTM measures need historical observations.
   Row numbers refer to the original Excel worksheet, even if it contains blank rows.
7. In Review notes, document an issue, your correction, or the outcome of a review.
   A row inspection attaches the note to that exact four-part output key. Notes
   append to history; they never edit model output or imply that a model retrained.
8. Download the full run to capture results and notes together. The original
   source file is included, so treat the bundle like the source workbook.

The latest review note supplies the run's review-status label. This is an audit
trail, not issue-by-issue tracking: a resolution note does not automatically close
every previously flagged issue. Multiple notes and conflicting findings remain
visible. New runs are immutable snapshots with separate review histories.

## What the scores mean

Reference mode has no model score. Bridge checks establish arithmetic
reconciliation, not economic validity. Source checks identify duplicate keys and
missing source months, while revenue meaning and period completeness explicitly
remain review items.

Model mode uses the existing weighted scorer: structure 10, coverage 20,
revenue-component accuracy 40, retention 20, reconciliation 10. Pass threshold:
95. The workbench tightens the shared comparison tolerance to 1e-6 absolute and
relative, including percentage values. Bridge display checks allow 0.01 source
units / 1e-6 relative tolerance. Source units are not assumed to be dollars.

The reference engine is the project's existing implementation. Agreement with it
is not independent financial adjudication or evidence of general model capability.
The frozen synthetic answer set and existing CLI remain available separately.
Model prose without structured references is labeled uncited. AI explanations
are checked for valid row citations, but their numeric claims and interpretation
are not automatically proven by citation existence.

The optional AI judge applies the existing anchored finance rubric to the workbook.
The app displays the judge separately; complete criterion evidence is included
in `result.json`. It is an additional model opinion, not a real human review.

## Failure handling and local operation

- Runs execute in two background workers. Progress can be reopened from history.
- Invalid data marks a run failed and preserves its source fingerprint.
- Failed AI explanations leave the calculated analysis usable with a visible warning.
- Failed model calculations never silently become reference-model successes.
- Failed Excel exports retain JSON and CSV. Failed judges retain prior analysis.
- A server restart marks unfinished runs interrupted; retrying creates a new run.
- Atomic JSON writes and a lock protect run and note storage within one process.
- Localhost binding, Host/Origin checks, escaped interface text, CSV formula
  escaping, ZIP expansion limits, and allowlisted downloads protect the local workflow.
- Provider errors are redacted before display. Credentials remain in environment
  or Keychain, outside run artifacts.
- Python 3.10+ serves the app without new Python dependencies. The existing
  spreadsheet builder uses the bundled Node runtime and `@oai/artifact-tool`.
- The original CLI and pre-existing working-tree edits are retained.

## Build and verification map

| Component | Purpose |
| --- | --- |
| `app.py` | Local HTTP app, saved runs, worker pipeline, evidence, feedback, exports |
| `web/index.html` | Upload, analysis, source, checks, and review screens |
| `web/style.css` | Responsive desktop/mobile interface |
| `web/app.js` | Upload, polling, filtering, source links, downloads, review notes |
| `customer_cube.py` | Existing deterministic customer revenue engine |
| `model_eval.py` | Existing candidate prompt, strict validation, Excel export |
| `llm_judge.py` | Existing provider adapter and optional rubric-based judge |
| `tests/test_workbench.py` | Upload-to-review HTTP behavior, failures, source preservation, numeric errors |

Run all checks with `python3 -m unittest discover -s tests -v`. The HTTP integration
test needs permission to bind localhost. AI paths are covered with controlled
responses; any live trial must be reported separately from these software tests.

Verified on November 15, 2026:

- All 46 tests passed, including an HTTP upload → analysis → row-linked review
  note → downloadable run bundle, bad-origin rejection, and failure handling.
- A live calculated run processed the 1,000-row synthetic source into 604 cube rows.
- A live AI-assisted run used the existing configured `gpt-5.6-sol` provider and
  returned four insights with valid cube-row citations. It completed without warnings.
- Browser QA followed a product citation to 194 source rows, saved an explicitly
  labeled test review note, and checked the available download links.
- The AI-assisted run exported Excel successfully. All seven sheet previews were
  visually reviewed; the workbook's formula-error scan found zero matches.
- Model-calculation and separate-judge paths were exercised with controlled
  responses in tests. Neither was measured as a live model benchmark in this build.

The v2 learning workspace now adds model comparisons, issue-resolution history,
reviewer-authored regression cases, and model-call timing. See [LEARNING.md](LEARNING.md)
for scope and limitations. Configurable fiscal calendars, a broader adjudicated
gold set, and token-cost measurement remain future work.
