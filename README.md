# Customer Cube Eval

## Interactive customer-analysis app

Start the local workbench from this directory:

```bash
python3 app.py
```

Open **http://127.0.0.1:8765**. Choose **Try synthetic example** to run the
included 1,000-row synthetic workbook, or upload a workbook with the same four
columns. The app builds quarterly and annual revenue bridges, retention, and
customer/product cuts. Inspect output rows against original Excel row numbers,
save review notes, and download Excel, CSV, or a complete reproducibility bundle.
Run history and notes survive app restarts. Use `--port 8766` for another port.

Three modes are available:

- **Calculated analysis:** existing deterministic engine, no API call. No model
  performance score is assigned to these results.
- **Calculated analysis + AI explanation:** the engine calculates; the configured
  model explains the results with citations to existing output rows. Citation
  existence is checked; the prose still needs human review.
- **Test a model:** the configured model generates the cube; the engine checks
  it against a reference computed for this exact upload. The workbench uses
  `1e-6` absolute and relative comparison tolerances, including rates. This mode
  currently supports calendar-year monthly revenue (`period_sum`) only.

An optional separate AI judge reviews the exported workbook with the existing
finance rubric. Its judgment is shown separately from the engine comparison.
The two calls may use the same configured model; they do not share a response.

AI calls use the existing `candidate_model_config.json` and
`llm_judge_config.json`. Supply `OPENAI_API_KEY` in the server environment or use
the existing configured macOS Keychain service; keys never go to the browser.
The app does not install credentials or change either model configuration.
AI modes send source or derived analysis data to the configured provider.
The default calculated mode stays local.

Uploads are limited to 5 MB / 5,000 source rows / 30 MB expanded XLSX contents.
Non-finite and negative amounts are rejected. Missing customer/product months
are zero, duplicates are summed, returning customers are New, and boundary
periods may be partial. Confirm these assumptions before using real exports.
The app supports monthly-flow and month-end-balance conventions; it does not
infer currency, source completeness, or the business cause of a change.

Runs are stored under ignored `outputs/workbench/<run-id>/`; full bundles
contain the original upload, normalized source evidence, calculation results,
available model prompts/responses, exports, and review notes. Each upload is
SHA-256 fingerprinted. The app binds to localhost and is intended for one local
user; hosted authentication, multi-user permissions, and public deployment are
outside this version. Excel exports are calculation snapshots: rerun the app
after changing source data. If the bundled Node/artifact runtime is unavailable,
CSV and JSON remain usable and the app reports the Excel export failure.

See [WORKBENCH.md](WORKBENCH.md) for the workflow, validation, and implementation map.

For a plain-English technical walkthrough of the entire project, see
[`HOW_IT_WORKS.md`](HOW_IT_WORKS.md).

This project evaluates whether a language model can turn one raw Excel workbook
into a reconciled SaaS customer cube. The input contract is intentionally narrow:
the first worksheet must contain exactly these four columns:

- `Customer name`
- `Month`
- `Amount`
- `Product`

The automated path calls the model under test, saves its exact structured
response, renders that response into a reviewable Excel workbook, and sends that
generated workbook to a separate LLM judge. You no longer need to create or
supply a candidate workbook yourself.

## Run the complete model eval

```bash
python3 model_eval.py run fixtures/customer_cube_raw.xlsx
```

That one command:

1. sends the raw workbook data to the candidate model configured in
   `candidate_model_config.json`;
2. saves the model's exact response;
3. creates `outputs/model_eval/latest/candidate_output.xlsx` from that response;
4. calls the independent judge configured in `llm_judge_config.json`; and
5. saves the validated score in
   `outputs/model_eval/latest/validated_report.json`.

Open the model-generated workbook and score:

```bash
open outputs/model_eval/latest/candidate_output.xlsx
python3 -m json.tool outputs/model_eval/latest/validated_report.json | less
```

Press `q` to exit the report view. The candidate and judge configurations are
separate on purpose, so you can change the model being tested without changing
the grader. The candidate model receives the task definitions but never sees
the grading rubric.

To replay and grade a response you already saved, without calling the candidate
model again:

```bash
python3 model_eval.py run fixtures/customer_cube_raw.xlsx \
  --candidate-response /path/to/saved_model_response.json
```

## Quick start

Python 3.10+ is sufficient; the eval runner has no third-party dependencies.

```bash
python eval_runner.py evaluate
python -m unittest discover -s tests -v
```

The bundled reference candidate should score `100.00/100` across 604 gold rows.

Open the finished reference Excel workbook directly:

```bash
open outputs/customer_cube_eval/customer_cube_reference_output.xlsx
```

It includes an executive summary, quarterly and annual cubes, data-quality
checks, metric definitions, the finance-review rubric, and the 1,000-row source
data tab.

## Grade an existing candidate workbook (optional legacy route)

Use `llm_judge.py` when you already have a candidate Excel workbook and want
another language model to grade it directly from the raw data and a rubric. The
judge receives:

- the raw customer-revenue workbook;
- the candidate's Excel output; and
- the default editable rubric in `rubrics/customer_cube_finance_template.json`.

It does **not** receive or require the frozen expected JSON or reference output.
The model returns 0–4 criterion ratings, explanations, and workbook citations.
The Python harness converts each anchored rating into weighted points, adds the
points itself, and assigns `PASS`, `REVIEW`, `FAIL`, or `FAIL - CRITICAL`. Any
declared critical failure forces a fail even when the numerical score is above
the threshold. Validated reports also record SHA-256 fingerprints for
reproducibility.

### Manual route: model-agnostic

Build one self-contained prompt:

```bash
python3 llm_judge.py prepare \
  fixtures/customer_cube_raw.xlsx \
  /path/to/candidate_output.xlsx \
  --output outputs/llm_judge/judge_prompt.md
```

Paste that prompt into the model you want to use. Save its JSON-only response as
`outputs/llm_judge/model_response.json`, then validate and score it:

```bash
python3 llm_judge.py validate \
  outputs/llm_judge/model_response.json \
  --report outputs/llm_judge/validated_report.json
```

The `prepare` command is local-only: it makes no API call and sends no workbook
data anywhere.

### Automated OpenAI route

In `llm_judge_config.json`, set `provider` to `openai_responses` and choose the
model. Keep the API key in the environment:

```bash
export OPENAI_API_KEY="your-key"
python3 llm_judge.py grade \
  fixtures/customer_cube_raw.xlsx \
  /path/to/candidate_output.xlsx \
  --report outputs/llm_judge/validated_report.json \
  --prompt-output outputs/llm_judge/judge_prompt.md \
  --raw-response-output outputs/llm_judge/model_response.json
```

Never put an API key in the JSON config. In automated mode the prompt contains
the workbook contents, so use only a provider and data policy appropriate for
the source data. On macOS, the grader can instead read the key from the Keychain
service named by `keychain_service`; an environment variable takes precedence.

Both automated configs set `store` to `false`, so the Responses API request asks
OpenAI not to retain the response object for later retrieval.

### Local or other model command

Set `provider` to `command` and provide a command array. The prompt is delivered
on standard input, `{model}` is replaced with the configured model name, and the
command must write only its JSON response to standard output. For example:

```json
{
  "provider": "command",
  "model": "qwen3:32b",
  "command": ["ollama", "run", "{model}"]
}
```

Start by editing the rubric's assumptions, metric definitions, anchored
criteria, weights, critical failures, and thresholds. Criterion weights must
total 100. The finance-template rubric scores numerical correctness (30),
movement methodology (15), data-quality detection (10), traceability (15),
analytical judgment (15), practical usefulness (10), and presentation (5). It
passes at 85 and routes scores from 70 through 84.99 to human review.

LLM judgments can vary. Before using this as a benchmark, grade several known
good and known bad workbooks with a fixed model, inspect its cited evidence, and
tune the rubric until the model's deductions match the behavior you intend to
measure.

### Rubric library and imports

List the available rubric presets:

```bash
python3 llm_judge.py list-rubrics
```

The bundled presets are:

| Rubric | Intended use |
|---|---|
| `customer_cube_finance_template.json` | Default 0–4 finance-review rubric imported from the supplied Customer Cube template. |
| `customer_cube_llm_judge.json` | Original balanced point-based scorecard. |
| `customer_cube_strict_accuracy.json` | Heavier weighting on movements and retention; 90-point pass threshold. |
| `customer_cube_auditability.json` | Heavier weighting on traceability and reviewer usability. |
| `customer_cube_flexible_output.json` | More credit for flexible layouts and dimensional cutting approaches. |

Select one with `--rubric`:

```bash
python3 llm_judge.py grade \
  fixtures/customer_cube_raw.xlsx \
  /path/to/candidate_output.xlsx \
  --rubric rubrics/customer_cube_finance_template.json \
  --report outputs/llm_judge/finance_review_report.json
```

The four variants import the balanced rubric with `extends`. The finance
template replaces the scorecard and adds a 0–4 rating scale, review threshold,
critical failures, and dataset variants; the other three change selected
weights and settings. To import another rubric, add a JSON file under
`rubrics/` and either define a complete 100-point rubric or extend an existing
one. Every resolved rubric must still contain unique criterion IDs and weights
totaling exactly 100.

## Generic Excel-output scoring

Use `workbook_scorer.py` when the answer key and candidate are both Excel
workbooks and you do not want customer-cube-specific logic:

```bash
python3 workbook_scorer.py \
  /path/to/expected_output.xlsx \
  /path/to/candidate_output.xlsx
```

The expected workbook is the gold answer key. Every nonblank gold cell is
compared with the same worksheet and cell coordinate in the candidate by
default. Formatting, colors, widths, and other presentation choices are
ignored. The score allocates 10 points to required worksheets, 20 points to
expected-cell coverage, and 70 points to value accuracy.

Edit `workbook_score_config.json`, or copy it and pass the copy with `--config`,
to reuse the scorer for different workbook shapes:

```bash
python3 workbook_scorer.py gold.xlsx candidate.xlsx \
  --config my_workbook_eval.json \
  --json-report score_report.json
```

The main controls are:

| Setting | Purpose |
|---|---|
| `include_sheets` / `exclude_sheets` | Select the gold tabs that count. |
| `sheet_map` | Map a gold tab to a differently named candidate tab. |
| `ranges` | Score only configured A1 ranges on each gold tab. |
| `ignore_ranges` | Exclude notes, checks, timestamps, or volatile cells. |
| `sheet_weights` | Give an entire tab more or less importance. |
| `range_weights` | Give important outputs, such as headline KPIs, extra weight. |
| `absolute_tolerance` / `relative_tolerance` | Control acceptable numeric differences. |
| `formula_mode` | Compare cached values, exact formulas, or both. |
| `penalize_extra_cells` / `penalize_extra_sheets` | Make the comparison strict. |

For example, this compares only output tables, accepts a differently named
quarterly tab, and triples the importance of annual retention outputs:

```json
{
  "include_sheets": ["Quarterly", "Annual"],
  "sheet_map": {"Quarterly": "Qtr Output"},
  "ranges": {
    "Quarterly": ["A4:M200"],
    "Annual": ["A4:M100"]
  },
  "ignore_ranges": {
    "Quarterly": ["N4:N200"],
    "Annual": ["N4:N100"]
  },
  "range_weights": [
    {"sheet": "Annual", "range": "K5:M100", "weight": 3.0}
  ],
  "absolute_tolerance": 0.01,
  "relative_tolerance": 0.000001,
  "formula_mode": "values",
  "pass_threshold": 95.0
}
```

This comparison is coordinate-based, so row and column order matter. If two
outputs can legitimately sort rows differently, normalize their order first or
use the customer-cube semantic scorer, which keys rows by business fields.
Legacy binary `.xls` files should be saved as `.xlsx`; `.xlsx` and `.xlsm` are
supported directly.

To run the reference logic repeatedly on another workbook:

```bash
python eval_runner.py analyze \
  --input /path/to/raw_customer_data.xlsx \
  --output /tmp/customer_cube.json
```

Choose one or several cuts by repeating `--slice`:

```bash
python eval_runner.py analyze \
  --input /path/to/raw_customer_data.xlsx \
  --slice overall \
  --slice product \
  --output /tmp/product_cube.json
```

To score a previously saved candidate result:

```bash
python eval_runner.py score /tmp/candidate_cube.json
```

## Candidate contract

Replace `build_customer_cube` in `candidate.py` with the model, agent, or
implementation under test. It receives normalized `RevenueRecord` objects and a
`CubeConfig`, and returns `CandidateResult(rows=...)`.

Each row is uniquely keyed by:

```text
(grain, period, slice_type, slice_value)
```

Required numeric outputs are:

```text
opening_revenue, new, upsell, downsell, cross_sell, churn,
closing_revenue, gross_retention, net_retention, ttm_gross_retention
```

`downsell` and `churn` are positive loss magnitudes. Retention is `null` when
opening revenue is zero.

## Metric definitions

Classification happens at customer × product, but the customer's total
relationship determines New versus Churn.

| Metric | Definition |
|---|---|
| New | Customer total was zero in the prior comparable period and is positive now; all current product revenue is New. |
| Upsell | Existing customer and existing product; current product revenue exceeds prior product revenue. |
| Downsell | Customer remains active, while an existing product decreases or disappears. |
| Cross-sell | Existing customer begins a product that was absent in the prior period. |
| Churn | Customer total falls to zero; all prior product revenue is Churn. |
| GRR | `(Opening − Downsell − Churn) / Opening`, floored at 0% and capped at 100%. |
| NRR | `(Opening + Upsell + Cross-sell − Downsell − Churn) / Opening`, floored at 0% with no upper cap. New is excluded. |
| TTM GRR | In `period_sum` mode, GRR compares the rolling 12-month revenue window with the preceding 12 months. In `period_end` mode, it compares period-end revenue with the same month one year earlier. |

The bridge must always reconcile:

```text
Closing = Opening + New + Upsell + Cross-sell − Downsell − Churn
```

Because Reactivation was not requested as a separate metric, a customer with
zero revenue in the immediately prior period is treated as New when it returns.

## Period and amount conventions

Defaults live in `eval_config.json`:

- `amount_mode: period_sum` treats `Amount` as monthly recognized recurring
  revenue and sums it into quarterly and annual periods.
- Change it to `period_end` if `Amount` is a month-end MRR/ARR balance; quarterly
  and annual outputs will then use the last month in each period.
- `fiscal_year_start_month: 1` creates calendar quarters and years. Set `4` for
  an April fiscal-year start, for example.
- Missing customer/product/month rows are treated as zero. Duplicate rows for
  the same key are summed. Negative amounts are rejected.

These are the two business choices worth confirming before this becomes a
production benchmark: whether `Amount` is a monthly flow or month-end balance,
and whether returning customers should remain folded into New or receive a
separate Reactivation category.

## Scoring

| Section | Weight | What is graded |
|---|---:|---|
| Structure | 10% | Valid grains/cuts, finite numbers, unique row keys |
| Coverage | 20% | All expected quarterly/annual and slice rows; no extras |
| Movement accuracy | 40% | Opening, closing, New, Upsell, Downsell, Cross-sell, Churn |
| Retention accuracy | 20% | GRR, NRR, rolling TTM GRR, including correct `null`s |
| Reconciliation | 10% | Revenue bridge and retention formulas tie internally |

Default pass threshold is 95. Dollar tolerance is `$0.01`; rates use a `1e-6`
relative tolerance. The CLI reports the first missing/extra keys and value
mismatches, plus accuracy by period grain and cut.

## Fixture coverage

The single bundled raw workbook contains exactly 1,000 fully synthetic data
rows spanning January 2022 through December 2025, with 50 dummy customers and
five dummy products. It covers:

- full new-logo acquisition and full-customer churn;
- product upsell and downsell;
- product additions (cross-sell) and product removal while the customer remains;
- product reallocation inside one customer;
- duplicate customer/month/product source rows;
- quarterly, annual, rolling-TTM, overall, product, and customer cuts.

The expected JSON is frozen in `scenarios/customer_cube_expected.json`. Do not
generate expected values at eval runtime; freezing the gold set prevents a bug
in the candidate from silently changing the answer key.

## Project map

```text
candidate.py                         swappable candidate entry point
customer_cube.py                     deterministic reference implementation
eval_runner.py                       evaluate / analyze / score CLI
workbook_scorer.py                   generic XLSX-against-XLSX scorer
models.py                            typed input, output, and report contracts
scoring.py                           weighted scorer and diagnostics
xlsx_reader.py                       strict, dependency-free XLSX input adapter
eval_config.json                     conventions, slices, tolerances, threshold
workbook_score_config.json           generic workbook comparison settings
llm_judge.py                          answer-key-free LLM grading harness
llm_judge_config.json                 manual/API/command model adapter settings
rubrics/customer_cube_finance_template.json default anchored finance-review rubric
rubrics/customer_cube_llm_judge.json balanced base judge rubric
fixtures/customer_cube_raw.xlsx      the one raw Excel fixture
scenarios/customer_cube_expected.json frozen gold rows
tests/                               taxonomy, tie-out, ingestion, scoring tests
outputs/customer_cube_eval/          visually verified reference workbook
```

`scripts/generate_gold.py` is deliberately separate from the runtime scorer.
Use it only when intentionally versioning a fixture change, then review the
diff and the reference workbook before accepting the new gold set.
