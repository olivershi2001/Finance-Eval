"""Saved model comparisons, issue history, and reviewer-authored regression cases.

No provider calls happen here. Original run results are read, never rewritten.
"""
from __future__ import annotations

from collections import Counter
import hashlib
import json
import math
from pathlib import Path
import re
import secrets
from datetime import datetime, timezone

from scoring import MOVEMENT_FIELDS, RETENTION_FIELDS

KEY_FIELDS = ("grain", "period", "slice_type", "slice_value")
METRICS = (*MOVEMENT_FIELDS, *RETENTION_FIELDS)


def timestamp():
    return datetime.now(timezone.utc).isoformat()


def fingerprint(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()).hexdigest()


def valid_model(value):
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._:/-]{0,119}", value):
        raise ValueError("Enter a model ID using letters, numbers, dots, dashes, slashes or colons (up to 120 characters)")
    return value


class Lab:
    def __init__(self, store, lock):
        self.store = Path(store)
        self.lock = lock

    def read(self, path):
        with self.lock:
            return json.loads(path.read_text(encoding="utf-8"))

    def write(self, path, data):
        with self.lock:
            path.parent.mkdir(parents=True, exist_ok=True)
            temp = path.with_suffix(".tmp")
            temp.write_text(json.dumps(data, indent=2, allow_nan=False), encoding="utf-8")
            temp.replace(path)

    def run(self, run_id):
        if not isinstance(run_id, str) or not re.fullmatch(r"[a-f0-9]{24}", run_id):
            raise ValueError("Invalid run identifier")
        directory = self.store / run_id
        meta = self.read(directory / "run.json")
        result = self.read(directory / "result.json") if (directory / "result.json").exists() else None
        return directory, meta, result

    def comparison(self, identity):
        if not re.fullmatch(r"[a-f0-9]{24}", identity):
            raise ValueError("Invalid comparison identifier")
        batch = self.read(self.store / "comparisons" / (identity + ".json"))
        items = []
        for run_id in batch["run_ids"]:
            _, meta, result = self.run(run_id)
            score = result.get("score") if result else None
            items.append({"id": run_id, "model": meta.get("selected_model", meta.get("model")),
                          "status": meta["status"], "stage": meta["stage"], "error": meta.get("error"),
                          "score": score["score"] if score else None,
                          "passed": score["passed"] if score else None,
                          "model_seconds": meta.get("model_seconds"), "total_seconds": meta.get("total_seconds"),
                          "prompt_sha256": meta.get("prompt_sha256"),
                          "config_sha256": meta.get("config_sha256"),
                          "failures": sum(i["category"] != "reviewer" for i in self.run_failures(run_id))
                          if meta["status"] in ("complete", "failed") else None})
        settled = all(x["status"] in ("complete", "failed") for x in items)
        hashes = {x["prompt_sha256"] for x in items if x["prompt_sha256"]}
        all_recorded = bool(items) and all(x["prompt_sha256"] for x in items)
        batch.update(candidates=items, status="complete" if settled else "running",
                     same_prompt=(len(hashes) == 1) if all_recorded else None,
                     caveat="Single trials on one dataset; scores are not a general model ranking. Cost is not measured.")
        return batch

    def comparisons(self):
        return [self.comparison(p.stem) for p in sorted((self.store / "comparisons").glob("*.json"), reverse=True)]

    def create_comparison(self, body, create_run):
        directory, meta, _ = self.run(body.get("source_run_id"))
        models = body.get("models")
        if not isinstance(models, list) or not 2 <= len(models) <= 4:
            raise ValueError("Compare between two and four different model IDs")
        models = [valid_model(m) for m in models]
        if len(set(models)) != len(models):
            raise ValueError("Choose different model IDs for this comparison")
        if meta["amount_mode"] != "period_sum":
            raise ValueError("Model comparison currently supports monthly revenue only")
        data = (directory / "source.xlsx").read_bytes()
        if hashlib.sha256(data).hexdigest() != meta["source_sha256"]:
            raise ValueError("Saved source fingerprint changed; upload it again")
        identity = secrets.token_hex(12)
        batch = {"id": identity, "created_at": timestamp(), "source_run_id": meta["id"],
                 "source_sha256": meta["source_sha256"], "filename": meta["filename"],
                 "amount_mode": meta["amount_mode"], "models": models, "run_ids": []}
        # Persist before scheduling: a failed request can still be recovered from history.
        destination = self.store / "comparisons" / (identity + ".json")
        self.write(destination, batch)
        for model in models:
            run = create_run(data, meta["filename"], mode="model", amount_mode=meta["amount_mode"],
                             demo=meta.get("demo", False), selected_model=model, comparison_id=identity)
            batch["run_ids"].append(run["id"])
            self.write(destination, batch)
        return self.comparison(identity)

    def run_failures(self, run_id):
        directory, meta, result = self.run(run_id)
        failures = []
        def add(category, row_key, metric, expected, actual, detail, tag=""):
            identity = fingerprint([run_id, category, row_key, metric, tag])[:24]
            failures.append({"id": identity, "run_id": run_id, "model": meta.get("selected_model", meta.get("model")),
                             "filename": meta["filename"], "category": category, "row_key": row_key,
                             "metric": metric, "expected": expected, "actual": actual, "detail": detail,
                             "created_at": meta["created_at"], "status": "open", "history": []})
        if meta["status"] == "failed":
            add("execution", None, None, None, None, meta.get("error", "Run failed"))
        if result and meta["mode"] == "model":
            expected = {tuple(r[k] for k in KEY_FIELDS): r for r in result["reference_rows"]}
            actual = {tuple(r[k] for k in KEY_FIELDS): r for r in result["rows"]}
            counts = Counter(tuple(r[k] for k in KEY_FIELDS) for r in result["rows"])
            for key in sorted(expected.keys() - actual.keys()):
                add("coverage", list(key), None, None, None, "Expected output row is missing")
            for key in sorted(actual.keys() - expected.keys()):
                add("coverage", list(key), None, None, None, "Unexpected output row")
            for key, count in counts.items():
                if count > 1:
                    add("structure", list(key), None, 1, count, "Duplicate output row key")
            for difference in result["differences"]:
                field = difference["metric"]
                add("retention" if field in RETENTION_FIELDS else "revenue", difference["row_key"], field,
                    difference["expected"], difference["actual"], "Model value differs from the reference engine")
        notes = self.read(directory / "feedback.json")
        for note in notes:
            if note["status"] == "flagged":
                add("reviewer", note.get("row_key"), None, None, None, note["note"], note["id"])
        return failures

    def issues(self):
        updates = self.read(self.store / "issue_history.json") if (self.store / "issue_history.json").exists() else {}
        result = []
        for file in sorted(self.store.glob("*/run.json")):
            for issue in self.run_failures(file.parent.name):
                events = updates.get(issue["id"], [])
                issue["history"] = events
                if events:
                    issue["status"] = events[-1]["status"]
                result.append(issue)
        return sorted(result, key=lambda x: x["created_at"], reverse=True)

    def update_issue(self, identity, body):
        if identity not in {i["id"] for i in self.issues()}:
            raise ValueError("Issue not found")
        status = body.get("status")
        note = body.get("note", "")
        if status not in ("open", "investigating", "resolved", "dismissed"):
            raise ValueError("Choose a supported issue status")
        if not isinstance(note, str) or not 1 <= len(note.strip()) <= 4000:
            raise ValueError("Explain this status change in 1–4,000 characters")
        file = self.store / "issue_history.json"
        with self.lock:
            history = self.read(file) if file.exists() else {}
            events = history.setdefault(identity, [])
            events.append({"status": status, "note": note.strip(), "created_at": timestamp()})
            self.write(file, history)
        return events

    def cases(self):
        cases = [self.read(p) for p in (self.store / "cases").glob("*/case.json")]
        for case in cases:
            if fingerprint({k: v for k, v in case.items() if k != "case_sha256"}) != case.get("case_sha256"):
                raise ValueError("Regression case fingerprint changed; restore the original case file")
        return sorted(cases, key=lambda x: x["created_at"], reverse=True)

    def create_case(self, body):
        directory, meta, result = self.run(body.get("run_id"))
        if meta["status"] != "complete" or not result:
            raise ValueError("Create a case from a completed analysis")
        notes = self.read(directory / "feedback.json")
        note = next((n for n in notes if n["id"] == body.get("note_id")), None)
        if note is None or not note.get("row_key"):
            raise ValueError("Choose a saved review note attached to an output row")
        assertion = body.get("assertion", "metric_equals")
        if assertion not in ("metric_equals", "row_absent", "row_present"):
            raise ValueError("Choose a supported test assertion")
        metric = body.get("metric") if assertion == "metric_equals" else None
        if assertion == "metric_equals" and metric not in METRICS:
            raise ValueError("Choose a supported numeric metric")
        if assertion == "metric_equals" and "expected" not in body:
            raise ValueError("Enter an explicit expected value")
        expected = body.get("expected") if assertion == "metric_equals" else None
        if assertion == "metric_equals" and expected is None:
            if metric not in RETENTION_FIELDS:
                raise ValueError("Only retention metrics support a null expected value")
        elif assertion == "metric_equals" and (isinstance(expected, bool) or not isinstance(expected, (int, float)) or not math.isfinite(expected) or expected < 0):
            raise ValueError("Expected value must be a finite non-negative number or null for an unavailable retention rate")
        if metric in ("gross_retention", "ttm_gross_retention") and expected is not None and expected > 1:
            raise ValueError("Gross retention is a fraction between 0 and 1")
        title, rationale = body.get("title"), body.get("rationale")
        if not isinstance(title, str) or not 1 <= len(title.strip()) <= 160:
            raise ValueError("Give this test a title of 1–160 characters")
        if not isinstance(rationale, str) or not 1 <= len(rationale.strip()) <= 4000:
            raise ValueError("Explain why this expected value is correct")
        data = (directory / "source.xlsx").read_bytes()
        if hashlib.sha256(data).hexdigest() != meta["source_sha256"]:
            raise ValueError("Saved source fingerprint changed")
        definition = {"source_sha256": meta["source_sha256"], "amount_mode": meta["amount_mode"],
                      "assertion": assertion,
                      "row_key": note["row_key"], "metric": metric, "expected": expected,
                      "absolute_tolerance": 1e-6, "relative_tolerance": 1e-6}
        identity = fingerprint(definition)[:24]
        destination = self.store / "cases" / identity
        case = {**definition, "id": identity, "title": title.strip(), "rationale": rationale.strip(),
                "created_at": timestamp(), "origin_run_id": meta["id"], "origin_note_id": note["id"],
                "original_note": note["note"], "filename": meta["filename"], "provenance": "Reviewer-entered expected value"}
        with self.lock:
            if (destination / "case.json").exists():
                return self.read(destination / "case.json")
            destination.mkdir(parents=True, exist_ok=True)
            (destination / "source.xlsx").write_bytes(data)
            case["case_sha256"] = fingerprint(case)
            self.write(destination / "case.json", case)
        return case

    def evaluate_cases(self, run_id):
        _, meta, result = self.run(run_id)
        rows = {}
        counts = Counter()
        for row in result["rows"] if result else []:
            key = tuple(row[k] for k in KEY_FIELDS)
            rows[key] = row
            counts[key] += 1
        outcomes = []
        for case in self.cases():
            outcome = {"case_id": case["id"], "title": case["title"], "case_sha256": case["case_sha256"],
                       "run_id": run_id, "expected": case["expected"], "actual": None,
                       "row_key": case["row_key"], "metric": case["metric"]}
            if case["source_sha256"] != meta["source_sha256"] or case["amount_mode"] != meta["amount_mode"]:
                outcome.update(status="skipped", reason="Different source data or amount convention")
            elif meta["status"] not in ("complete", "failed"):
                outcome.update(status="pending", reason="Run has not finished")
            elif meta["status"] == "failed":
                outcome.update(status="fail", reason="Run failed before producing a complete result")
            elif case.get("assertion") in ("row_absent", "row_present"):
                count = counts[tuple(case["row_key"])]
                expected_count = 0 if case["assertion"] == "row_absent" else 1
                outcome.update(actual=count, expected=expected_count, status="pass" if count == expected_count else "fail",
                               reason=f"Expected {expected_count} matching row(s); found {count}")
            elif counts[tuple(case["row_key"])] != 1:
                outcome.update(status="fail", reason="Expected exactly one matching output row")
            else:
                actual = rows[tuple(case["row_key"])][case["metric"]]
                expected = case["expected"]
                equal = expected is None and actual is None
                if expected is not None and actual is not None:
                    equal = math.isfinite(actual) and math.isclose(actual, expected,
                        abs_tol=case["absolute_tolerance"], rel_tol=case["relative_tolerance"])
                outcome.update(actual=actual, status="pass" if equal else "fail",
                               reason="Matches reviewer answer" if equal else "Differs from reviewer answer")
            outcomes.append(outcome)
        return outcomes

    def rerun_case(self, identity, body, create_run):
        case = next((c for c in self.cases() if c["id"] == identity), None)
        if case is None:
            raise ValueError("Regression case not found")
        model = valid_model(body.get("model"))
        if case["amount_mode"] != "period_sum":
            raise ValueError("Model regression runs currently support monthly revenue only")
        data = (self.store / "cases" / identity / "source.xlsx").read_bytes()
        if hashlib.sha256(data).hexdigest() != case["source_sha256"]:
            raise ValueError("Regression source fingerprint changed")
        # Only the source reaches the candidate; expected values and review notes stay blind.
        return create_run(data, case["filename"], mode="model", amount_mode=case["amount_mode"],
                          selected_model=model, regression_case_id=identity)

    def suite_report(self):
        cases = self.cases()
        evaluations = []
        for file in self.store.glob("*/run.json"):
            run_id = file.parent.name
            _, meta, _ = self.run(run_id)
            for result in self.evaluate_cases(run_id):
                if result["status"] != "skipped":
                    evaluations.append({**result, "mode": meta["mode"], "model": meta.get("selected_model", meta.get("model")),
                                        "created_at": meta["created_at"]})
        return {"cases": cases, "evaluations": sorted(evaluations, key=lambda x: x["created_at"], reverse=True),
                "caveat": "Expected values are reviewer-authored. Passing these specific assertions is not a general model benchmark."}
