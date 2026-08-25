"""Resumable-batch-crawl helper, shared by every country script.

Each country's register has thousands of entries and a single real-website
crawl per record can take several seconds to tens of seconds - a 200-record
run is a long-running process that may get interrupted (network blip,
machine sleep, Ctrl-C). Rather than re-crawling everything from scratch on
every run, each processed record is appended to a JSONL checkpoint file
immediately after it's fetched, keyed by the register's own unique ID.
Re-running the script skips any ID already present in the checkpoint and
picks up exactly where it left off, and the final .xlsx is always rebuilt
fresh from the full checkpoint contents (not just the current run's rows).
"""
import json
import os


_RESULT_FIELD_CANDIDATES = ("Best Email", "Best email", "Email")  # England/
# Slovakia, Norway, Turkey respectively - whichever one a given row has.


def _has_result(rec):
    return any(rec.get(f) for f in _RESULT_FIELD_CANDIDATES)


def load_checkpoint(path):
    """Returns {unique_id: row_dict} for every record already processed.
    When the same ID appears more than once (a resumed/retried run appends
    rather than overwrites - see append_checkpoint below), the later entry
    normally wins so a successful retry replaces an earlier failure. But if
    the later entry is blank while an earlier one for the same ID already
    found a result, the earlier (data-bearing) one is kept instead - a
    blank duplicate should never silently erase a real answer that was
    already found."""
    rows = {}
    if not os.path.exists(path):
        return rows
    is_turkey = "turkey" in os.path.basename(path).lower()
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            key = rec["Website"] if is_turkey and "Website" in rec else rec["_id"]
            existing = rows.get(key)
            if existing is not None and _has_result(existing) and not _has_result(rec):
                continue
            rows[key] = rec
    return rows


def append_checkpoint(path, unique_id, row):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    rec = dict(row)
    rec["_id"] = unique_id
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def export_xlsx_from_checkpoint(path, fieldnames, out_path, sheet_name, text_columns=None):
    from output_utils import write_xlsx

    rows = load_checkpoint(path)
    out_rows = [{k: v for k, v in r.items() if k != "_id"} for r in rows.values()]
    write_xlsx(out_rows, fieldnames, out_path, sheet_name=sheet_name, text_columns=text_columns)
    return len(out_rows)
