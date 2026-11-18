"""Resolve run IDs to run directories, in benchmark folders or directly under the runs root."""
import re
from pathlib import Path


def run_directory(root, run_id, lane=None):
    root = Path(root)
    if not re.fullmatch(r"[a-zA-Z0-9][a-zA-Z0-9._-]*", str(run_id)) or run_id in ("atomic", "e2e"):
        raise ValueError("Unsafe run ID")
    if lane is not None and lane not in ("atomic", "e2e"):
        raise ValueError("Unknown benchmark lane")
    matches = [p for p in (root / run_id, root / "atomic" / run_id, root / "e2e" / run_id) if p.is_dir()]
    if len(matches) > 1:
        raise ValueError("Duplicate run ID in registry: " + run_id)
    if matches:
        return matches[0]
    return root / lane / run_id if lane else root / run_id


def attempt_metadata(root):
    root = Path(root)
    for pattern in ("*/tasks/*/metadata.json", "atomic/*/tasks/*/metadata.json", "e2e/*/tasks/*/metadata.json"):
        yield from root.glob(pattern)
