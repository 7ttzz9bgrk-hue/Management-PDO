import os
from typing import Iterable


def normalize_path(path: str) -> str:
    """Return a normalized real path for secure path comparison."""
    return os.path.realpath(os.path.abspath(path))


def is_allowed_path(candidate_path: str, allowed_paths: Iterable[str]) -> bool:
    """Check whether candidate_path resolves to one of the configured allowed paths."""
    normalized_candidate = normalize_path(candidate_path)
    normalized_allowed = {normalize_path(p) for p in allowed_paths}
    return normalized_candidate in normalized_allowed
