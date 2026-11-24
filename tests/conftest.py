"""Shared fixtures for the DOCX parser test suite."""

from __future__ import annotations

import json
import pathlib
import sys

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

SAMPLES = ROOT / "samples"
EXPECTED = SAMPLES / "expected"
EDGE = SAMPLES / "edge"


@pytest.fixture(scope="session")
def root() -> pathlib.Path:
    return ROOT


@pytest.fixture(scope="session")
def samples() -> pathlib.Path:
    return SAMPLES


@pytest.fixture(scope="session")
def edge() -> pathlib.Path:
    return EDGE


def load_snapshot(name: str):
    """Load samples/expected/<name>.json as a list of block dicts."""
    path = EXPECTED / (name + ".json")
    return json.loads(path.read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def snapshot():
    return load_snapshot
