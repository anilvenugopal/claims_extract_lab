import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import local_harness  # noqa: E402  (installs stubs, imports src/app.py unchanged)
from make_samples import build_all  # noqa: E402


@pytest.fixture(scope="session")
def h():
    return local_harness


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    """{filename: Path} of freshly generated sample documents."""
    return build_all(tmp_path_factory.mktemp("samples"))
